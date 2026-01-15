import type {
  ISharePointClient,
  ListItemQueryOptions,
  SearchRequestOptions,
  SearchResult,
  SharePointConfig,
  FieldDefinition,
  IBatch,
  UserInfo,
  SPBasePermissions,
  SiteGroup,
  FileVersion,
  WebInfo,
  ListInfo,
  AttachmentInfo,
} from '../types'
import { getServerRelativePath } from '../utils/urlUtils'
import { adaptFileMetadata } from '../utils/metadataAdapter'
import { Logger } from '../utils/debug'

// Internal Type
interface InternalBatchItem {
  url: string
  method: string
  body?: any
  headers?: any
}

// --- Internal Cache Helper ---
class InternalCache {
  private enabled: boolean
  private prefix = 'sp-rest-cache::'

  constructor(enabled: boolean) {
    this.enabled = enabled
  }

  get<T>(key: string): T | null {
    if (!this.enabled) return null
    const item = sessionStorage.getItem(this.prefix + key)
    if (!item) return null
    try {
      const parsed = JSON.parse(item)
      // Simple expiry check (20 minutes default)
      if (Date.now() > parsed.expiry) {
        sessionStorage.removeItem(this.prefix + key)
        return null
      }
      return parsed.data
    } catch {
      return null
    }
  }

  set(key: string, data: any, ttlMinutes = 20) {
    if (!this.enabled) return
    const record = {
      data,
      expiry: Date.now() + ttlMinutes * 60 * 1000,
    }
    sessionStorage.setItem(this.prefix + key, JSON.stringify(record))
  }
}

export class RestSharePointClient implements ISharePointClient {
  private baseUrl: string
  private authProvider?: () => Promise<Record<string, string>>
  private cache: InternalCache
  private logger: Logger

  // Digest Caching (Memory only)
  private digestCache: string | null = null
  private digestExpiry: number = 0

  constructor(options: SharePointConfig) {
    let url = options.baseUrl
    if (
      options.devBaseUrl &&
      typeof location !== 'undefined' &&
      (location.hostname === 'localhost' || location.hostname === '127.0.0.1')
    ) {
      url = options.devBaseUrl
    }

    this.baseUrl = url.replace(/\/$/, '')
    this.authProvider = options.authProvider
    this.cache = new InternalCache(options.enableCache ?? true)
    this.logger = new Logger(options.debug)
  }

  // --- Centralized Request Handler ---
  private async request<T = any>(
    endpoint: string,
    options: {
      method?: string
      body?: any
      headers?: Record<string, string>
      isWrite?: boolean
      skipMetadata?: boolean // If true, doesn't unwrap .d or .d.results
      targetPath?: string // Optional: The file/folder path to determine the correct API root
      abortSignal?: AbortSignal
    } = {}
  ): Promise<T> {
    const { method = 'GET', body, isWrite = false, targetPath, abortSignal } = options

    let finalEndpoint = endpoint;
    let contextUrl: string | undefined = undefined;

    // 1. Determine API Root based on targetPath
    if (targetPath) {
        const apiRoot = this.getApiRoot(targetPath);
        contextUrl = apiRoot;

        // If endpoint is relative, prepend the calculated apiRoot
        if (!endpoint.startsWith('http')) {
            finalEndpoint = `${apiRoot}${endpoint}`;
        } else {
            finalEndpoint = endpoint;
        }
    } else {
        // Default behavior
        if (endpoint.startsWith('http')) {
            finalEndpoint = endpoint;
        } else {
            finalEndpoint = `${this.baseUrl}${endpoint}`;
        }
    }

    // 2. Determine Digest Context (if not already set by targetPath logic)
    if (isWrite && !contextUrl) {
       // Try to infer from endpoint if it's absolute
       if (finalEndpoint.startsWith('http')) {
           try {
             const urlObj = new URL(finalEndpoint);
             contextUrl = this.getApiRoot(urlObj.pathname);
           } catch { /* ignore */ }
       }
    }

    const headers = await this.getHeaders(isWrite, contextUrl)
    if (options.headers) {
      Object.entries(options.headers).forEach(([k, v]) => headers.set(k, v))
    }

    const fetchOptions: RequestInit = {
      method,
      headers,
      signal: abortSignal
    }

    if (body) {
      // Check for Blob or ArrayBuffer to avoid JSON stringification
      const isBinary =
        body instanceof Blob ||
        body instanceof ArrayBuffer ||
        (typeof Buffer !== 'undefined' && Buffer.isBuffer(body))

      fetchOptions.body = (
        isBinary || typeof body === 'string' ? body : JSON.stringify(body)
      ) as BodyInit
    }

    this.logger.log(`Request: ${method} ${finalEndpoint}`, body ? { body } : '')

    const response = await fetch(finalEndpoint, fetchOptions)

    if (!response.ok) {
      const errorText = await response.text()
      this.logger.error(
        `Request Failed: ${response.status} ${response.statusText}`,
        errorText
      )
      throw new Error(
        `SharePoint Request Failed: ${method} ${endpoint} - ${response.status} ${response.statusText}\n${errorText}`
      )
    }

    // Handle empty responses (e.g. 204 No Content)
    if (response.status === 204) {
      return null as T
    }

    const data = await response.json()
    this.logger.log(`Response: ${response.status}`, data)

    if (options.skipMetadata) {
      return data as T
    }

    // Unwrap SharePoint OData verbose response
    if (data && data.d) {
      // If it's a collection, return results
      if (data.d.results) {
        return data.d.results as T
      }
      return data.d as T
    }

    return data as T
  }

  // --- Batching Implementation ---
  async executeBatch(builder: (batch: IBatch) => void, abortSignal?: AbortSignal): Promise<void> {
    const queue: InternalBatchItem[] = []

    const proxy: IBatch = {
      createListItem: (list, payload) => {
        queue.push({
          method: 'POST',
          url: `/_api/web/lists/getbytitle('${list}')/items`,
          body: payload,
        })
      },
      updateListItem: (list, id, payload) => {
        queue.push({
          method: 'MERGE',
          url: `/_api/web/lists/getbytitle('${list}')/items(${id})`,
          body: payload,
          headers: { 'IF-MATCH': '*' },
        })
      },
      deleteListItem: (list, id) => {
        queue.push({
          method: 'POST',
          url: `/_api/web/lists/getbytitle('${list}')/items(${id})/recycle()`,
        })
      },
      deleteFile: (url) => {
        queue.push({
          method: 'POST',
          url: `/_api/web/getfilebyserverrelativeurl('${url}')/recycle()`,
        })
      },
    }

    builder(proxy) // Fill Queue
    if (queue.length === 0) return
    this.logger.log(`Executing Batch with ${queue.length} items`)
    await this.processInternalBatch(queue, abortSignal)
  }

  private async processInternalBatch(requests: InternalBatchItem[], abortSignal?: AbortSignal) {
    const batchGuid = 'batch_' + this.generateUuid()
    const changesetGuid = 'changeset_' + this.generateUuid()
    const digest = await this.getRequestDigest()

    let body = `--${batchGuid}\r\nContent-Type: multipart/mixed; boundary=${changesetGuid}\r\n\r\n`

    for (const req of requests) {
      body += `--${changesetGuid}\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\n`
      const fullUrl = `${this.baseUrl}${req.url}`
      body += `${req.method} ${fullUrl} HTTP/1.1\r\nAccept: application/json;odata=verbose\r\n`
      if (req.headers)
        Object.entries(req.headers).forEach(
          ([k, v]) => (body += `${k}: ${v}\r\n`)
        )
      if (req.body) {
        body += `Content-Type: application/json;odata=verbose\r\n\r\n${JSON.stringify(
          req.body
        )}\r\n`
      } else {
        body += `\r\n`
      }
    }
    body += `--${changesetGuid}--\r\n--${batchGuid}--\r\n`

    // Note: Batch request uses fetch directly because of custom body structure
    const headers = new Headers({
      'X-RequestDigest': digest,
      'Content-Type': `multipart/mixed; boundary=${batchGuid}`,
    })
    if (this.authProvider) Object.assign(headers, await this.authProvider())

    const res = await fetch(`${this.baseUrl}/_api/$batch`, {
      method: 'POST',
      headers,
      body,
      signal: abortSignal
    })
    if (!res.ok) throw new Error(`Batch failed: ${await res.text()}`)
  }

  // --- Search ---
  async search<T = any>(opts: SearchRequestOptions, abortSignal?: AbortSignal): Promise<SearchResult<T>> {
    const kql = this.buildKql(opts)
    this.logger.log(`Search KQL: ${kql}`)

    const userSelects = opts.selectFields || []
    const userExpands = opts.expandFields || []
    const needsHydration =
      userExpands.length > 0 || userSelects.some((f) => f.includes('/'))

    let searchSelect = [
      'Title',
      'Path',
      'OriginalPath',
      'UniqueId',
      'HitHighlightedSummary',
      'DefaultEncodingURL', // <--- Mapped to EncodedAbsUrl (The most reliable direct link)
      'PictureURL', // <--- Specifically for images
      'PictureThumbnailURL', // <--- Fallback for images
      ...userSelects.filter((f) => !f.includes('/')),
    ]

    if (opts.mapping) {
      searchSelect = [
        ...searchSelect,
        ...Object.keys(opts.mapping).filter((k) => !k.includes('/')),
      ]
    }

    if (needsHydration) {
      if (!searchSelect.includes('ListId')) searchSelect.push('ListId')
      if (!searchSelect.includes('ListItemId')) searchSelect.push('ListItemId')
    }

    // Dedupe
    searchSelect = [...new Set(searchSelect)]

    const payload = {
      request: {
        Querytext: kql,
        RowLimit: opts.rowLimit || 10,
        StartRow: opts.startRow || 0,
        SelectProperties: { results: searchSelect },
        Refiners:
          opts.refiners && opts.refiners.length > 0
            ? opts.refiners.join(',')
            : undefined,
        HitHighlightedProperties: { results: ['Contents', 'Title'] },
        SummaryLength: 250,
        EnableStemming: true,
        TrimDuplicates: false,
        SortList: opts.sortList
          ? {
              results: opts.sortList.map((s) => ({
                Property: s.property,
                Direction: s.direction === 'ascending' ? 0 : 1,
              })),
            }
          : undefined,
      },
    }

    // Search returns complex object, using skipMetadata to handle manual parsing
    const data = await this.request<any>(`/_api/search/postquery`, {
      method: 'POST',
      body: payload,
      isWrite: true,
      skipMetadata: true,
      abortSignal
    })

    const refinerData = data.d.postquery.PrimaryQueryResult.RefinementResults

    let refiners = undefined;
    if (refinerData) {
      refiners = refinerData.Refiners.results.map((r: any) => ({
        filterName: r.Name, // e.g., "RefinableString00"
        options: r.Entries.results.map((e: any) => ({
          label: e.RefinementName, // e.g., "Asia"
          count: e.RefinementCount, // e.g., 42
          token: e.RefinementToken, // Used for the actual filtering logic
        })),
      }))
    }

    const rows =
      data.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results

    let items = rows.map((r: any) => {
      const map: any = {}
      r.Cells.results.forEach((c: any) => (map[c.Key] = c.Value))

      // We prioritize DefaultEncodingURL as it is the most standard direct link
      let directUrl = [
        map.DefaultEncodingURL,
        map.PictureURL,
        map.OriginalPath,
        map.Path,
      ].find((url) => url && !url.toLowerCase().includes('dispform.aspx'))

      // If we have a filename in 'Title' and a 'Path', we can often swap out the Form part
      if (!directUrl || directUrl.toLowerCase().includes('dispform.aspx')) {
        const base = map.OriginalPath || map.Path || ''
        if (base.includes('/Forms/DispForm.aspx')) {
          const libraryPath = base.split('/Forms/')[0]
          const fileName = map.Title || ''
          if (fileName.includes('.')) {
            directUrl = `${libraryPath}/${fileName}`
          }
        }
      }

      map.DirectLink = directUrl || map.Path

      if (opts.includeRelativePath && map.DirectLink) {
        try {
          map.relativePath = decodeURIComponent(
            new URL(map.DirectLink).pathname
          )
        } catch {
          map.relativePath = map.DirectLink
        }
      }
      return map
    })

    // Hydration
    if (needsHydration && items.length > 0) {
      const listSelect = userSelects.filter((f) => !this.isSearchOnlyProp(f))

      await Promise.all(
        items.map(async (item: any) => {
          if (item.ListId && item.ListItemId) {
            try {
              const params: string[] = []
              if (listSelect.length > 0)
                params.push(`$select=${listSelect.join(',')}`)
              if (userExpands.length > 0)
                params.push(`$expand=${userExpands.join(',')}`)
              const qs = params.length > 0 ? `?${params.join('&')}` : ''

              const hydrated = await this.request(
                `/_api/web/lists/getById('${item.ListId}')/items(${item.ListItemId})${qs}`,
                { abortSignal }
              )
              if (hydrated) {
                Object.assign(item, hydrated)
              }
            } catch {
              /* ignore */
            }
          }
        })
      )
    }

    if (opts.mapping) {
      items = items.map((map: any) => {
        const out: any = {}
        const mapping = opts.mapping!
        Object.entries(mapping).forEach(([k, v]) => {
          // 1. Get Value from Source
          let val: any = map
          if (k.includes('.')) {
            const parts = k.split('.')
            for (const p of parts) {
                if (val && typeof val === 'object') {
                    // @ts-ignore
                    val = val[p]
                } else {
                    val = null
                }
            }
          } else {
            // @ts-ignore
            val = map[k]
          }

          // 2. Assign Value to Destination
          if (v.includes('.')) {
            const parts = v.split('.')
            let current = out
            for (let i = 0; i < parts.length - 1; i++) {
              const part = parts[i]
              // @ts-ignore
              if (!current[part]) current[part] = {}
              // @ts-ignore
              current = current[part]
            }
            // @ts-ignore
            current[parts[parts.length - 1]] = val
          } else {
            out[v] = val
          }
        })
        // @ts-ignore
        if (!out.url) out.url = map.Path
        // @ts-ignore
        if (opts.includeRelativePath) out.relativePath = map.relativePath
        return out
      })
    }

    console.log('refiners', refiners)
    return {
      items,
      totalHits: data.d.postquery.PrimaryQueryResult.RelevantResults.TotalRows,
      startRow: opts.startRow || 0,
      refiners: refiners,
    }
  }

  // --- Cached Metadata Methods ---

  async getCurrentUser(abortSignal?: AbortSignal) {
    const cacheKey = 'CurrentUser'
    const cached = this.cache.get<any>(cacheKey)
    if (cached) return cached

    const data = await this.request<any>(`/_api/web/currentuser`, { abortSignal })
    this.cache.set(cacheKey, data)
    return data
  }

  async getListFields(list: string, webUrl?: string, abortSignal?: AbortSignal) {
    const cacheKey = `Fields:${webUrl || 'current'}:${list}`
    const cached = this.cache.get<FieldDefinition[]>(cacheKey)
    if (cached) return cached

    // If webUrl is provided, use it as the base for the request context
    // This supports cross-site schema fetching
    const endpoint = `/_api/web/lists/getbytitle('${list}')/fields?$filter=Hidden eq false`;

    // Pass webUrl as targetPath if it looks like a URL, or handle it in request logic?
    // request() uses targetPath to determine apiRoot.
    // If webUrl is the API root (e.g. /sites/Other), we pass it as targetPath.

    const requestOptions: any = { abortSignal };
    if (webUrl) {
        requestOptions.targetPath = webUrl;
    }

    const data = await this.request<any[]>(endpoint, requestOptions)

    const mapped = data.map((f: any) => ({
      InternalName: f.InternalName,
      Title: f.Title,
      TypeAsString: f.TypeAsString,
      Hidden: f.Hidden,
      Choices: f.Choices?.results || [],
      TermSetId: f.TermSetId || undefined
    }))

    this.cache.set(cacheKey, mapped)
    return mapped
  }

  async getFieldChoices(list: string, field: string, abortSignal?: AbortSignal) {
    const cacheKey = `Choices:${list}:${field}`
    const cached = this.cache.get<string[]>(cacheKey)
    if (cached) return cached

    const data = await this.request<any>(
      `/_api/web/lists/getbytitle('${list}')/fields/getByInternalNameOrTitle('${field}')`,
      { abortSignal }
    )

    const choices = data.Choices?.results || []
    this.cache.set(cacheKey, choices)
    return choices
  }

  async searchTerm(
    termSetId: string,
    label: string,
    abortSignal?: AbortSignal
  ): Promise<{ Label: string; TermGuid: string } | null> {
    // Uses the modern Taxonomy API (v2.1)
    // $filter=labels/any(l:l/name eq 'Label') or defaults to name matching
    try {
      // NOTE: We wrap the label in quotes. If label contains quotes, they need escaping.
      const safeLabel = label.replace(/'/g, "''")

      const endpoint = `/_api/v2.1/termStore/termSets/${termSetId}/terms?$filter=labels/any(l:l/name eq '${safeLabel}') or name eq '${safeLabel}'&$select=id,name,labels`

      const response = await this.request<{ value: any[] }>(endpoint, { abortSignal })

      const terms = response.value
      if (terms && terms.length > 0) {
        // Return first match
        const t = terms[0]
        // Prefer the localized name matching the input or just the default name
        return {
            Label: t.names?.[0]?.name || t.name || label,
            TermGuid: t.id
        }
      }
    } catch (e) {
      this.logger.warn(`SearchTerm failed for ${label} in ${termSetId}`, e)
    }
    return null
  }

  // --- Other Methods (Standard) ---
  private buildKql(opts: SearchRequestOptions): string {
    const parts: string[] = []
    const txt = opts.query?.trim() || '*'

    if (txt !== '*') {
      const escapedTxt = txt.replace(/"/g, '""')
      if (opts.searchTitleOnly) {
        // Logic for "Name Only" search
        parts.push(`(Title:"${escapedTxt}*" OR Filename:"${escapedTxt}*")`)
      } else {
        // Logic for "Everything" search (Title, Name, and Content)
        parts.push(
          `(Title:"${escapedTxt}*" OR Filename:"${escapedTxt}*" OR ${escapedTxt}*)`
        )
      }
    } else {
      parts.push('*')
    }

    if (opts.fileTypes?.exclude?.length) {
      opts.fileTypes.exclude.forEach((ext) => {
        if (ext.toLowerCase() === 'aspx') {
          // INSTEAD OF EXCLUDING .aspx EXTENSION:
          // Exclude Site Pages and Wiki Pages specifically by ContentClass
          // This keeps Images (which are usually STS_ListItem_Picture) visible.
          parts.push('-contentclass:STS_ListItem_850') // Site Pages
          parts.push('-contentclass:STS_ListItem_WebPageLibrary') // Wiki/Legacy Pages
        } else {
          // Use the minus operator for standard file types
          parts.push(`-FileType:${ext}`)
        }
      })
    }
    if (opts.fileTypes?.include?.length) {
      parts.push(
        `(${opts.fileTypes.include.map((e) => `FileType:${e}`).join(' OR ')})`
      )
    }

    if (opts.scope) {
      const scopes = Array.isArray(opts.scope) ? opts.scope : [opts.scope]
      let origin = ''
      try {
        origin = new URL(this.baseUrl).origin
      } catch {
        /* ignore */
      }
      const normalizedScopes = scopes.map((s) => {
        const safeS = String(s || '')
        if (safeS.toLowerCase().startsWith('http')) return `Path:"${safeS}*"`
        if (safeS.startsWith('/')) {
          // Server Relative
          return `Path:"${origin}${safeS}*"`
        }
        // Site Relative
        return `Path:"${this.baseUrl}/${safeS}*"`
      })
      parts.push(`(${normalizedScopes.join(' OR ')})`)
    }

    if (opts.resultType === 'items') {
      // 1. Exclude Folders (0x0120)
      parts.push('-ContentTypeId:0x0120*')

      // 2. Exclude the Library/List containers themselves
      // STS_List_* catches STS_List_DocumentLibrary, STS_List_Links, etc.
      parts.push('-contentclass:STS_List_*')

      // 3. (Optional but recommended) Exclude the Site and Web objects
      parts.push('-contentclass:STS_Web')
      parts.push('-contentclass:STS_Site')
    } else if (opts.resultType === 'folders') {
      // Only Folders
      parts.push('ContentTypeId:0x0120*')
    }

    if (opts.filters) {
      for (const [key, value] of Object.entries(opts.filters)) {
        if (!value || (Array.isArray(value) && value.length === 0)) continue
        let mp = key
        if (opts.mapping) {
          const mapping = opts.mapping
          const found = Object.keys(mapping).find(
            (k) => mapping[k] === key
          )
          if (found) mp = found
        }
        if (Array.isArray(value))
          parts.push(`(${value.map((v) => `${mp}:"${v}"`).join(' OR ')})`)
        else parts.push(`${mp}:"${value}"`)
      }
    }
    return parts.join(' AND ')
  }

  async createListItem(list: string, payload: any, abortSignal?: AbortSignal) {
    return await this.request(`/_api/web/lists/getbytitle('${list}')/items`, {
      method: 'POST',
      body: payload,
      isWrite: true,
      abortSignal
    })
  }

  async updateListItem(list: string, id: number, payload: any, abortSignal?: AbortSignal) {
    await this.request(`/_api/web/lists/getbytitle('${list}')/items(${id})`, {
      method: 'POST',
      body: payload,
      headers: {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      abortSignal
    })
  }

  async deleteListItem(list: string, id: number, abortSignal?: AbortSignal) {
    await this.request(`/_api/web/lists/getbytitle('${list}')/items(${id})`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      abortSignal
    })
  }

  async getListItemById(
    list: string,
    id: number,
    select?: string[],
    expand?: string[],
    abortSignal?: AbortSignal
  ) {
    const params: string[] = []
    if (select && select.length > 0) params.push(`$select=${select.join(',')}`)
    if (expand && expand.length > 0) params.push(`$expand=${expand.join(',')}`)

    const q = params.length > 0 ? `?${params.join('&')}` : ''
    return await this.request(
      `/_api/web/lists/getbytitle('${list}')/items(${id})${q}`,
      { abortSignal }
    )
  }

  async getListItems<T = any>(
    listTitle: string,
    options?: ListItemQueryOptions,
    abortSignal?: AbortSignal
  ): Promise<T[]> {
    const params: string[] = []

    if (options?.select && options.select.length > 0) {
      params.push(`$select=${options.select.join(',')}`)
    }
    if (options?.expand && options.expand.length > 0) {
      params.push(`$expand=${options.expand.join(',')}`)
    }
    if (options?.filter) {
      params.push(`$filter=${encodeURIComponent(options.filter)}`)
    }
    if (options?.top) {
      params.push(`$top=${options.top}`)
    }
    if (options?.orderBy) {
      const dir = options.ascending === false ? 'desc' : 'asc'
      params.push(`$orderby=${options.orderBy} ${dir}`)
    }

    const q = params.length > 0 ? `?${params.join('&')}` : ''
    return await this.request(
      `/_api/web/lists/getbytitle('${listTitle}')/items${q}`,
      { abortSignal }
    )
  }

  async getItemAttachments(
    list: string,
    id: number,
    abortSignal?: AbortSignal
  ): Promise<AttachmentInfo[]> {
    const results = await this.request<any[]>(
      `/_api/web/lists/getbytitle('${list}')/items(${id})/AttachmentFiles`,
      { abortSignal }
    )
    return results.map((a) => ({
      FileName: a.FileName,
      ServerRelativeUrl: a.ServerRelativeUrl,
    }))
  }

  async addAttachment(
    list: string,
    id: number,
    fileName: string,
    file: Blob | ArrayBuffer,
    abortSignal?: AbortSignal
  ): Promise<void> {
    // For binary upload, we need to ensure Content-Type is not application/json
    await this.request(
      `/_api/web/lists/getbytitle('${list}')/items(${id})/AttachmentFiles/add(FileName='${fileName}')`,
      {
        method: 'POST',
        body: file,
        isWrite: true,
        headers: {
          // Overwriting Content-Type to undefined or null isn't standard in Headers object interaction
          // but we can set it to application/octet-stream or rely on fetch logic if we handled headers map better.
          // In request() helper, we merge headers. If we pass a header that conflicts, we need to ensure it wins.
          'Content-Type': 'application/octet-stream',
        },
        abortSignal
      }
    )
  }

  async deleteAttachment(
    list: string,
    id: number,
    fileName: string,
    abortSignal?: AbortSignal
  ): Promise<void> {
    await this.request(
      `/_api/web/lists/getbytitle('${list}')/items(${id})/AttachmentFiles/getByFileName('${fileName}')`,
      {
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*',
        },
        isWrite: true,
        abortSignal
      }
    )
  }

  async uploadFile(url: string, name: string, file: any, abortSignal?: AbortSignal) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)

    const data = await this.request<any>(
      `/_api/web/getfolderbyserverrelativeurl('${fullUrl}')/files/add(url='${name}', overwrite=true)`,
      {
        method: 'POST',
        body: file,
        headers: {
          'Content-Type': 'application/octet-stream',
        },
        isWrite: true,
        targetPath: fullUrl, // Triggers central getApiRoot
        abortSignal
      }
    )
    return data.ServerRelativeUrl
  }

  async downloadFile(url: string, abortSignal?: AbortSignal) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    const apiRoot = this.getApiRoot(fullUrl)

    const headers = await this.getHeaders(false)
    const res = await fetch(
      `${apiRoot}/_api/web/getfilebyserverrelativeurl('${fullUrl}')/$value`,
      { headers, signal: abortSignal }
    )
    return await res.blob()
  }

  async updateFileMetadata(url: string, payload: any, abortSignal?: AbortSignal) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)

    // 1. Get List Item details
    const meta = await this.request<any>(
      `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/ListItemAllFields?$expand=ParentList`,
      { targetPath: fullUrl, abortSignal } // Triggers central getApiRoot
    )

    if (!meta || !meta['ParentList']) {
      throw new Error(`Could not determine List for file: ${url}`)
    }

    const listTitle = meta.ParentList.Title
    const itemId = meta.Id

    // 2. Adapt Payload
    // Pass the apiRoot (calculated from fullUrl) to ensure we fetch fields from the correct site
    const apiRoot = this.getApiRoot(fullUrl)
    const formValues = await adaptFileMetadata(this, listTitle, payload, apiRoot)

    // 3. Execute ValidateUpdateListItem on the Item
    // Note: We need to use the same apiRoot for the list call.
    // Since adaptFileMetadata uses the passed apiRoot for schema fetching, we are good.

    const response = await this.request<any>(
      `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/ValidateUpdateListItem`,
      {
        method: 'POST',
        body: {
          formValues,
          bNewDocumentUpdate: false
        },
        isWrite: true,
        targetPath: fullUrl, // Use file path to determine site again
        abortSignal
      }
    )

    // Check for field-level errors in ValidateUpdateListItem response
    // Response structure: { value: [{ FieldName, ErrorCode, ErrorMessage }] }
    // Or { d: { ValidateUpdateListItem: { results: [...] } } }

    // Our request helper handles unwrapping 'd', but ValidateUpdateListItem action usually returns nested object in 'd'.
    // e.g. d.ValidateUpdateListItem.results

    // If request helper unwraps 'd', we might get { ValidateUpdateListItem: { results: [] } }
    // Or if it unwraps 'd.results' (for collections), but this is a single action returning complex type.

    let results: any[] = []
    if (Array.isArray(response)) {
        results = response;
    } else if (response && response.ValidateUpdateListItem && response.ValidateUpdateListItem.results) {
        results = response.ValidateUpdateListItem.results;
    } else if (response && response.value) {
        results = response.value;
    }

    const errors = results.filter((r: any) => r.ErrorCode !== 0)
    if (errors.length > 0) {
        const msg = errors.map((e: any) => `${e.FieldName}: ${e.ErrorMessage}`).join('; ')
        throw new Error(`UpdateFileMetadata failed: ${msg}`)
    }
  }

  async deleteFile(url: string, abortSignal?: AbortSignal) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    await this.request(`/_api/web/getfilebyserverrelativeurl('${fullUrl}')`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      targetPath: fullUrl,
      abortSignal
    })
  }

  async createFolder(url: string, abortSignal?: AbortSignal) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    await this.request(`/_api/web/folders`, {
      method: 'POST',
      body: {
        __metadata: { type: 'SP.Folder' },
        ServerRelativeUrl: fullUrl,
      },
      isWrite: true,
      targetPath: fullUrl,
      abortSignal
    })
  }

  // --- Webs & Lists ---
  async getWebInfo(abortSignal?: AbortSignal): Promise<WebInfo> {
    const w = await this.request<any>(`/_api/web`, { abortSignal })
    return {
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }
  }

  async getSubwebs(abortSignal?: AbortSignal): Promise<WebInfo[]> {
    const webs = await this.request<any[]>(`/_api/web/webs`, { abortSignal })
    return webs.map((w) => ({
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }))
  }

  async getLists(abortSignal?: AbortSignal): Promise<ListInfo[]> {
    const lists = await this.request<any[]>(`/_api/web/lists`, { abortSignal })
    return lists.map((l) => ({
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl,
    }))
  }

  async getList(listTitle: string, abortSignal?: AbortSignal): Promise<ListInfo> {
    const l = await this.request<any>(
      `/_api/web/lists/getbytitle('${listTitle}')`,
      { abortSignal }
    )
    return {
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl,
    }
  }

  async createList(
    title: string,
    description?: string,
    template = 100,
    abortSignal?: AbortSignal
  ): Promise<ListInfo> {
    const l = await this.request<any>(`/_api/web/lists`, {
      method: 'POST',
      isWrite: true,
      body: {
        __metadata: { type: 'SP.List' },
        Title: title,
        Description: description,
        BaseTemplate: template,
      },
      abortSignal
    })
    return {
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl,
    }
  }

  async deleteList(title: string, abortSignal?: AbortSignal): Promise<void> {
    await this.request(`/_api/web/lists/getbytitle('${title}')`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      abortSignal
    })
  }

  // --- Helpers ---
  private async getHeaders(isWrite = false, contextUrl?: string): Promise<Headers> {
    const headers = new Headers({
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    })
    if (this.authProvider) {
      const a = await this.authProvider()
      Object.entries(a).forEach(([k, v]) => headers.append(k, v))
    }
    if (isWrite) {
      headers.append('X-RequestDigest', await this.getRequestDigest(contextUrl))
    }
    return headers
  }

  private async getRequestDigest(customBaseUrl?: string): Promise<string> {
    // Note: Digest is per-site (web). If we are targeting a different site, we need a different digest.
    // For simplicity, we cache based on the DEFAULT baseUrl.
    // If a customBaseUrl is provided (different from default), we bypass cache or should cache separately.
    // To support cross-site writes properly, we should really cache by URL.

    // For now, if customBaseUrl is passed and differs, we fetch fresh.
    const targetUrl = customBaseUrl || this.baseUrl;
    const isDefault = targetUrl === this.baseUrl;

    if (isDefault && this.digestCache && Date.now() < this.digestExpiry)
      return this.digestCache

    // We use a simplified fetch here to avoid recursion with `request` which might call getHeaders
    const h = new Headers({ Accept: 'application/json;odata=verbose' })
    if (this.authProvider) Object.assign(h, await this.authProvider())

    const res = await fetch(`${targetUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: h,
    })
    if (!res.ok) return ''
    const d = await res.json()

    const token = d.d.GetContextWebInformation.FormDigestValue

    if (isDefault) {
        this.digestCache = token
        this.digestExpiry =
          Date.now() +
          d.d.GetContextWebInformation.FormDigestTimeoutSeconds * 1000 -
          60000
    }
    return token
  }

  /**
   * Determines the correct SharePoint API root URL based on the file path.
   * If the file path points to a different site collection (e.g. /sites/Other),
   * this returns the URL to that site (https://tenant/sites/Other),
   * allowing cross-site uploads/queries.
   */
  private getApiRoot(urlOrPath: string): string {
    // 1. Get Origin
    let origin = ''
    try {
        origin = new URL(this.baseUrl).origin
    } catch {
        return this.baseUrl // Fallback
    }

    // 2. Parse Path
    // Normalize input (handle absolute URLs passed as targetPath)
    const serverRelativePath = getServerRelativePath(urlOrPath, this.baseUrl)

    // Heuristic: SharePoint site collections usually start with /sites/ or /teams/
    // We look for the first 2 segments: /sites/SiteName or /teams/TeamName
    // Or / (root site)

    const path = serverRelativePath.startsWith('/') ? serverRelativePath : `/${serverRelativePath}`
    const parts = path.split('/').filter(p => p) // Remove empty

    if (parts.length >= 2) {
        const category = parts[0].toLowerCase()
        if (category === 'sites' || category === 'teams') {
            return `${origin}/${parts[0]}/${parts[1]}`
        }
    }

    // Fallback: If it's the root site collection (path starts with something else usually, or just /)
    // But we might be in a subweb of the current baseUrl.
    // Without exact knowledge of topology, we usually default to this.baseUrl
    // OR return the Site Collection Root if we detect we are breaking out of the current web context?

    // User Request: "query and upload to different sharepoints based on where the relativeurl is pointing"
    // This strongly implies identifying the Site Collection.

    // If the path doesn't match /sites/ or /teams/, it might be the root site collection.
    // e.g. /Shared Documents -> https://tenant.com
    if (parts.length > 0 && !['sites', 'teams'].includes(parts[0].toLowerCase())) {
         return origin
    }

    return this.baseUrl
  }

  private generateUuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0
      return (c == 'x' ? r : (r & 0x3) | 0x8).toString(16)
    })
  }

  async getSiteUsers(abortSignal?: AbortSignal): Promise<UserInfo[]> {
    return await this.request<UserInfo[]>(
      `/_api/web/siteusers?$filter=PrincipalType eq 1`,
      { abortSignal }
    )
  }

  async searchUsers(query: string, abortSignal?: AbortSignal): Promise<UserInfo[]> {
    const pickerEndpoint = `/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

    const payload = {
      queryParams: {
        QueryString: query,
        MaximumEntitySuggestions: 15,
        AllowEmailAddresses: true,
        PrincipalSource: 15,
        PrincipalType: 1,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        EnabledClaimProviders: "",
        ForceClaims: false,
        Required: false,
        SharePointGroupID: 0,
        UrlZone: 0
      }
    };

    let finalResults: UserInfo[] = [];

    // 1. Try People Picker (Best for direct directory resolution)
    try {
      const response = await this.request<any>(pickerEndpoint, {
        method: 'POST',
        body: payload,
        isWrite: true,
        skipMetadata: true,
        abortSignal
      });

      const pickerResults = JSON.parse(response.d.ClientPeoplePickerSearchUser);
      finalResults = pickerResults.map((r: any) => ({
        Id: 0, // Picker doesn't give Site ID
        Title: r.DisplayText,
        Email: r.EntityData.Email || r.Description,
        LoginName: r.Key
      }));
    } catch (err) {
      this.logger.error('People Picker failed', err);
    }

    //fallback for user search, this is needed if the general search does not match enough results to do more of a Fuzzy search for users
    if (finalResults.length < 5 && query.length > 2) {
      try {
        // 1. Tokenize: Split "Tom Greener" -> ["Tom", "Greener"]
        const searchTerms = query.trim().split(/\s+/).filter(t => t.length > 0);

        // 2. Helper to Title Case (e.g., "tom" -> "Tom")
        // SharePoint 'siteusers' is often case-sensitive. Title Casing increases match probability.
        const toTitleCase = (str: string) =>
            str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();

        // 3. Build Broad OData Filter (Use 'OR' to gather candidates)
        // Logic: Get users who match "Tom" OR "Greener".
        // We rely on 'OR' because 'AND' in OData fails if strict case/order is slightly off.
        const filterConditions = searchTerms.map(term => {
          const cleanTerm = encodeURIComponent(toTitleCase(term.replace(/'/g, "''")));
          return `(substringof('${cleanTerm}', Title) or substringof('${cleanTerm}', Email))`;
        });

        // Join with 'or' to get a wide list of potential matches
        const filterString = filterConditions.join(' or ');

        // Fetch top 20 candidates to ensure we capture the right person even with common names
        const siteUsersUrl = `/_api/web/siteusers?$filter=${filterString}&$top=20`;

        const siteUsers = await this.request<any[]>(siteUsersUrl, { abortSignal });

        // 4. Client-Side Refinement (Case Insensitive + Order Independent)
        // Now we strictly check that the user matches ALL terms (e.g., must have "Tom" AND "Greener")
        const lowerQueryParts = searchTerms.map(t => t.toLowerCase());

        siteUsers.forEach((u: any) => {
          const titleLower = (u.Title || "").toLowerCase();
          const emailLower = (u.Email || "").toLowerCase();

          // Check: Does this user contain ALL the words from the query?
          // This allows "Tom Greener" to match "Greener, Tom"
          const isMatch = lowerQueryParts.every(part =>
              titleLower.includes(part) || emailLower.includes(part)
          );

          if (isMatch) {
            // Deduplicate
            const exists = finalResults.some(existing =>
                (existing.Email && u.Email && existing.Email.toLowerCase() === u.Email.toLowerCase()) ||
                (existing.LoginName && u.LoginName && existing.LoginName.toLowerCase() === u.LoginName.toLowerCase())
            );

            if (!exists) {
              finalResults.push({
                Id: u.Id,
                Title: u.Title,
                Email: u.Email,
                LoginName: u.LoginName
              });
            }
          }
        });

      } catch (siteErr) {
        this.logger.warn('Site User fallback search failed', siteErr);
      }
    }

    return finalResults;
  }

  async ensureUser(loginName: string, abortSignal?: AbortSignal): Promise<UserInfo> {
    const data = await this.request<any>(`/_api/web/ensureUser`, {
      method: 'POST',
      isWrite: true,
      body: { logonName: loginName },
      abortSignal
    })
    return data
  }

  async getUserGroups(email?: string, abortSignal?: AbortSignal): Promise<SiteGroup[]> {
    let endpoint = ''

    if (email) {
      // Step 1: Get the User ID/LoginName from Email
      const user = await this.getUserByEmail(email)
      endpoint = `/_api/web/siteusers/getByLoginName(@v)/groups?@v='${encodeURIComponent(
        user.LoginName
      )}'`
    } else {
      // Current User
      endpoint = `/_api/web/currentuser/groups`
    }

    return await this.request<SiteGroup[]>(endpoint, { abortSignal })
  }

  async addUserToGroup(groupName: string, loginName: string, abortSignal?: AbortSignal): Promise<void> {
    const user = await this.ensureUser(loginName)
    await this.request(`/_api/web/sitegroups/getByName('${groupName}')/users`, {
      method: 'POST',
      isWrite: true,
      body: { LoginName: user.LoginName },
      abortSignal
    })
  }

  async removeUserFromGroup(
    groupName: string,
    loginName: string,
    abortSignal?: AbortSignal
  ): Promise<void> {
    await this.request(
      `/_api/web/sitegroups/getByName('${groupName}')/users/removeByLoginName(@v)?@v='${encodeURIComponent(
        loginName
      )}'`,
      {
        method: 'POST',
        isWrite: true,
        abortSignal
      }
    )
  }

  async createGroup(
    groupName: string,
    description?: string,
    abortSignal?: AbortSignal
  ): Promise<SiteGroup> {
    return await this.request<SiteGroup>(`/_api/web/sitegroups`, {
      method: 'POST',
      isWrite: true,
      body: {
        __metadata: { type: 'SP.Group' },
        Title: groupName,
        Description: description,
      },
      abortSignal
    })
  }

  async getUserEffectivePermissions(
    email?: string,
    abortSignal?: AbortSignal
  ): Promise<SPBasePermissions> {
    let endpoint = ''

    if (email) {
      const user = await this.getUserByEmail(email)
      endpoint = `/_api/web/getUserEffectivePermissions(@v)?@v='${encodeURIComponent(
        user.LoginName
      )}'`
    } else {
      endpoint = `/_api/web/effectiveBasePermissions`
    }

    return await this.request<SPBasePermissions>(endpoint, { abortSignal })
  }

  // Helper to resolve Email -> LoginName
  private async getUserByEmail(email: string): Promise<UserInfo> {
    // We filter site users to find the specific email
    const users = await this.request<UserInfo[]>(
      `/_api/web/siteusers?$filter=Email eq '${email}'`
    )
    const user = users[0]
    if (!user)
      throw new Error(`User with email ${email} not found on this site.`)
    return user
  }

  async getFileVersions(url: string, abortSignal?: AbortSignal): Promise<FileVersion[]> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    // We expand CreatedBy to get the user name
    const endpoint = `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/versions?$expand=CreatedBy`

    const results = await this.request<any[]>(endpoint, { abortSignal })

    return results.map((v: any) => ({
      VersionLabel: v.VersionLabel,
      Created: v.Created,
      CheckInComment: v.CheckInComment,
      IsCurrentVersion: v.IsCurrentVersion,
      Size: parseInt(v.Size, 10), // REST API often returns Size as string
      Url: v.Url, // This is the relative URL to this specific version
      CreatedBy: {
        Id: v.CreatedBy.Id,
        Title: v.CreatedBy.Title,
        Email: v.CreatedBy.Email,
      },
    }))
  }

  async getVersionHistoryLink(serverRelativeUrl: string, abortSignal?: AbortSignal): Promise<string> {
    // 1. Escape single quotes for the OData query (e.g. "O'Neil.docx" -> "O''Neil.docx")
    // We assume serverRelativeUrl is decoded or we need to handle it.
    // The user provided logic: const safeUrl = fullUrl.replace(/'/g, "''")
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const safeUrl = fullUrl.replace(/'/g, "''")

    // 2. Fetch the Item ID and Parent List ID
    // We use targetPath to ensure the request goes to the correct site (e.g. /sites/ProposalExpertiseHubUAT/V2)
    const data = await this.request<any>(
      `/_api/web/getFileByServerRelativeUrl('${safeUrl}')/ListItemAllFields?$select=Id,ParentList/Id&$expand=ParentList`,
      {
        method: 'GET',
        targetPath: fullUrl, // vital for calculating the correct apiRoot
        abortSignal
      }
    )

    if (!data || !data.ParentList) {
      throw new Error(`Could not resolve List ID for file: ${fullUrl}`)
    }

    const itemId = data.Id
    const listId = data.ParentList.Id

    // 3. Use builder
    const fileSiteRoot = this.getApiRoot(fullUrl)
    return this.getVersionHistoryLinkByItem(listId, itemId, fileSiteRoot)
  }

  getVersionHistoryLinkByItem(
    listId: string,
    itemId: number,
    webUrl?: string
  ): string {
    const root = webUrl ? webUrl.replace(/\/$/, '') : this.baseUrl
    // Format: .../Versions.aspx?list={GUID}&ID=123
    // Note: The list GUID should be wrapped in curly braces
    return `${root}/_layouts/15/Versions.aspx?list={${listId}}&ID=${itemId}`
  }

  private isSearchOnlyProp(field: string): boolean {
    const searchOnlyPattern =
      /^(Refinable|HitHighlighted|Path$|OriginalPath$|Rank$|DocId$|WorkId$|Piw)/i
    return searchOnlyPattern.test(field)
  }
}
