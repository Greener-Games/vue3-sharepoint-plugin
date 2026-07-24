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
  AttachmentInfo,
  RefinerResult,
  WebInfo,
  ListInfo,
} from '../types'
import { getServerRelativePath } from '../utils/urlUtils'
import { adaptFileMetadata } from '../utils/metadataAdapter'
import { Logger } from '../utils/debug'

// Internal Type
interface InternalBatchItem {
  url: string
  method: string
  body?: unknown
  headers?: Record<string, string>
  listTitle?: string
}

/** Internal interface for cached items */
interface CacheRecord<T> {
  data: T
  expiry: number
}

// --- Internal Cache Helper ---
class InternalCache {
  private enabled: boolean
  private prefix = 'sp-rest-cache::'

  constructor(enabled: boolean) {
    this.enabled = enabled
  }

  /** Retrieves a cached item if it exists and hasn't expired */
  get<T>(key: string): T | null {
    if (!this.enabled) return null
    const item = sessionStorage.getItem(this.prefix + key)
    if (!item) return null
    try {
      const parsed = JSON.parse(item) as CacheRecord<T>
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

  /** Caches an item with a specific TTL */
  set<T>(key: string, data: T, ttlMinutes = 20): void {
    if (!this.enabled) return
    const record: CacheRecord<T> = {
      data,
      expiry: Date.now() + ttlMinutes * 60 * 1000,
    }
    sessionStorage.setItem(this.prefix + key, JSON.stringify(record))
  }
}

/** Interface for SharePoint OData Verbose responses */
interface SPVerboseResponse<T> {
  d?: T & {
    results?: T extends Array<infer U> ? U[] : never
    __next?: string
    GetContextWebInformation?: {
      FormDigestValue: string
      FormDigestTimeoutSeconds: number
    }
    postquery?: {
        PrimaryQueryResult: {
            RelevantResults: {
                Table: {
                    Rows: {
                        results: Array<{ Cells: { results: Array<{ Key: string, Value: unknown }> } }>
                    }
                }
                TotalRows: number
            }
            RefinementResults?: {
                Refiners: {
                    results: Array<{
                        Name: string
                        Entries: {
                            results: Array<{
                                RefinementName: string
                                RefinementCount: number
                                RefinementToken: string
                            }>
                        }
                    }>
                }
            }
        }
    }
    ValidateUpdateListItem?: {
        results: Array<{
            FieldName: string
            ErrorCode: number
            ErrorMessage: string
        }>
    }
  }
}

/** Interface for SharePoint OData V2.1 / REST responses */
interface SPODataResponse<T> {
  value?: T
  '@odata.nextLink'?: string
  'odata.nextLink'?: string
  ClientPeoplePickerSearchUser?: string
}

export class RestSharePointClient implements ISharePointClient {
  private baseUrl: string
  private authProvider?: () => Promise<Record<string, string>>
  private cache: InternalCache
  private logger: Logger

  // Digest Caching (Memory only)
  private digestCache: string | null = null
  private digestExpiry = 0

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
  /** Gets the configured base URL */
  public getBaseUrl(): string {
    return this.baseUrl
  }

  /** Performs a SharePoint REST API request */
  public async request<T>(
    endpoint: string,
    options: Omit<RequestInit, 'body'> & {
      isWrite?: boolean
      skipMetadata?: boolean // If true, doesn't unwrap .d or .d.results
      targetPath?: string // Optional: The file/folder path to determine the correct API root
      body?: unknown
    } = {}
  ): Promise<T> {
    const {
      method = 'GET',
      body,
      isWrite = false,
      targetPath,
      signal,
      skipMetadata: _skipMetadata,
      ...fetchProps
    } = options

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
      Object.entries(options.headers).forEach(([k, v]) => headers.set(k, String(v)))
    }

    const fetchOptions: RequestInit = {
      ...fetchProps,
      method,
      headers,
      signal: signal
    }

    if (body) {
      // Check for Blob or ArrayBuffer to avoid JSON stringification
      const isBinary =
        body instanceof Blob ||
        body instanceof ArrayBuffer

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
        [errorText]
      )
      throw new Error(
        `SharePoint Request Failed: ${method} ${endpoint} - ${response.status} ${response.statusText}\n${errorText}`
      )
    }

    // Handle empty responses (e.g. 204 No Content)
    if (response.status === 204) {
      return null as unknown as T
    }

    const data = (await response.json()) as SPVerboseResponse<T> & SPODataResponse<T>
    this.logger.log(`Response: ${response.status}`, [data])

    if (options.skipMetadata) {
      return data as unknown as T
    }

    // Unwrap SharePoint OData verbose response
    if (data && data.d) {
        const d = data.d;
        if (d && typeof d === 'object' && 'results' in d && d.results) {
            return d.results as unknown as T;
        }
        return d;
    }

    return data as unknown as T
  }

  /** Gets information about the current web */
  async getWebInfo(signal?: AbortSignal): Promise<WebInfo> {
    const w = await this.request<WebInfo>(`/_api/web?$select=Id,Title,Url,Description`, { signal: signal })
    return {
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }
  }

  // --- Batching Implementation ---
  /** Executes multiple SharePoint operations in a single batch */
  async executeBatch(builder: (batch: IBatch) => void, signal?: AbortSignal): Promise<void> {
    const queue: InternalBatchItem[] = []

    const proxy: IBatch = {
      createListItem: (list, payload) => {
        queue.push({
          method: 'POST',
          url: `/_api/web/lists/getbytitle('${list}')/items`,
          body: payload,
          listTitle: list
        })
      },
      updateListItem: (list, id, payload) => {
        queue.push({
          method: 'MERGE',
          url: `/_api/web/lists/getbytitle('${list}')/items(${id})`,
          body: payload,
          headers: { 'IF-MATCH': '*' },
          listTitle: list
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
    await this.processInternalBatch(queue, signal)
  }

  private async processInternalBatch(requests: InternalBatchItem[], signal?: AbortSignal) {
    const batchGuid = 'batch_' + this.generateUuid()
    const changesetGuid = 'changeset_' + this.generateUuid()
    const digest = await this.getRequestDigest()

    // Pre-process items that need __metadata
    for (const req of requests) {
      if (req.listTitle && req.body && typeof req.body === 'object' && !('__metadata' in req.body)) {
        const type = await this.getListEntityType(req.listTitle)
        req.body = {
          __metadata: { type },
          ...req.body
        }
      }
    }

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
    if (this.authProvider) {
        const authHeaders = await this.authProvider()
        Object.entries(authHeaders).forEach(([k, v]) => headers.set(k, v))
    }

    const res = await fetch(`${this.baseUrl}/_api/$batch`, {
      method: 'POST',
      headers,
      body,
      signal: signal
    })
    if (!res.ok) throw new Error(`Batch failed: ${await res.text()}`)
  }

  // --- Search ---
  /** Executes a search using SharePoint Search API */
  async search<T>(opts: SearchRequestOptions, signal?: AbortSignal): Promise<SearchResult<T>> {
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
        SortList: opts.sortList && opts.sortList.length > 0
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
    const response = await this.request<SPVerboseResponse<Record<string, unknown>>>(`/_api/search/postquery`, {
      method: 'POST',
      body: payload,
      isWrite: true,
      skipMetadata: true,
      signal
    })

    const postquery = response.d?.postquery
    if (!postquery) {
        return { items: [], totalHits: 0, startRow: opts.startRow || 0 }
    }

    const refinerData = postquery.PrimaryQueryResult.RefinementResults

    let refiners: RefinerResult[] | undefined = undefined;
    if (refinerData) {
      refiners = refinerData.Refiners.results.map((r) => ({
        name: r.Name, // e.g., "RefinableString00"
        values: r.Entries.results.map((e) => ({
          label: e.RefinementName, // e.g., "Asia"
          count: e.RefinementCount, // e.g., 42
          token: e.RefinementToken, // Used for the actual filtering logic
        })),
      }))
    }

    const rows = postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results

    let items: Record<string, unknown>[] = rows.map((r) => {
      const map: Record<string, unknown> = {}
      r.Cells.results.forEach((c) => (map[c.Key] = c.Value))

      // We prioritize DefaultEncodingURL as it is the most standard direct link
      let directUrl = [
        map.DefaultEncodingURL,
        map.PictureURL,
        map.OriginalPath,
        map.Path,
      ].find((url) => typeof url === 'string' && !url.toLowerCase().includes('dispform.aspx')) as string | undefined

      // If we have a filename in 'Title' and a 'Path', we can often swap out the Form part
      if ((!directUrl || directUrl.toLowerCase().includes('dispform.aspx')) && typeof map.Path === 'string') {
        const base = (map.OriginalPath as string) || (map.Path) || ''
        if (base.includes('/Forms/DispForm.aspx')) {
          const libraryPath = base.split('/Forms/')[0]
          const fileName = (map.Title as string) || ''
          if (fileName.includes('.')) {
            directUrl = `${libraryPath}/${fileName}`
          }
        }
      }

      map.DirectLink = directUrl || map.Path

      if (opts.includeRelativePath && typeof map.DirectLink === 'string') {
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
        items.map(async (item) => {
          if (typeof item.ListId === 'string' && (typeof item.ListItemId === 'string' || typeof item.ListItemId === 'number')) {
            try {
              const params: string[] = []
              if (listSelect.length > 0)
                params.push(`$select=${listSelect.join(',')}`)
              if (userExpands.length > 0)
                params.push(`$expand=${userExpands.join(',')}`)
              const qs = params.length > 0 ? `?${params.join('&')}` : ''

              const hydrated = await this.request<Record<string, unknown>>(
                `/_api/web/lists/getById('${item.ListId}')/items(${item.ListItemId})${qs}`,
                { signal }
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
      const mapping = opts.mapping
      items = items.map((map) => {
        const out: Record<string, unknown> = {}
        Object.entries(mapping).forEach(([k, v]) => {
          // 1. Get Value from Source
          let val: unknown = map
          if (k.includes('.')) {
            const parts = k.split('.')
            for (const p of parts) {
                if (val && typeof val === 'object' && p in (val)) {
                    val = (val as Record<string, unknown>)[p]
                } else {
                    val = null
                }
            }
          } else {
            val = map[k]
          }

          // 2. Assign Value to Destination
          if (v.includes('.')) {
            const parts = v.split('.')
            let current = out
            for (let i = 0; i < parts.length - 1; i++) {
              const part = parts[i]
              if (!current[part] || typeof current[part] !== 'object') {
                  current[part] = {}
              }
              current = current[part] as Record<string, unknown>
            }
            current[parts[parts.length - 1]] = val
          } else {
            out[v] = val
          }
        })
        if (!out.url && typeof map.Path === 'string') out.url = map.Path
        if (opts.includeRelativePath) out.relativePath = map.relativePath
        return out
      })
    }

    return {
      items: items as unknown as T[],
      totalHits: postquery.PrimaryQueryResult.RelevantResults.TotalRows,
      startRow: opts.startRow || 0,
      refiners: refiners,
    }
  }

  // --- Cached Metadata Methods ---

  /** Gets information about the current user */
  async getCurrentUser(signal?: AbortSignal): Promise<UserInfo> {
    const cacheKey = 'CurrentUser'
    const cached = this.cache.get<UserInfo>(cacheKey)
    if (cached) return cached

    const data = await this.request<UserInfo>(`/_api/web/currentuser`, { signal })
    this.cache.set(cacheKey, data)
    return data
  }

  /** Gets all visible fields for a list */
  async getListFields(list: string, webUrl?: string, signal?: AbortSignal): Promise<FieldDefinition[]> {
    const cacheKey = `Fields:${webUrl || 'current'}:${list}`
    const cached = this.cache.get<FieldDefinition[]>(cacheKey)
    if (cached) return cached

    const endpoint = `/_api/web/lists/getbytitle('${list}')/fields?$filter=Hidden eq false`;
    const requestOptions = { signal, targetPath: webUrl };

    const data = await this.request<Array<Record<string, unknown>>>(endpoint, requestOptions)

    const mapped = data.map((f) => ({
      InternalName: f.InternalName as string,
      Title: f.Title as string,
      TypeAsString: f.TypeAsString as string,
      Hidden: f.Hidden as boolean,
      Choices: (f.Choices as { results: string[] })?.results || [],
      TermSetId: (f.TermSetId as string) || undefined
    }))

    this.cache.set(cacheKey, mapped)
    return mapped
  }

  /** Gets choices for a Choice field */
  async getFieldChoices(list: string, field: string, signal?: AbortSignal): Promise<string[]> {
    const cacheKey = `Choices:${list}:${field}`
    const cached = this.cache.get<string[]>(cacheKey)
    if (cached) return cached

    const data = await this.request<Record<string, unknown>>(
      `/_api/web/lists/getbytitle('${list}')/fields/getByInternalNameOrTitle('${field}')`,
      { signal }
    )

    const choices = (data.Choices as { results: string[] })?.results || []
    this.cache.set(cacheKey, choices)
    return choices
  }

  /** Searches for a taxonomy term by label */
  async searchTerm(
    termSetId: string,
    label: string,
    signal?: AbortSignal
  ): Promise<{ Label: string; TermGuid: string } | null> {
    try {
      const safeLabel = label.replace(/'/g, "''")
      const endpoint = `/_api/v2.1/termStore/termSets/${termSetId}/terms?$filter=labels/any(l:l/name eq '${safeLabel}') or name eq '${safeLabel}'&$select=id,name,labels`
      const response = await this.request<SPODataResponse<Array<{ id: string, name: string, names?: Array<{ name: string }> }>>>(endpoint, { signal })
      const terms = response.value
      if (terms && terms.length > 0) {
        const t = terms[0]
        return {
            Label: t.names?.[0]?.name || t.name || label,
            TermGuid: t.id
        }
      }
    } catch (e: unknown) {
      this.logger.warn(`SearchTerm failed for ${label} in ${termSetId}`, [e])
    }
    return null
  }

  // --- Other Methods (Standard) ---
  private buildKql(opts: SearchRequestOptions): string {
    const parts: string[] = []
    const txt = opts.query?.trim() || '*'

    if (txt !== '*') {
      const escapedTxt = txt.replace(/"/g, '""')
      const fields = opts.searchFields && opts.searchFields.length > 0 ? opts.searchFields : ['Title', 'Filename']
      const fieldSearch = fields.map(f => `${f}:"${escapedTxt}*"`).join(' OR ')
      if (opts.searchTitleOnly) {
        parts.push(`(${fieldSearch})`)
      } else {
        parts.push(`(${fieldSearch} OR ${escapedTxt}*)`)
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
          parts.push(`(${value.map((v) => `${mp}:"${String(v)}"`).join(' OR ')})`)
        else parts.push(`${mp}:"${String(value)}"`)
      }
    }
    return parts.join(' AND ')
  }

  /** Creates a new item in a list */
  async createListItem<T>(list: string, payload: Record<string, unknown>, signal?: AbortSignal): Promise<T> {
    const type = await this.getListEntityType(list);
    const body = {
      __metadata: { type },
      ...payload
    };

    return await this.request<T>(`/_api/web/lists/getbytitle('${list}')/items`, {
      method: 'POST',
      body,
      isWrite: true,
      signal
    })
  }

  /** Updates an existing item in a list */
  async updateListItem(list: string, id: number, payload: Record<string, unknown>, signal?: AbortSignal): Promise<void> {
    const type = await this.getListEntityType(list);
    const body = {
      __metadata: { type },
      ...payload
    };

    await this.request(`/_api/web/lists/getbytitle('${list}')/items(${id})`, {
      method: 'POST',
      body,
      headers: {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      signal
    })
  }

  /** Recycles a list item */
  async deleteListItem(list: string, id: number, signal?: AbortSignal): Promise<void> {
    await this.request(`/_api/web/lists/getbytitle('${list}')/items(${id})`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      signal
    })
  }

  /** Retrieves a single list item by its ID */
  async getListItemById<T>(
    list: string,
    id: number,
    select?: string[],
    expand?: string[],
    signal?: AbortSignal
  ): Promise<T> {
    const params: string[] = []
    if (select && select.length > 0) params.push(`$select=${select.join(',')}`)
    if (expand && expand.length > 0) params.push(`$expand=${expand.join(',')}`)

    const q = params.length > 0 ? `?${params.join('&')}` : ''
    return await this.request<T>(
      `/_api/web/lists/getbytitle('${list}')/items(${id})${q}`,
      { signal }
    )
  }

  /** Retrieves multiple items from a list */
  async getListItems<T>(
    listTitle: string,
    options?: ListItemQueryOptions,
    signal?: AbortSignal
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
    const endpoint = `/_api/web/lists/getbytitle('${listTitle}')/items${q}`

    if (options?.getAll) {
      let allItems: T[] = []
      let nextUrl: string | undefined = endpoint

      interface ListResponse {
        d?: {
          results?: unknown
          __next?: string
        }
        value?: unknown
        '@odata.nextLink'?: string
        'odata.nextLink'?: string
      }

      while (nextUrl) {
        const response: ListResponse = await this.request<ListResponse>(nextUrl, {
          skipMetadata: true,
          signal
        })

        if (response && response.d) {
          const results = response.d.results !== undefined ? response.d.results : response.d
          if (Array.isArray(results)) {
            allItems = allItems.concat(results as T[])
            if (options?.onProgress) {
              options.onProgress(results as Record<string, unknown>[])
            }
          } else {
            allItems.push(results as T)
            if (options?.onProgress) {
              options.onProgress([results as Record<string, unknown>])
            }
          }
          nextUrl = response.d.__next
        } else if (response && response.value) {
          allItems = allItems.concat(response.value as T[])
          if (options?.onProgress) {
            options.onProgress(response.value as Record<string, unknown>[])
          }
          nextUrl = response['@odata.nextLink'] || response['odata.nextLink']
        } else {
          nextUrl = undefined
        }
      }
      return allItems
    }

    return await this.request<T[]>(endpoint, { signal })
  }

  /** Gets attachment files for a list item */
  async getItemAttachments(
    list: string,
    id: number,
    signal?: AbortSignal
  ): Promise<AttachmentInfo[]> {
    const results = await this.request<AttachmentInfo[]>(
      `/_api/web/lists/getbytitle('${list}')/items(${id})/AttachmentFiles`,
      { signal }
    )
    return results.map((a) => ({
      FileName: a.FileName,
      ServerRelativeUrl: a.ServerRelativeUrl,
    }))
  }

  /** Adds an attachment to a list item */
  async addAttachment(
    list: string,
    id: number,
    fileName: string,
    file: Blob | ArrayBuffer,
    signal?: AbortSignal
  ): Promise<void> {
    // For binary upload, we need to ensure Content-Type is not application/json
    await this.request(
      `/_api/web/lists/getbytitle('${list}')/items(${id})/AttachmentFiles/add(FileName='${fileName}')`,
      {
        method: 'POST',
        body: file,
        isWrite: true,
        headers: {
          'Content-Type': 'application/octet-stream',
        },
        signal
      }
    )
  }

  /** Deletes an attachment from a list item */
  async deleteAttachment(
    list: string,
    id: number,
    fileName: string,
    signal?: AbortSignal
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
        signal
      }
    )
  }

  /** Uploads a file to a specific folder */
  async uploadFile(url: string, name: string, file: Blob | ArrayBuffer, signal?: AbortSignal): Promise<string> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)

    const data = await this.request<{ ServerRelativeUrl: string }>(
      `/_api/web/getfolderbyserverrelativeurl('${fullUrl}')/files/add(url='${name}', overwrite=true)`,
      {
        method: 'POST',
        body: file,
        headers: {
          'Content-Type': 'application/octet-stream',
        },
        isWrite: true,
        targetPath: fullUrl, // Triggers central getApiRoot
        signal
      }
    )
    return data.ServerRelativeUrl
  }

  /** Downloads a file as a Blob */
  async downloadFile(url: string, signal?: AbortSignal): Promise<Blob> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    const apiRoot = this.getApiRoot(fullUrl)

    const headers = await this.getHeaders(false)
    const res = await fetch(
      `${apiRoot}/_api/web/getfilebyserverrelativeurl('${fullUrl}')/$value`,
      { headers, signal: signal }
    )
    return await res.blob()
  }

  /** Updates metadata for a file */
  async updateFileMetadata(url: string, payload: Record<string, unknown>, signal?: AbortSignal): Promise<void> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)

    // 1. Get List Item details
    const meta = await this.request<{ Id: number, ParentList: { Title: string } }>(
      `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/ListItemAllFields?$expand=ParentList`,
      { targetPath: fullUrl, signal } // Triggers central getApiRoot
    )

    if (!meta || !meta.ParentList) {
      throw new Error(`Could not determine List for file: ${url}`)
    }

    const listTitle = meta.ParentList.Title
    const itemId = meta.Id

    // 2. Adapt Payload
    const formValues = await adaptFileMetadata(this, listTitle, payload)

    // 3. Execute ValidateUpdateListItem on the Item
    const response = await this.request<SPVerboseResponse<{ ValidateUpdateListItem: { results: Array<{ FieldName: string, ErrorCode: number, ErrorMessage: string }> } }>>(
      `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/ValidateUpdateListItem`,
      {
        method: 'POST',
        body: {
          formValues,
          bNewDocumentUpdate: false
        },
        isWrite: true,
        targetPath: fullUrl, // Use file path to determine site again
        signal
      }
    )

    const results = response.d?.ValidateUpdateListItem?.results || []

    const errors = results.filter((r) => r.ErrorCode !== 0)
    if (errors.length > 0) {
        const msg = errors.map((e) => `${e.FieldName}: ${e.ErrorMessage}`).join('; ')
        throw new Error(`UpdateFileMetadata failed: ${msg}`)
    }
  }

  /** Recycles a file */
  async deleteFile(url: string, signal?: AbortSignal): Promise<void> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    await this.request(`/_api/web/getfilebyserverrelativeurl('${fullUrl}')`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      targetPath: fullUrl,
      signal
    })
  }

  /** Creates a folder at the specified path */
  async createFolder(url: string, signal?: AbortSignal): Promise<void> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    await this.request(`/_api/web/folders`, {
      method: 'POST',
      body: {
        __metadata: { type: 'SP.Folder' },
        ServerRelativeUrl: fullUrl,
      },
      isWrite: true,
      targetPath: fullUrl,
      signal
    })
  }

  /** Gets subwebs of the current web */
  async getSubwebs(signal?: AbortSignal): Promise<WebInfo[]> {
    const webs = await this.request<WebInfo[]>(`/_api/web/webs`, { signal })
    return webs.map((w) => ({
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }))
  }

  /** Gets all lists in the current web */
  async getLists(signal?: AbortSignal): Promise<ListInfo[]> {
    const lists = await this.request<ListInfo[]>(`/_api/web/lists`, { signal })
    return lists.map((l) => ({
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl,
    }))
  }

  /** Gets a single list by its title */
  async getList(listTitle: string, signal?: AbortSignal): Promise<ListInfo> {
    const l = await this.request<ListInfo>(
      `/_api/web/lists/getbytitle('${listTitle}')`,
      { signal }
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

  /** Creates a new list */
  async createList(
    title: string,
    description?: string,
    template = 100,
    signal?: AbortSignal
  ): Promise<ListInfo> {
    const l = await this.request<ListInfo>(`/_api/web/lists`, {
      method: 'POST',
      isWrite: true,
      body: {
        __metadata: { type: 'SP.List' },
        Title: title,
        Description: description,
        BaseTemplate: template,
      },
      signal
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

  /** Deletes a list by its title */
  async deleteList(title: string, signal?: AbortSignal): Promise<void> {
    await this.request(`/_api/web/lists/getbytitle('${title}')`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
      signal
    })
  }

  // --- Helpers ---
  private async getListEntityType(listTitle: string): Promise<string> {
    const cacheKey = `ListEntityType:${listTitle}`;
    const cached = this.cache.get<string>(cacheKey);
    if (cached) return cached;

    try {
      const list = await this.request<{ ListItemEntityTypeFullName: string }>(`/_api/web/lists/getbytitle('${listTitle}')?$select=ListItemEntityTypeFullName`);
      const type = list.ListItemEntityTypeFullName || 'SP.Data.ListItem';
      this.cache.set(cacheKey, type);
      return type;
    } catch {
      return 'SP.Data.ListItem';
    }
  }

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
    const targetUrl = customBaseUrl || this.baseUrl;
    const isDefault = targetUrl === this.baseUrl;

    if (isDefault && this.digestCache && Date.now() < this.digestExpiry)
      return this.digestCache

    const h = new Headers({ Accept: 'application/json;odata=verbose' })
    if (this.authProvider) {
        const authHeaders = await this.authProvider()
        Object.entries(authHeaders).forEach(([k, v]) => h.set(k, v))
    }

    const res = await fetch(`${targetUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: h,
    })
    if (!res.ok) return ''
    const response = (await res.json()) as SPVerboseResponse<{ GetContextWebInformation: { FormDigestValue: string, FormDigestTimeoutSeconds: number } }>

    const contextInfo = response.d?.GetContextWebInformation
    if (!contextInfo) return ''

    const token = contextInfo.FormDigestValue

    if (isDefault) {
        this.digestCache = token
        this.digestExpiry =
          Date.now() +
          contextInfo.FormDigestTimeoutSeconds * 1000 -
          60000
    }
    return token
  }

  private getApiRoot(urlOrPath: string): string {
    let origin = ''
    try {
        origin = new URL(this.baseUrl).origin
    } catch {
        return this.baseUrl // Fallback
    }

    const serverRelativePath = getServerRelativePath(urlOrPath, this.baseUrl)
    const path = serverRelativePath.startsWith('/') ? serverRelativePath : `/${serverRelativePath}`
    const parts = path.split('/').filter(p => p)

    if (parts.length >= 2) {
        const category = parts[0].toLowerCase()
        if (category === 'sites' || category === 'teams') {
            return `${origin}/${parts[0]}/${parts[1]}`
        }
    }
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

  /** Gets all users on the site */
  async getSiteUsers(signal?: AbortSignal): Promise<UserInfo[]> {
    return await this.request<UserInfo[]>(
      `/_api/web/siteusers?$filter=PrincipalType eq 1`,
      { signal }
    )
  }

  /** Searches for users by title or email */
  async searchUsers(query: string, signal?: AbortSignal): Promise<UserInfo[]> {
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
      const response = await this.request<{ ClientPeoplePickerSearchUser: string }>(pickerEndpoint, {
        method: 'POST',
        body: payload,
        isWrite: true,
        signal
      });

      const pickerResults = JSON.parse(response.ClientPeoplePickerSearchUser) as Array<{ DisplayText: string, EntityData: { Email?: string }, Description: string, Key: string }>;
      finalResults = pickerResults.map((r) => ({
        Id: 0, // Picker doesn't give Site ID
        Title: r.DisplayText,
        Email: r.EntityData.Email || r.Description,
        LoginName: r.Key
      }));
    } catch (err: unknown) {
      this.logger.error('People Picker failed', [err]);
    }

    //fallback for user search, this is needed if the general search does not match enough results to do more of a Fuzzy search for users
    if (finalResults.length < 5 && query.length > 2) {
      try {
        const searchTerms = query.trim().split(/\s+/).filter(t => t.length > 0);
        const toTitleCase = (str: string) =>
            str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();

        const filterConditions = searchTerms.map(term => {
          const cleanTerm = encodeURIComponent(toTitleCase(term.replace(/'/g, "''")));
          return `(substringof('${cleanTerm}', Title) or substringof('${cleanTerm}', Email))`;
        });

        const filterString = filterConditions.join(' or ');
        const siteUsersUrl = `/_api/web/siteusers?$filter=${filterString}&$top=20`;

        const siteUsers = await this.request<UserInfo[]>(siteUsersUrl, { signal });

        const lowerQueryParts = searchTerms.map(t => t.toLowerCase());

        siteUsers.forEach((u) => {
          const titleLower = (u.Title || "").toLowerCase();
          const emailLower = (u.Email || "").toLowerCase();

          const isMatch = lowerQueryParts.every(part =>
              titleLower.includes(part) || emailLower.includes(part)
          );

          if (isMatch) {
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

      } catch (siteErr: unknown) {
        this.logger.warn('Site User fallback search failed', [siteErr]);
      }
    }

    return finalResults;
  }

  /** Ensures a user exists on the site */
  async ensureUser(loginName: string, signal?: AbortSignal): Promise<UserInfo> {
    const data = await this.request<UserInfo>(`/_api/web/ensureUser`, {
      method: 'POST',
      isWrite: true,
      body: { logonName: loginName },
      signal
    })
    return data
  }

  /** Gets groups for a specific user or current user */
  async getUserGroups(email?: string, signal?: AbortSignal): Promise<SiteGroup[]> {
    let endpoint = ''

    if (email) {
      const user = await this.getUserByEmail(email)
      endpoint = `/_api/web/siteusers/getByLoginName(@v)/groups?@v='${encodeURIComponent(
        user.LoginName
      )}'`
    } else {
      endpoint = `/_api/web/currentuser/groups`
    }

    return await this.request<SiteGroup[]>(endpoint, { signal })
  }

  /** Adds a user to a SharePoint group */
  async addUserToGroup(groupName: string, loginName: string, signal?: AbortSignal): Promise<void> {
    const user = await this.ensureUser(loginName)
    await this.request(`/_api/web/sitegroups/getByName('${groupName}')/users`, {
      method: 'POST',
      isWrite: true,
      body: { LoginName: user.LoginName },
      signal
    })
  }

  /** Removes a user from a SharePoint group */
  async removeUserFromGroup(
    groupName: string,
    loginName: string,
    signal?: AbortSignal
  ): Promise<void> {
    await this.request(
      `/_api/web/sitegroups/getByName('${groupName}')/users/removeByLoginName(@v)?@v='${encodeURIComponent(
        loginName
      )}'`,
      {
        method: 'POST',
        isWrite: true,
        signal
      }
    )
  }

  /** Creates a new SharePoint group */
  async createGroup(
    groupName: string,
    description?: string,
    signal?: AbortSignal
  ): Promise<SiteGroup> {
    return await this.request<SiteGroup>(`/_api/web/sitegroups`, {
      method: 'POST',
      isWrite: true,
      body: {
        __metadata: { type: 'SP.Group' },
        Title: groupName,
        Description: description,
      },
      signal
    })
  }

  /** Gets effective permissions for a user */
  async getUserEffectivePermissions(
    email?: string,
    signal?: AbortSignal
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

    return await this.request<SPBasePermissions>(endpoint, { signal })
  }

  // Helper to resolve Email -> LoginName
  private async getUserByEmail(email: string): Promise<UserInfo> {
    const users = await this.request<UserInfo[]>(
      `/_api/web/siteusers?$filter=Email eq '${email}'`
    )
    const user = users[0]
    if (!user)
      throw new Error(`User with email ${email} not found on this site.`)
    return user
  }

  /** Gets version history for a file */
  async getFileVersions(url: string, signal?: AbortSignal): Promise<FileVersion[]> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    const endpoint = `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/versions?$expand=CreatedBy`

    const results = await this.request<Array<Record<string, unknown>>>(endpoint, { signal })

    return results.map((v) => ({
      VersionLabel: v.VersionLabel as string,
      Created: v.Created as string,
      CheckInComment: v.CheckInComment as string,
      IsCurrentVersion: v.IsCurrentVersion as boolean,
      Size: parseInt(v.Size as string, 10),
      Url: v.Url as string,
      CreatedBy: {
        Id: (v.CreatedBy as UserInfo)?.Id,
        Title: (v.CreatedBy as UserInfo)?.Title,
        Email: (v.CreatedBy as UserInfo)?.Email,
      },
    }))
  }

  /** Gets a link to the version history page for a file */
  async getVersionHistoryLink(serverRelativeUrl: string, signal?: AbortSignal): Promise<string> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const safeUrl = fullUrl.replace(/'/g, "''")

    const data = await this.request<{ Id: number, ParentList: { Id: string } }>(
      `/_api/web/getFileByServerRelativeUrl('${safeUrl}')/ListItemAllFields?$select=Id,ParentList/Id&$expand=ParentList`,
      {
        method: 'GET',
        targetPath: fullUrl, // vital for calculating the correct apiRoot
        signal
      }
    )

    if (!data || !data.ParentList) {
      throw new Error(`Could not resolve List ID for file: ${fullUrl}`)
    }

    const itemId = data.Id
    const listId = data.ParentList.Id

    const fileSiteRoot = this.getApiRoot(fullUrl)
    return this.getVersionHistoryLinkByItem(listId, itemId, fileSiteRoot)
  }

  /** Generates a version history link from list ID and item ID */
  getVersionHistoryLinkByItem(
    listId: string,
    itemId: number,
    webUrl?: string
  ): string {
    const root = webUrl ? webUrl.replace(/\/$/, '') : this.baseUrl
    return `${root}/_layouts/15/Versions.aspx?list={${listId}}&ID=${itemId}`
  }

  private isSearchOnlyProp(field: string): boolean {
    const searchOnlyPattern =
      /^(Refinable|HitHighlighted|Path$|OriginalPath$|Rank$|DocId$|WorkId$|Piw)/i
    return searchOnlyPattern.test(field)
  }
}
