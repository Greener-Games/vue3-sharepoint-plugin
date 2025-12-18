import {
  ISharePointClient,
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
import { Logger } from '../utils/debug'

// Internal Type (Fixes your error)
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
    this.baseUrl = options.baseUrl.replace(/\/$/, '')
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
      } = {}
  ): Promise<T> {
    const { method = 'GET', body, isWrite = false } = options

    const headers = await this.getHeaders(isWrite)
    if (options.headers) {
      Object.entries(options.headers).forEach(([k, v]) => headers.set(k, v))
    }

    const fullUrl = endpoint.startsWith('http')
        ? endpoint
        : `${this.baseUrl}${endpoint}`

    const fetchOptions: RequestInit = {
      method,
      headers,
    }

    if (body) {
      // Check for Blob or ArrayBuffer to avoid JSON stringification
      const isBinary =
          body instanceof Blob ||
          body instanceof ArrayBuffer ||
          (typeof Buffer !== 'undefined' && Buffer.isBuffer(body))

      fetchOptions.body =
          isBinary || typeof body === 'string' ? body : JSON.stringify(body)
    }

    this.logger.log(`Request: ${method} ${fullUrl}`, body ? { body } : '')

    const response = await fetch(fullUrl, fetchOptions)

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
  async executeBatch(builder: (batch: IBatch) => void): Promise<void> {
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
    await this.processInternalBatch(queue)
  }

  private async processInternalBatch(requests: InternalBatchItem[]) {
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
    })
    if (!res.ok) throw new Error(`Batch failed: ${await res.text()}`)
  }

  // --- Search ---
  async search<T = any>(opts: SearchRequestOptions): Promise<SearchResult<T>> {
    const kql = this.buildKql(opts)
    this.logger.log(`Search KQL: ${kql}`)

    const select = opts.selectFields || [
      'Title',
      'Path',
      'OriginalPath',
      'UniqueId',
      'HitHighlightedSummary',
      ...Object.keys(opts.mapping || {}),
    ]
    const payload = {
      request: {
        Querytext: kql,
        RowLimit: opts.rowLimit || 10,
        StartRow: opts.startRow || 0,
        SelectProperties: { results: select },
        HitHighlightedProperties: { results: ['Contents', 'Title'] },
        SummaryLength: 250,
        EnableStemming: true,
        TrimDuplicates: false,
      },
    }

    // Search returns complex object, using skipMetadata to handle manual parsing
    const data = await this.request<any>(`/_api/search/postquery`, {
      method: 'POST',
      body: payload,
      isWrite: true,
      skipMetadata: true,
    })

    const rows =
        data.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results

    const items = rows.map((r: any) => {
      const map: any = {}
      r.Cells.results.forEach((c: any) => (map[c.Key] = c.Value))

      if (opts.includeRelativePath && map.Path) {
        try {
          map.relativePath = decodeURIComponent(new URL(map.Path).pathname)
        } catch {
          map.relativePath = map.Path
        }
      }

      if (opts.mapping) {
        const out: any = {}
        Object.entries(opts.mapping).forEach(([k, v]) => (out[v] = map[k]))
        if (!out.url) out.url = map.Path
        if (opts.includeRelativePath) out.relativePath = map.relativePath
        return out
      }
      return map
    })
    return {
      items,
      totalHits: data.d.postquery.PrimaryQueryResult.RelevantResults.TotalRows,
      startRow: opts.startRow || 0,
    }
  }

  // --- Cached Metadata Methods ---

  async getCurrentUser() {
    const cacheKey = 'CurrentUser'
    const cached = this.cache.get<any>(cacheKey)
    if (cached) return cached

    const data = await this.request<any>(`/_api/web/currentuser`)
    this.cache.set(cacheKey, data)
    return data
  }

  async getListFields(list: string) {
    const cacheKey = `Fields:${list}`
    const cached = this.cache.get<FieldDefinition[]>(cacheKey)
    if (cached) return cached

    const data = await this.request<any[]>(
        `/_api/web/lists/getbytitle('${list}')/fields?$filter=Hidden eq false`
    )

    const mapped = data.map((f: any) => ({
      InternalName: f.InternalName,
      Title: f.Title,
      TypeAsString: f.TypeAsString,
      Hidden: f.Hidden,
      Choices: f.Choices?.results || [],
    }))

    this.cache.set(cacheKey, mapped)
    return mapped
  }

  async getFieldChoices(list: string, field: string) {
    const cacheKey = `Choices:${list}:${field}`
    const cached = this.cache.get<string[]>(cacheKey)
    if (cached) return cached

    const data = await this.request<any>(
        `/_api/web/lists/getbytitle('${list}')/fields/getByInternalNameOrTitle('${field}')`
    )

    const choices = data.Choices?.results || []
    this.cache.set(cacheKey, choices)
    return choices
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

    if (opts.fileTypes?.include?.length) {
      parts.push(
          `(${opts.fileTypes.include
              .map((e) => `FileExtension:${e}`)
              .join(' OR ')})`
      )
    }
    if (opts.fileTypes?.exclude?.length) {
      parts.push(
          `(${opts.fileTypes.exclude
              .map((e) => `NOT FileExtension:${e}`)
              .join(' AND ')})`
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
      // Exclude Folders (0x0120 is Folder ContentType)
      parts.push(`(NOT ContentTypeId:0x0120*)`)
    } else if (opts.resultType === 'folders') {
      // Only Folders
      parts.push(`ContentTypeId:0x0120*`)
    }

    if (opts.filters) {
      for (const [key, value] of Object.entries(opts.filters)) {
        if (!value || (Array.isArray(value) && value.length === 0)) continue
        let mp = key
        if (opts.mapping) {
          const found = Object.keys(opts.mapping).find(
              (k) => opts.mapping![k] === key
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

  async createListItem(list: string, payload: any) {
    return await this.request(`/_api/web/lists/getbytitle('${list}')/items`, {
      method: 'POST',
      body: payload,
      isWrite: true,
    })
  }

  async updateListItem(list: string, id: number, payload: any) {
    await this.request(`/_api/web/lists/getbytitle('${list}')/items(${id})`, {
      method: 'POST',
      body: payload,
      headers: {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
      },
      isWrite: true,
    })
  }

  async deleteListItem(list: string, id: number) {
    await this.request(`/_api/web/lists/getbytitle('${list}')/items(${id})`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
    })
  }

  async getListItemById(list: string, id: number, select?: string[]) {
    const q = select ? `?$select=${select.join(',')}` : ''
    return await this.request(
        `/_api/web/lists/getbytitle('${list}')/items(${id})${q}`
    )
  }

  async getItemAttachments(
      list: string,
      id: number
  ): Promise<AttachmentInfo[]> {
    const results = await this.request<any[]>(
        `/_api/web/lists/getbytitle('${list}')/items(${id})/AttachmentFiles`
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
      file: Blob | ArrayBuffer
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
        }
    )
  }

  async deleteAttachment(
      list: string,
      id: number,
      fileName: string
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
        }
    )
  }

  async uploadFile(url: string, name: string, file: any) {
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
        }
    )
    return data.ServerRelativeUrl
  }

  async downloadFile(url: string) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    // We can't use centralized request for Blob return nicely without generic complexity or flag
    // So sticking to native fetch for blob download, but using helper for headers could be good.
    // However, downloadFile logic is simple enough.
    const headers = await this.getHeaders(false)
    const res = await fetch(
        `${this.baseUrl}/_api/web/getfilebyserverrelativeurl('${fullUrl}')/$value`,
        { headers }
    )
    return await res.blob()
  }

  async updateFileMetadata(url: string, payload: any) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    // 1. Get List Item URI
    const meta = await this.request<any>(
        `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/ListItemAllFields`
    )
    const uri = meta.__metadata.uri
    const type = meta.__metadata.type

    // 2. Update
    await this.request(uri, {
      method: 'POST',
      body: { __metadata: { type }, ...payload },
      headers: {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
      },
      isWrite: true,
    })
  }

  async deleteFile(url: string) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    await this.request(`/_api/web/getfilebyserverrelativeurl('${fullUrl}')`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
    })
  }

  async createFolder(url: string) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    await this.request(`/_api/web/folders`, {
      method: 'POST',
      body: {
        __metadata: { type: 'SP.Folder' },
        ServerRelativeUrl: fullUrl,
      },
      isWrite: true,
    })
  }

  // --- Webs & Lists ---
  async getWebInfo(): Promise<WebInfo> {
    const w = await this.request<any>(`/_api/web`)
    return {
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }
  }

  async getSubwebs(): Promise<WebInfo[]> {
    const webs = await this.request<any[]>(`/_api/web/webs`)
    return webs.map((w) => ({
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }))
  }

  async getLists(): Promise<ListInfo[]> {
    const lists = await this.request<any[]>(`/_api/web/lists`)
    return lists.map((l) => ({
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl,
    }))
  }

  async getList(listTitle: string): Promise<ListInfo> {
    const l = await this.request<any>(
        `/_api/web/lists/getbytitle('${listTitle}')`
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
      template = 100
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

  async deleteList(title: string): Promise<void> {
    await this.request(`/_api/web/lists/getbytitle('${title}')`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
      },
      isWrite: true,
    })
  }

  // --- Helpers ---
  private async getHeaders(isWrite = false): Promise<Headers> {
    const headers = new Headers({
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    })
    if (this.authProvider) {
      const a = await this.authProvider()
      Object.entries(a).forEach(([k, v]) => headers.append(k, v))
    }
    if (isWrite) {
      headers.append('X-RequestDigest', await this.getRequestDigest())
    }
    return headers
  }

  private async getRequestDigest(): Promise<string> {
    if (this.digestCache && Date.now() < this.digestExpiry)
      return this.digestCache

    // We use a simplified fetch here to avoid recursion with `request` which might call getHeaders
    const h = new Headers({ Accept: 'application/json;odata=verbose' })
    if (this.authProvider) Object.assign(h, await this.authProvider())

    const res = await fetch(`${this.baseUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: h,
    })
    if (!res.ok) return ''
    const d = await res.json()
    this.digestCache = d.d.GetContextWebInformation.FormDigestValue
    this.digestExpiry =
        Date.now() +
        d.d.GetContextWebInformation.FormDigestTimeoutSeconds * 1000 -
        60000
    return this.digestCache!
  }

  private generateUuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0
      return (c == 'x' ? r : (r & 0x3) | 0x8).toString(16)
    })
  }

  async getSiteUsers(): Promise<UserInfo[]> {
    return await this.request<UserInfo[]>(
        `/_api/web/siteusers?$filter=PrincipalType eq 1`
    )
  }

  async ensureUser(loginName: string): Promise<UserInfo> {
    const data = await this.request<any>(`/_api/web/ensureUser`, {
      method: 'POST',
      isWrite: true,
      body: { logonName: loginName },
    })
    return data
  }

  async getUserGroups(email?: string): Promise<SiteGroup[]> {
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

    return await this.request<SiteGroup[]>(endpoint)
  }

  async addUserToGroup(groupName: string, loginName: string): Promise<void> {
    const user = await this.ensureUser(loginName)
    await this.request(`/_api/web/sitegroups/getByName('${groupName}')/users`, {
      method: 'POST',
      isWrite: true,
      body: { LoginName: user.LoginName },
    })
  }

  async removeUserFromGroup(
      groupName: string,
      loginName: string
  ): Promise<void> {
    await this.request(
        `/_api/web/sitegroups/getByName('${groupName}')/users/removeByLoginName(@v)?@v='${encodeURIComponent(
            loginName
        )}'`,
        {
          method: 'POST',
          isWrite: true,
        }
    )
  }

  async createGroup(
      groupName: string,
      description?: string
  ): Promise<SiteGroup> {
    return await this.request<SiteGroup>(`/_api/web/sitegroups`, {
      method: 'POST',
      isWrite: true,
      body: {
        __metadata: { type: 'SP.Group' },
        Title: groupName,
        Description: description,
      },
    })
  }

  async getUserEffectivePermissions(
      email?: string
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

    return await this.request<SPBasePermissions>(endpoint)
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

  async getFileVersions(url: string): Promise<FileVersion[]> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    // We expand CreatedBy to get the user name
    const endpoint = `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/versions?$expand=CreatedBy`

    const results = await this.request<any[]>(endpoint)

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

  getVersionHistoryLink(serverRelativeUrl: string): string {
    // The classic Versions.aspx page can accept a FileName parameter
    // This works on both SharePoint On-Prem and Online
    return `${
        this.baseUrl
    }/_layouts/15/Versions.aspx?FileName=${encodeURIComponent(
        serverRelativeUrl
    )}`
  }
}
