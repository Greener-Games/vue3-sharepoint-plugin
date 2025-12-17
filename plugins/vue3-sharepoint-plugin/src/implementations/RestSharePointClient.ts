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
} from '../types'

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

  // Digest Caching (Memory only)
  private digestCache: string | null = null
  private digestExpiry: number = 0

  constructor(options: SharePointConfig) {
    this.baseUrl = options.baseUrl.replace(/\/$/, '')
    this.authProvider = options.authProvider
    this.cache = new InternalCache(options.enableCache ?? true)
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
      },
    }

    const headers = await this.getHeaders(true)
    const res = await fetch(`${this.baseUrl}/_api/search/postquery`, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
    })
    if (!res.ok) throw new Error(res.statusText)
    const data = await res.json()
    const rows =
      data.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results

    const items = rows.map((r: any) => {
      const map: any = {}
      r.Cells.results.forEach((c: any) => (map[c.Key] = c.Value))
      if (opts.mapping) {
        const out: any = {}
        Object.entries(opts.mapping).forEach(([k, v]) => (out[v] = map[k]))
        if (!out.url) out.url = map.Path
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

    const headers = await this.getHeaders(false)
    const res = await fetch(`${this.baseUrl}/_api/web/currentuser`, { headers })
    if (!res.ok) throw new Error('User fetch failed')
    const data = await res.json()

    this.cache.set(cacheKey, data.d)
    return data.d
  }

  async getListFields(list: string) {
    const cacheKey = `Fields:${list}`
    const cached = this.cache.get<FieldDefinition[]>(cacheKey)
    if (cached) return cached

    const headers = await this.getHeaders(false)
    const res = await fetch(
      `${this.baseUrl}/_api/web/lists/getbytitle('${list}')/fields?$filter=Hidden eq false`,
      { headers }
    )
    if (!res.ok) throw new Error('Fields fetch failed')
    const data = await res.json()

    const mapped = data.d.results.map((f: any) => ({
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

    const headers = await this.getHeaders(false)
    const res = await fetch(
      `${this.baseUrl}/_api/web/lists/getbytitle('${list}')/fields/getByInternalNameOrTitle('${field}')`,
      { headers }
    )
    if (!res.ok) throw new Error('Choices fetch failed')
    const data = await res.json()

    const choices = data.d.Choices?.results || []
    this.cache.set(cacheKey, choices)
    return choices
  }

  // --- Other Methods (Standard) ---
  private buildKql(opts: SearchRequestOptions): string {
    const parts: string[] = []
    const txt = opts.query?.trim() || '*'
    if (txt !== '*')
      parts.push(
        opts.searchTitleOnly ? `Title:"${txt.replace(/"/g, '""')}*"` : txt
      )
    else parts.push('*')
    if (opts.fileTypes?.length)
      parts.push(
        `(${opts.fileTypes.map((e) => `FileExtension:${e}`).join(' OR ')})`
      )
    if (opts.scope) {
      const scopes = Array.isArray(opts.scope) ? opts.scope : [opts.scope]
      parts.push(`(${scopes.map((s) => `Path:"${s}*"`).join(' OR ')})`)
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

  // (Boilerplate implementations for create/update/upload/download etc...)
  async createListItem(list: string, payload: any) {
    const h = await this.getHeaders(true)
    // Fetch List Type if needed, usually we can guess SP.Data.[List]ListItem but safer to fetch
    // For brevity in this snippet, assuming standard
    const res = await fetch(
      `${this.baseUrl}/_api/web/lists/getbytitle('${list}')/items`,
      { method: 'POST', headers: h, body: JSON.stringify(payload) }
    )
    return (await res.json()).d
  }
  async updateListItem(list: string, id: number, payload: any) {
    const h = await this.getHeaders(true)
    h.append('X-HTTP-Method', 'MERGE')
    h.append('IF-MATCH', '*')
    await fetch(
      `${this.baseUrl}/_api/web/lists/getbytitle('${list}')/items(${id})`,
      { method: 'POST', headers: h, body: JSON.stringify(payload) }
    )
  }
  async deleteListItem(list: string, id: number) {
    const h = await this.getHeaders(true)
    h.append('X-HTTP-Method', 'DELETE')
    h.append('IF-MATCH', '*')
    await fetch(
      `${this.baseUrl}/_api/web/lists/getbytitle('${list}')/items(${id})`,
      { method: 'POST', headers: h }
    )
  }
  async getListItemById(list: string, id: number, select?: string[]) {
    const h = await this.getHeaders(false)
    const q = select ? `?$select=${select.join(',')}` : ''
    const res = await fetch(
      `${this.baseUrl}/_api/web/lists/getbytitle('${list}')/items(${id})${q}`,
      { headers: h }
    )
    return (await res.json()).d
  }
  async uploadFile(url: string, name: string, file: any) {
    const h = await this.getHeaders(true)
    h.delete('Content-Type')
    const res = await fetch(
      `${this.baseUrl}/_api/web/getfolderbyserverrelativeurl('${url}')/files/add(url='${name}', overwrite=true)`,
      { method: 'POST', headers: h, body: file }
    )
    return (await res.json()).d.ServerRelativeUrl
  }
  async downloadFile(url: string) {
    const h = await this.getHeaders(false)
    const res = await fetch(
      `${this.baseUrl}/_api/web/getfilebyserverrelativeurl('${url}')/$value`,
      { headers: h }
    )
    return await res.blob()
  }
  async updateFileMetadata(url: string, payload: any) {
    const h = await this.getHeaders(false)
    const meta = await (
      await fetch(
        `${this.baseUrl}/_api/web/getfilebyserverrelativeurl('${url}')/ListItemAllFields`,
        { headers: h }
      )
    ).json()
    const uri = meta.d.__metadata.uri
    const type = meta.d.__metadata.type
    const h2 = await this.getHeaders(true)
    h2.append('X-HTTP-Method', 'MERGE')
    h2.append('IF-MATCH', '*')
    await fetch(uri, {
      method: 'POST',
      headers: h2,
      body: JSON.stringify({ __metadata: { type }, ...payload }),
    })
  }
  async deleteFile(url: string) {
    const h = await this.getHeaders(true)
    h.append('X-HTTP-Method', 'DELETE')
    h.append('IF-MATCH', '*')
    await fetch(
      `${this.baseUrl}/_api/web/getfilebyserverrelativeurl('${url}')`,
      { method: 'POST', headers: h }
    )
  }
  async createFolder(url: string) {
    const h = await this.getHeaders(true)
    await fetch(`${this.baseUrl}/_api/web/folders`, {
      method: 'POST',
      headers: h,
      body: JSON.stringify({
        __metadata: { type: 'SP.Folder' },
        ServerRelativeUrl: url,
      }),
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
    // We can cache the digest token for ~20 mins (it lasts 30)
    // For this implementation, let's fetch it if missing
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
    const headers = await this.getHeaders(false)
    const res = await fetch(
      `${this.baseUrl}/_api/web/siteusers?$filter=PrincipalType eq 1`,
      { headers }
    )
    return (await res.json()).d.results
  }

  async getUserGroups(email?: string): Promise<SiteGroup[]> {
    const headers = await this.getHeaders(false)
    let endpoint = ''

    if (email) {
      // Step 1: Get the User ID/LoginName from Email
      const user = await this.getUserByEmail(email)
      endpoint = `${
        this.baseUrl
      }/_api/web/siteusers/getByLoginName(@v)/groups?@v='${encodeURIComponent(
        user.LoginName
      )}'`
    } else {
      // Current User
      endpoint = `${this.baseUrl}/_api/web/currentuser/groups`
    }

    const res = await fetch(endpoint, { headers })
    if (!res.ok)
      throw new Error(`Failed to fetch groups for ${email || 'current user'}`)
    return (await res.json()).d.results
  }

  async getUserEffectivePermissions(
    email?: string
  ): Promise<SPBasePermissions> {
    const headers = await this.getHeaders(false)
    let endpoint = ''

    if (email) {
      const user = await this.getUserByEmail(email)
      endpoint = `${
        this.baseUrl
      }/_api/web/getUserEffectivePermissions(@v)?@v='${encodeURIComponent(
        user.LoginName
      )}'`
    } else {
      endpoint = `${this.baseUrl}/_api/web/effectiveBasePermissions`
    }

    const res = await fetch(endpoint, { headers })
    if (!res.ok)
      throw new Error(
        `Failed to fetch permissions for ${email || 'current user'}`
      )
    const data = await res.json()
    return data.d // Returns { High: number, Low: number }
  }

  // Helper to resolve Email -> LoginName
  private async getUserByEmail(email: string): Promise<UserInfo> {
    const headers = await this.getHeaders(false)
    // We filter site users to find the specific email
    const res = await fetch(
      `${this.baseUrl}/_api/web/siteusers?$filter=Email eq '${email}'`,
      { headers }
    )
    const data = await res.json()
    const user = data.d.results[0]
    if (!user)
      throw new Error(`User with email ${email} not found on this site.`)
    return user
  }

  async getFileVersions(url: string): Promise<FileVersion[]> {
    const headers = await this.getHeaders(false)
    // We expand CreatedBy to get the user name
    const endpoint = `${this.baseUrl}/_api/web/getfilebyserverrelativeurl('${url}')/versions?$expand=CreatedBy`

    const res = await fetch(endpoint, { headers })
    if (!res.ok) throw new Error(`Failed to fetch versions for ${url}`)
    const data = await res.json()

    return data.d.results.map((v: any) => ({
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
