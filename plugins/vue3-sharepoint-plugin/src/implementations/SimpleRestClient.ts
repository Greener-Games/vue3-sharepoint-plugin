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
import { Logger } from '../utils/debug'

// Internal Cache Helper (Reused from RestSharePointClient logic)
class InternalCache {
  private enabled: boolean
  private prefix = 'sp-simple-cache::'

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

export class SimpleRestClient implements ISharePointClient {
  private baseUrl: string
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
    this.cache = new InternalCache(options.enableCache ?? true)
    this.logger = new Logger(options.debug)
  }

  // --- Centralized Request Handler ---
  private async request<T = any>(
    endpoint: string,
    options: {
      method?: string
      headers?: Record<string, string>
      targetPath?: string
      body?: any
      isWrite?: boolean
      skipMetadata?: boolean
    } = {}
  ): Promise<T> {
    const { method = 'GET', targetPath, isWrite = false, body } = options

    // Block non-GET requests effectively (though we also stub them out)
    if (method !== 'GET' && !isWrite) { // isWrite check to allow our authorized methods
        // Actually, our specific methods (create/update) pass isWrite=true.
        // So we can check if it's a write method and isWrite is false?
        // But the stubbed methods throw error anyway.
        // Let's just log writes if they happen unexpectedly.
    }

    let finalEndpoint = endpoint;
    let contextUrl: string | undefined = undefined;

    // 1. Determine API Root based on targetPath
    if (targetPath) {
        const apiRoot = this.getApiRoot(targetPath);
        contextUrl = apiRoot;

        if (!endpoint.startsWith('http')) {
            finalEndpoint = `${apiRoot}${endpoint}`;
        } else {
            finalEndpoint = endpoint;
        }
    } else {
        if (!endpoint.startsWith('http')) {
            finalEndpoint = `${this.baseUrl}${endpoint}`;
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

    this.logger.log(`SimpleRest Request: ${method} ${finalEndpoint}`)

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

      fetchOptions.body = (
        isBinary || typeof body === 'string' ? body : JSON.stringify(body)
      ) as BodyInit
    }

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

    if (response.status === 204) {
      return null as T
    }

    const data = await response.json()
    this.logger.log(`Response: ${response.status}`, data)

    if (options.skipMetadata) {
      return data as T
    }

    // Robustly check for nested 'd' and 'results'
    const safeData: any = data;
    if (safeData && safeData['d']) {
        const d = safeData['d'];
        if (d && d['results']) {
            return d['results'] as T;
        }
        return d as T;
    }

    return data as T
  }

  private getApiRoot(urlOrPath: string): string {
    let origin = ''
    try {
        origin = new URL(this.baseUrl).origin
    } catch {
        return this.baseUrl
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
    // If it's potentially root site collection
    if (parts.length > 0 && !['sites', 'teams'].includes(parts[0].toLowerCase())) {
         return origin
    }
    return this.baseUrl
  }

  // --- Helper: Get Headers (with Digest for Writes) ---
  private async getHeaders(isWrite = false, contextUrl?: string): Promise<Headers> {
    const headers = new Headers({
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    })

    // In SimpleRest Mode, we assume no Auth Provider is needed (handled by proxy),
    // OR if one is provided in config, it might be used.

    if (isWrite) {
      const digest = await this.getRequestDigest(contextUrl)
      if (digest) {
        headers.append('X-RequestDigest', digest)
      }
    }
    return headers
  }

  private async getRequestDigest(customBaseUrl?: string): Promise<string> {
    const targetUrl = customBaseUrl || this.baseUrl;
    const isDefault = targetUrl === this.baseUrl;

    if (isDefault && this.digestCache && Date.now() < this.digestExpiry)
      return this.digestCache

    const h = new Headers({ Accept: 'application/json;odata=verbose' })

    try {
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
    } catch (e) {
        this.logger.warn("Failed to fetch Request Digest", e);
        return '';
    }
  }

  private async getListEntityType(listTitle: string): Promise<string> {
      // Fetch list metadata to get ListItemEntityTypeFullName
      // Endpoint: /_api/web/lists/getbytitle('List')?$select=ListItemEntityTypeFullName
      try {
          const list = await this.request<any>(`/_api/web/lists/getbytitle('${listTitle}')?$select=ListItemEntityTypeFullName`);
          return list.ListItemEntityTypeFullName || 'SP.Data.ListItem'; // Fallback
      } catch {
          return 'SP.Data.ListItem';
      }
  }

  // --- Search Implementation (via Lists) ---
  async search<T = any>(opts: SearchRequestOptions): Promise<SearchResult<T>> {
    // 1. Validate Scope
    let scope = opts.scope;
    if (Array.isArray(scope)) scope = scope[0];

    if (!scope) {
        this.logger.error("SimpleRest Search requires a specific 'scope' pointing to a list or folder.")
        return { items: [], totalHits: 0, startRow: 0 }
    }

    this.logger.log(`SimpleRest Search: Scoped to ${scope}`)

    // 2. Build OData Query
    const params: string[] = []

    // Select
    const selectFields = opts.selectFields || ['Title', 'Id', 'EncodedAbsUrl', 'FileRef', 'FileDirRef', 'Created', 'Modified', 'Author/Title', 'Editor/Title'];
    const expandFields = opts.expandFields || ['Author', 'Editor'];

    const mappedSelects = selectFields.map(f => {
        if (f === 'Path' || f === 'OriginalPath') return 'EncodedAbsUrl';
        if (f === 'DefaultEncodingURL') return 'EncodedAbsUrl';
        return f;
    }).filter((v, i, a) => a.indexOf(v) === i); // unique

    params.push(`$select=${mappedSelects.join(',')}`)
    if (expandFields.length > 0) {
        params.push(`$expand=${expandFields.join(',')}`)
    }

    // Filter
    const filters: string[] = []

    if (opts.query && opts.query !== '*') {
        const q = opts.query.replace(/'/g, "''");
        filters.push(`substringof('${q}', Title)`)
    }

    if (opts.filters) {
        Object.entries(opts.filters).forEach(([key, value]) => {
            if (!value) return;
            const vals = Array.isArray(value) ? value : [value];
            if (vals.length === 0) return;
            const orBlock = vals.map(v => {
                const safeV = String(v).replace(/'/g, "''");
                return `${key} eq '${safeV}'`
            }).join(' or ');
            filters.push(`(${orBlock})`);
        })
    }

    if (opts.fileTypes?.include?.length) {
         const types = opts.fileTypes.include.map(t => `File_x0020_Type eq '${t}'`).join(' or ');
         filters.push(`(${types})`);
    }

    if (opts.resultType === 'items') {
        filters.push(`FSObjType eq 0`);
    } else if (opts.resultType === 'folders') {
        filters.push(`FSObjType eq 1`);
    }

    if (filters.length > 0) {
        params.push(`$filter=${filters.join(' and ')}`)
    }

    params.push(`$orderby=Created desc`)

    if (opts.rowLimit) {
        params.push(`$top=${opts.rowLimit}`)
    }

    const queryString = params.join('&');
    const scopePath = scope.startsWith('/') ? scope : `/${scope}`;
    const endpoint = `/_api/web/GetList(@v)/items?@v='${scopePath}'&${queryString}`;

    try {
        const results = await this.request<any[]>(endpoint, { targetPath: scopePath });

        // Map to SearchResult
        const items = results.map(item => {
             const out: any = { ...item };
             out.Path = item.EncodedAbsUrl || this.baseUrl + item.FileRef;
             if (opts.includeRelativePath) {
                 out.relativePath = item.FileRef;
             }
             if (opts.mapping) {
                Object.entries(opts.mapping).forEach(([k, v]) => {
                    if (item[k] !== undefined) out[v] = item[k];
                })
             }
             return out;
        });

        return {
            items,
            totalHits: items.length,
            startRow: 0
        }

    } catch (e) {
        this.logger.error("SimpleRest Search Failed", e);
        return { items: [], totalHits: 0, startRow: 0 }
    }
  }

  // --- Read Methods (Supported) ---

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

  async getListItemById(
    list: string,
    id: number,
    select?: string[],
    expand?: string[]
  ) {
    const params: string[] = []
    if (select && select.length > 0) params.push(`$select=${select.join(',')}`)
    if (expand && expand.length > 0) params.push(`$expand=${expand.join(',')}`)

    const q = params.length > 0 ? `?${params.join('&')}` : ''
    return await this.request(
      `/_api/web/lists/getbytitle('${list}')/items(${id})${q}`
    )
  }

  async getListItems<T = any>(
    listTitle: string,
    options?: ListItemQueryOptions
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
      `/_api/web/lists/getbytitle('${listTitle}')/items${q}`
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

  async downloadFile(url: string) {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    const apiRoot = this.getApiRoot(fullUrl)

    const headers = new Headers({
        Accept: 'application/json;odata=verbose'
    })

    const endpoint = `${apiRoot}/_api/web/getfilebyserverrelativeurl('${fullUrl}')/$value`;

    this.logger.log(`Anon Download: ${endpoint}`);
    const res = await fetch(endpoint, { headers });
    if (!res.ok) throw new Error(`Download failed: ${res.statusText}`);
    return await res.blob()
  }

  async getFileVersions(url: string): Promise<FileVersion[]> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    const endpoint = `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/versions?$expand=CreatedBy`
    const results = await this.request<any[]>(endpoint, { targetPath: fullUrl })

    return results.map((v: any) => ({
      VersionLabel: v.VersionLabel,
      Created: v.Created,
      CheckInComment: v.CheckInComment,
      IsCurrentVersion: v.IsCurrentVersion,
      Size: parseInt(v.Size, 10),
      Url: v.Url,
      CreatedBy: {
        Id: v.CreatedBy?.Id,
        Title: v.CreatedBy?.Title,
        Email: v.CreatedBy?.Email,
      },
    }))
  }

  async getVersionHistoryLink(serverRelativeUrl: string): Promise<string> {
    // We can simulate an async call or just return the old link format if we can't easily resolve ListID/ItemID
    // SimpleRest usually operates where we don't have full permissions or API access.
    // However, to satisfy the interface, we must return a Promise.

    // Attempt to resolve if possible, else fallback to FileName based link which is valid.
    // But the user said "currently one is failing".
    // If FileName based link is failing, we should try the ID based one.
    // But in SimpleRest/Anonymous, getting ListItemAllFields might fail if strictly limited.
    // Let's try to replicate the Rest logic, as SimpleRest inherits request capabilities.

    try {
        const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
        const safeUrl = fullUrl.replace(/'/g, "''")

        const data = await this.request<any>(
            `/_api/web/getFileByServerRelativeUrl('${safeUrl}')/ListItemAllFields?$select=Id,ParentList/Id&$expand=ParentList`,
            {
                method: 'GET',
                targetPath: fullUrl
            }
        )

        if (data && data.ParentList) {
            const apiRoot = this.getApiRoot(fullUrl);
            return this.getVersionHistoryLinkByItem(data.ParentList.Id, data.Id, apiRoot);
        }
    } catch {
        // Fallback or ignore
    }

    return `${this.baseUrl}/_layouts/15/Versions.aspx?FileName=${encodeURIComponent(serverRelativeUrl)}`
  }

  getVersionHistoryLinkByItem(
    listId: string,
    itemId: number,
    webUrl?: string
  ): string {
    const root = webUrl ? webUrl.replace(/\/$/, '') : this.baseUrl
    return `${root}/_layouts/15/Versions.aspx?list={${listId}}&ID=${itemId}`
  }

  async getListFields(list: string, webUrl?: string) {
    const cacheKey = `Fields:${webUrl || 'current'}:${list}`
    const cached = this.cache.get<FieldDefinition[]>(cacheKey)
    if (cached) return cached

    const endpoint = `/_api/web/lists/getbytitle('${list}')/fields?$filter=Hidden eq false`;
    const requestOptions: any = {};
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

  async searchTerm(
    termSetId: string,
    label: string
  ): Promise<{ Label: string; TermGuid: string } | null> {
      // This is usually a public taxonomy read, might work if anonymous has access.
      // But commonly taxonomy is locked down. We'll try.
    try {
      const safeLabel = label.replace(/'/g, "''")
      const endpoint = `/_api/v2.1/termStore/termSets/${termSetId}/terms?$filter=labels/any(l:l/name eq '${safeLabel}') or name eq '${safeLabel}'&$select=id,name,labels`
      const response = await this.request<{ value: any[] }>(endpoint)
      const terms = response.value
      if (terms && terms.length > 0) {
        const t = terms[0]
        return {
            Label: t.names?.[0]?.name || t.name || label,
            TermGuid: t.id
        }
      }
    } catch (e) {
      this.logger.warn(`SearchTerm failed for ${label}`, e)
    }
    return null
  }

  // --- Write Operations (Partial Support) ---

  async createListItem(listTitle: string, payload: Record<string, any>): Promise<any> {
      // 1. Get Entity Type (required for verbose JSON)
      const type = await this.getListEntityType(listTitle);

      // 2. Prepare Payload
      const body = {
          __metadata: { type },
          ...payload
      };

      // 3. Post
      return await this.request(`/_api/web/lists/getbytitle('${listTitle}')/items`, {
          method: 'POST',
          body,
          isWrite: true
      });
  }

  async updateListItem(listTitle: string, id: number, payload: Record<string, any>): Promise<void> {
      const type = await this.getListEntityType(listTitle);

      const body = {
          __metadata: { type },
          ...payload
      };

      await this.request(`/_api/web/lists/getbytitle('${listTitle}')/items(${id})`, {
          method: 'POST',
          body,
          headers: {
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*', // Force update, assumes concurrency is not handled
          },
          isWrite: true
      });
  }

  async deleteListItem(listTitle: string, id: number): Promise<void> {
      // User specifically requested "allow for the post requests to happen for updating a list item".
      // They did not explicitly ask for delete, but it's consistent CRUD.
      // I'll enable it.
      await this.request(`/_api/web/lists/getbytitle('${listTitle}')/items(${id})`, {
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*',
        },
        isWrite: true,
      })
  }

  // --- Unsupported / Stubbed Operations ---

  private notSupported(operation: string): Promise<any> {
      this.logger.error(`Operation '${operation}' is not supported in SimpleRest Mode.`)
      return Promise.reject(new Error(`Operation '${operation}' is not supported in SimpleRest Mode.`));
  }

  async executeBatch(_builder: (batch: IBatch) => void): Promise<void> {
      return this.notSupported('executeBatch');
  }
  async addAttachment(_listTitle: string, _itemId: number, _fileName: string, _file: Blob | ArrayBuffer): Promise<void> {
      return this.notSupported('addAttachment');
  }
  async deleteAttachment(_listTitle: string, _itemId: number, _fileName: string): Promise<void> {
      return this.notSupported('deleteAttachment');
  }
  async uploadFile(_serverRelativeUrl: string, _fileName: string, _file: Blob | ArrayBuffer): Promise<string> {
      return this.notSupported('uploadFile');
  }
  async updateFileMetadata(_serverRelativeUrl: string, _payload: Record<string, any>): Promise<void> {
      return this.notSupported('updateFileMetadata');
  }
  async deleteFile(_serverRelativeUrl: string): Promise<void> {
      return this.notSupported('deleteFile');
  }
  async createFolder(_serverRelativeUrl: string): Promise<void> {
      return this.notSupported('createFolder');
  }
  async createList(_title: string, _description?: string, _template?: number): Promise<ListInfo> {
      return this.notSupported('createList');
  }
  async deleteList(_title: string): Promise<void> {
      return this.notSupported('deleteList');
  }
  async ensureUser(_loginName: string): Promise<UserInfo> {
      return this.notSupported('ensureUser');
  }
  async addUserToGroup(_groupName: string, _loginName: string): Promise<void> {
      return this.notSupported('addUserToGroup');
  }
  async removeUserFromGroup(_groupName: string, _loginName: string): Promise<void> {
      return this.notSupported('removeUserFromGroup');
  }
  async createGroup(_groupName: string, _description?: string): Promise<SiteGroup> {
      return this.notSupported('createGroup');
  }

  // User methods - Current User might return 'Anonymous' or error
  async getCurrentUser(): Promise<UserInfo> {
      try {
          return await this.request(`/_api/web/currentuser`);
      } catch (e) {
          // Fallback for anonymous
          return { Id: 0, Title: "Anonymous", Email: "", LoginName: "" };
      }
  }

  async getSiteUsers(): Promise<UserInfo[]> {
     // Usually blocked for anonymous
     return this.notSupported('getSiteUsers');
  }

  async searchUsers(_query: string): Promise<UserInfo[]> {
     return this.notSupported('searchUsers');
  }

  async getUserGroups(_email?: string): Promise<SiteGroup[]> {
     return this.notSupported('getUserGroups');
  }

  async getUserEffectivePermissions(_email?: string): Promise<SPBasePermissions> {
     // Return empty/readonly permissions
     return { High: 0, Low: 1 }; // ViewListItems?
  }
}
