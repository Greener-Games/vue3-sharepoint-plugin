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

/** Internal interface for cached items */
interface CacheRecord<T> {
  data: T
  expiry: number
}

// Internal Cache Helper (Reused from RestSharePointClient logic)
class InternalCache {
  private enabled: boolean
  private prefix = 'sp-simple-cache::'

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
    ListItemEntityTypeFullName?: string
  }
}

/** Interface for SharePoint OData V2.1 / REST responses */
interface SPODataResponse<T> {
  value?: T
  '@odata.nextLink'?: string
  'odata.nextLink'?: string
}

export class AnonymousClient implements ISharePointClient {
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

  /** Gets the configured base URL */
  public getBaseUrl(): string {
    return this.baseUrl
  }

  // --- Centralized Request Handler ---
  /** Performs a SharePoint REST API request */
  public async request<T>(
    endpoint: string,
    options: Omit<RequestInit, 'body'> & {
      targetPath?: string
      isWrite?: boolean
      skipMetadata?: boolean
      body?: unknown
    } = {}
  ): Promise<T> {
    const {
      method = 'GET',
      targetPath,
      isWrite = false,
      body,
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
      if (options.headers instanceof Headers) {
          options.headers.forEach((v, k) => headers.set(k, v))
      } else if (Array.isArray(options.headers)) {
          options.headers.forEach(([k, v]) => headers.set(k, v))
      } else {
          Object.entries(options.headers).forEach(([k, v]) => headers.set(k, v))
      }
    }

    this.logger.log(`SimpleRest Request: ${method} ${finalEndpoint}`)

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
      return null as unknown as T
    }

    const data = (await response.json()) as SPVerboseResponse<T> & SPODataResponse<T>
    this.logger.log(`Response: ${response.status}`, data)

    if (options.skipMetadata) {
      return data as unknown as T
    }

    // Robustly check for nested 'd' and 'results'
    if (data && data.d) {
        const d = data.d;
        if (d && typeof d === 'object' && 'results' in d && d.results) {
            return d.results as unknown as T;
        }
        return d as unknown as T;
    }

    return data as unknown as T
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
        const d = (await res.json()) as SPVerboseResponse<{ GetContextWebInformation: { FormDigestValue: string, FormDigestTimeoutSeconds: number } }>
        const contextInfo = d.d?.GetContextWebInformation
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
    } catch (e: unknown) {
        this.logger.warn("Failed to fetch Request Digest", e);
        return '';
    }
  }

  private async getListEntityType(listTitle: string): Promise<string> {
      // Fetch list metadata to get ListItemEntityTypeFullName
      // Endpoint: /_api/web/lists/getbytitle('List')?$select=ListItemEntityTypeFullName
      try {
          const list = await this.request<{ ListItemEntityTypeFullName: string }>(`/_api/web/lists/getbytitle('${listTitle}')?$select=ListItemEntityTypeFullName`);
          return list.ListItemEntityTypeFullName || 'SP.Data.ListItem'; // Fallback
      } catch {
          return 'SP.Data.ListItem';
      }
  }

  // --- Search Implementation (via Lists) ---
  async search<T>(opts: SearchRequestOptions, signal?: AbortSignal): Promise<SearchResult<T>> {
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
        const results = await this.request<Array<Record<string, unknown>>>(endpoint, { targetPath: scopePath, signal });

        // Map to SearchResult
        const items = results.map(item => {
             const out = { ...item } as Record<string, unknown>;
             out.Path = (item.EncodedAbsUrl as string) || this.baseUrl + (item.FileRef as string);
             if (opts.includeRelativePath) {
                 out.relativePath = item.FileRef;
             }
             if (opts.mapping) {
                Object.entries(opts.mapping).forEach(([k, v]) => {
                    if (item[k] !== undefined) out[v] = item[k];
                })
             }
             return out as unknown as T;
        });

        return {
            items,
            totalHits: items.length,
            startRow: 0
        }

    } catch (e: unknown) {
        this.logger.error("SimpleRest Search Failed", e);
        return { items: [], totalHits: 0, startRow: 0 }
    }
  }

  // --- Read Methods (Supported) ---

  async getWebInfo(signal?: AbortSignal): Promise<WebInfo> {
    const w = await this.request<WebInfo>(`/_api/web`, { signal })
    return {
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }
  }

  async getSubwebs(signal?: AbortSignal): Promise<WebInfo[]> {
    const webs = await this.request<WebInfo[]>(`/_api/web/webs`, { signal })
    return webs.map((w) => ({
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }))
  }

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

      while (nextUrl) {
        const response = await this.request<SPVerboseResponse<T[]> & SPODataResponse<T[]>>(nextUrl, {
          skipMetadata: true,
          signal
        })

        if (response && response.d) {
          const results = response.d.results || (response.d as unknown as T[])
          if (Array.isArray(results)) {
            allItems = allItems.concat(results)
            if (options?.onProgress) {
              options.onProgress(results as unknown as Record<string, unknown>[])
            }
          } else {
            allItems.push(results as unknown as T)
            if (options?.onProgress) {
              options.onProgress([results as unknown as Record<string, unknown>])
            }
          }
          nextUrl = response.d.__next
        } else if (response && response.value) {
          allItems = allItems.concat(response.value)
          if (options?.onProgress) {
            options.onProgress(response.value as unknown as Record<string, unknown>[])
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

  async downloadFile(url: string, signal?: AbortSignal): Promise<Blob> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    const apiRoot = this.getApiRoot(fullUrl)

    const headers = new Headers({
        Accept: 'application/json;odata=verbose'
    })

    const endpoint = `${apiRoot}/_api/web/getfilebyserverrelativeurl('${fullUrl}')/$value`;

    this.logger.log(`Anon Download: ${endpoint}`);
    const res = await fetch(endpoint, { headers, signal: signal });
    if (!res.ok) throw new Error(`Download failed: ${res.statusText}`);
    return await res.blob()
  }

  async getFileVersions(url: string, signal?: AbortSignal): Promise<FileVersion[]> {
    const fullUrl = getServerRelativePath(url, this.baseUrl)
    const endpoint = `/_api/web/getfilebyserverrelativeurl('${fullUrl}')/versions?$expand=CreatedBy`
    const results = await this.request<Array<Record<string, unknown>>>(endpoint, { targetPath: fullUrl, signal })

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

  async getVersionHistoryLink(serverRelativeUrl: string, signal?: AbortSignal): Promise<string> {
    try {
        const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
        const safeUrl = fullUrl.replace(/'/g, "''")

        const data = await this.request<{ Id: number, ParentList: { Id: string } }>(
            `/_api/web/getFileByServerRelativeUrl('${safeUrl}')/ListItemAllFields?$select=Id,ParentList/Id&$expand=ParentList`,
            {
                method: 'GET',
                targetPath: fullUrl,
                signal
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
      this.logger.warn(`SearchTerm failed for ${label}`, e)
    }
    return null
  }

  // --- Write Operations (Partial Support) ---

  async createListItem<T>(listTitle: string, payload: Record<string, unknown>, signal?: AbortSignal): Promise<T> {
      // 1. Get Entity Type (required for verbose JSON)
      const type = await this.getListEntityType(listTitle);

      // 2. Prepare Payload
      const body = {
          __metadata: { type },
          ...payload
      };

      // 3. Post
      return await this.request<T>(`/_api/web/lists/getbytitle('${listTitle}')/items`, {
          method: 'POST',
          body,
          isWrite: true,
          signal
      });
  }

  async updateListItem(listTitle: string, id: number, payload: Record<string, unknown>, signal?: AbortSignal): Promise<void> {
      const type = await this.getListEntityType(listTitle);

      const body = {
          __metadata: { type },
          ...payload
      };

      await this.request<void>(`/_api/web/lists/getbytitle('${listTitle}')/items(${id})`, {
          method: 'POST',
          body,
          headers: {
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*', // Force update, assumes concurrency is not handled
          },
          isWrite: true,
          signal
      });
  }

  async deleteListItem(listTitle: string, id: number, signal?: AbortSignal): Promise<void> {
      await this.request<void>(`/_api/web/lists/getbytitle('${listTitle}')/items(${id})`, {
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*',
        },
        isWrite: true,
        signal
      })
  }

  // --- Unsupported / Stubbed Operations ---

  private async notSupported(operation: string): Promise<never> {
      this.logger.error(`Operation '${operation}' is not supported in SimpleRest Mode.`)
      return Promise.reject(new Error(`Operation '${operation}' is not supported in SimpleRest Mode.`));
  }

  async executeBatch(_builder: (batch: IBatch) => void, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('executeBatch');
  }
  async addAttachment(_listTitle: string, _itemId: number, _fileName: string, _file: Blob | ArrayBuffer, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('addAttachment');
  }
  async deleteAttachment(_listTitle: string, _itemId: number, _fileName: string, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('deleteAttachment');
  }
  async uploadFile(_serverRelativeUrl: string, _fileName: string, _file: Blob | ArrayBuffer, _signal?: AbortSignal): Promise<string> {
      return await this.notSupported('uploadFile');
  }
  async updateFileMetadata(_serverRelativeUrl: string, _payload: Record<string, unknown>, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('updateFileMetadata');
  }
  async deleteFile(_serverRelativeUrl: string, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('deleteFile');
  }
  async createFolder(_serverRelativeUrl: string, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('createFolder');
  }
  async createList(_title: string, _description?: string, _template?: number, _signal?: AbortSignal): Promise<ListInfo> {
      return await this.notSupported('createList');
  }
  async deleteList(_title: string, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('deleteList');
  }
  async ensureUser(_loginName: string, _signal?: AbortSignal): Promise<UserInfo> {
      return await this.notSupported('ensureUser');
  }
  async addUserToGroup(_groupName: string, _loginName: string, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('addUserToGroup');
  }
  async removeUserFromGroup(_groupName: string, _loginName: string, _signal?: AbortSignal): Promise<void> {
      await this.notSupported('removeUserFromGroup');
  }
  async createGroup(_groupName: string, _description?: string, _signal?: AbortSignal): Promise<SiteGroup> {
      return await this.notSupported('createGroup');
  }

  // User methods - Current User might return 'Anonymous' or error
  async getCurrentUser(signal?: AbortSignal): Promise<UserInfo> {
      try {
          return await this.request<UserInfo>(`/_api/web/currentuser`, { signal });
      } catch (_e: unknown) {
          // Fallback for anonymous
          return { Id: 0, Title: "Anonymous", Email: "", LoginName: "" };
      }
  }

  async getSiteUsers(signal?: AbortSignal): Promise<UserInfo[]> {
     // Usually blocked for anonymous
     return await this.notSupported('getSiteUsers');
  }

  async searchUsers(_query: string, _signal?: AbortSignal): Promise<UserInfo[]> {
     return await this.notSupported('searchUsers');
  }

  async getUserGroups(_email?: string, _signal?: AbortSignal): Promise<SiteGroup[]> {
     return await this.notSupported('getUserGroups');
  }

  async getUserEffectivePermissions(_email?: string, _signal?: AbortSignal): Promise<SPBasePermissions> {


     // Return empty/readonly permissions
     return Promise.resolve({ High: 0, Low: 1 }); // ViewListItems?
  }


}
