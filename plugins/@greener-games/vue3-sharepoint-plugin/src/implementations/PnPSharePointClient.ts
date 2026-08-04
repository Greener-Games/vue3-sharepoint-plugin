import { spfi, type SPFI } from '@pnp/sp'
import { PnPLogging } from '@pnp/logging'
import { Caching } from '@pnp/queryable'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
import '@pnp/sp/files'
import '@pnp/sp/folders'
import '@pnp/sp/fields'
import '@pnp/sp/site-users/web'
import '@pnp/sp/site-groups/web'
import '@pnp/sp/batching'
import '@pnp/sp/attachments'
import { SortDirection } from '@pnp/sp/search'
import type {
  IBatch,
  ISharePointClient,
  ListItemQueryOptions,
  SearchRequestOptions,
  SearchResult,
  SharePointConfig,
  SiteGroup,
  SPBasePermissions,
  UserInfo,
  FieldDefinition,
  FileVersion,
  WebInfo,
  ListInfo,
  AttachmentInfo,
} from '../types'
import { getServerRelativePath } from '../utils/urlUtils'
import { adaptFileMetadata } from '../utils/metadataAdapter'
import { Logger } from '../utils/debug'

/** Internal interface for search cache entries */
interface SearchCacheEntry<T> {
  data: SearchResult<T>
  expiry: number
}

/** Interface for PnPjs Search Result structure to avoid 'any' */
interface PnPSearchResults {
  PrimaryQueryResult?: {
    RelevantResults?: {
      Table?: {
        Rows?: Array<{
          Cells?: Array<{ Key: string; Value: unknown }>
        }>
      }
      TotalRows?: number
    }
  }
}

/** Interface for PnPjs User structure */
interface PnPUser {
  Id: number
  Title: string
  Email: string
  LoginName: string
}

/** Interface for PnPjs List structure */
interface PnPList {
  Id: string
  Title: string
  Description: string
  ItemCount: number
  Hidden: boolean
  ImageUrl: string
}

/** Interface for PnPjs File Version structure */
interface PnPFileVersion {
  VersionLabel: string
  IsCurrentVersion: boolean
  Created: string
  Url: string
  Size: number
  CheckInComment?: string
  CreatedBy?: PnPUser
}

export class PnPSharePointClient implements ISharePointClient {
  private sp: SPFI
  private baseUrl: string
  private enableCache = false
  private logger: Logger

  constructor(options: SharePointConfig) {
    let url = options.baseUrl
    if (
      options.devBaseUrl &&
      typeof location !== 'undefined' &&
      (location.hostname === 'localhost' || location.hostname === '127.0.0.1')
    ) {
      url = options.devBaseUrl
    }

    this.baseUrl = url
    this.enableCache = !!options.enableCache
    this.logger = new Logger(options.debug)

    // 1. Initialize PnPjs with Warning logging to prevent hangs
    this.sp = spfi(this.baseUrl).using(PnPLogging(2)) // 2 = LogLevel.Warning

    // 2. Enable PnPjs Caching for Standard CRUD (if enabled)
    if (this.enableCache) {
      this.logger.log('✅ PnP: Caching Enabled (Session Storage)')
      this.sp.using(
        Caching({
          store: 'session',
          keyFactory: (u) => `${u}`,
          expireFunc: () => {
            const expiration = new Date()
            expiration.setMinutes(expiration.getMinutes() + 15)
            return expiration
          },
        })
      )
    }

    this.logger.log('PnP: Client Initialized')
  }

  /** Gets the configured base URL */
  public getBaseUrl(): string {
    return this.baseUrl
  }

  /** Performs a raw request using native fetch */
  public async request<T>(endpoint: string, options: RequestInit = {}): Promise<T> {
    const url = endpoint.startsWith('http') ? endpoint : `${this.baseUrl}${endpoint.startsWith('/') ? endpoint : `/${endpoint}`}`
    const response = await fetch(url, options)
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const contentType = response.headers.get("content-type");
    if (contentType && contentType.includes("application/json")) {
      return response.json() as Promise<T>;
    }
    return response.text() as unknown as T;
  }

  // --------------------------------------------------------------------------
  // 1. SEARCH (Native Fetch + Manual Cache)
  // --------------------------------------------------------------------------

  /** Executes a search using SharePoint Search API */
  async search<T>(
    options: SearchRequestOptions,
    _abortSignal?: AbortSignal
  ): Promise<SearchResult<T>> {
    const cacheKey = this.enableCache ? this.generateSearchKey(options) : ''

    // A. Check Cache
    if (this.enableCache) {
      const cached = this.readFromCache<SearchResult<T>>(cacheKey)
      if (cached) {
        this.logger.log('Search: Serving from Cache')
        return cached
      }
    }

    this.logger.log('Search: Starting PnPjs Request...')

    const kql = this.buildKql(options)
    this.logger.log(`Search KQL: ${kql}`)

    try {
      const userSelects = options.selectFields || []
      const userExpands = options.expandFields || []

      const needsHydration =
        userExpands.length > 0 || userSelects.some((f) => f.includes('/'))

      let searchSelect = [
        'Title',
        'Path',
        'HitHighlightedSummary',
        ...userSelects.filter((f) => !f.includes('/')),
      ]

      if (options.mapping) {
        searchSelect = [
          ...searchSelect,
          ...Object.keys(options.mapping).filter((k) => !k.includes('/')),
        ]
      }

      if (needsHydration) {
        if (!searchSelect.includes('ListId')) searchSelect.push('ListId')
        if (!searchSelect.includes('ListItemId'))
          searchSelect.push('ListItemId')
        if (!searchSelect.includes('UniqueId')) searchSelect.push('UniqueId')
      }

      // Dedupe
      searchSelect = [...new Set(searchSelect)]

      const searchResultsRaw = await this.sp.search({
        Querytext: kql,
        RowLimit: options.rowLimit || 10,
        StartRow: options.startRow || 0,
        SelectProperties: searchSelect,
        SortList: options.sortList?.map((s) => ({
          Property: s.property,
          Direction:
            s.direction === 'ascending'
              ? SortDirection.Ascending
              : SortDirection.Descending,
        })),
        TrimDuplicates: false,
        HitHighlightedProperties: ['Contents', 'Title'],
        SummaryLength: 250,
        EnableStemming: true,
        ClientType: 'ContentSearchRegular',
      })

      const searchResults = searchResultsRaw as unknown as PnPSearchResults

      this.logger.log(`Search Results:`, [searchResults])

      const relevantResults = searchResults.PrimaryQueryResult?.RelevantResults

      if (!relevantResults) {
        return { items: [], totalHits: 0, startRow: options.startRow || 0 }
      }

      const rawRows = relevantResults.Table?.Rows || []

      // E. Map (Phase 1: Raw to Flat Object)
      let items: Record<string, unknown>[] = rawRows.map((row) => {
        const map: Record<string, unknown> = {}
        if (row.Cells) {
          row.Cells.forEach((c) => (map[c.Key] = c.Value))
        } else {
          Object.assign(map, row)
        }

        if (options.includeRelativePath && typeof map.Path === 'string') {
          try {
            map.relativePath = decodeURIComponent(new URL(map.Path).pathname)
          } catch {
            map.relativePath = map.Path
          }
        }
        return map
      })

      // F. Hydration (Optional)
      if (needsHydration && items.length > 0) {
        const listSelect = userSelects.filter((f) => !this.isSearchOnlyProp(f))

        await Promise.all(
          items.map(async (item) => {
            if (typeof item.ListId === 'string' && (typeof item.ListItemId === 'string' || typeof item.ListItemId === 'number')) {
              const listId = item.ListId
              const itemId = typeof item.ListItemId === 'string' ? parseInt(item.ListItemId, 10) : item.ListItemId
              
              let q = this.sp.web.lists.getById(listId).items.getById(itemId)
              if (listSelect.length > 0) q = q.select(...listSelect)
              if (userExpands.length > 0) q = q.expand(...userExpands)

              try {
                const hydrated = await q()
                if (hydrated) {
                  Object.assign(item, hydrated)
                }
              } catch {
                /* ignore missing items */
              }
            }
          })
        )
      }

      // G. Final Mapping
      const mapping = options.mapping
      if (mapping) {
        items = items.map((map) => {
          const out: Record<string, unknown> = {}

          Object.entries(mapping).forEach(([k, v]) => {
            // 1. Get Value from Source (support dot notation e.g. "Author.Title")
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

            const destKey = v;

            // 2. Assign Value to Destination (support dot notation e.g. "owner.name")
            if (destKey.includes('.')) {
              const parts = destKey.split('.')
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
              out[destKey] = val
            }
          })

          if (!out.url && typeof map.Path === 'string') out.url = map.Path
          if (options.includeRelativePath) out.relativePath = map.relativePath
          return out
        })
      }

      const result: SearchResult<T> = {
        items: items as unknown as T[],
        totalHits: relevantResults.TotalRows || 0,
        startRow: options.startRow || 0,
      }

      // F. Cache Result
      if (this.enableCache) {
        this.writeToCache(cacheKey, result)
      }

      return result
    } catch (error: unknown) {
      this.logger.error('Search Error:', [error])
      throw error
    }
  }

  // --------------------------------------------------------------------------
  // 2. BATCHING
  // --------------------------------------------------------------------------

  /** Executes multiple SharePoint operations in a single batch */
  async executeBatch(builder: (batch: IBatch) => void, _abortSignal?: AbortSignal): Promise<void> {
    const promises: Promise<any>[] = []

    const proxy: IBatch = {
      createListItem: (list, payload) => {
        promises.push(this.sp.web.lists.getByTitle(list).items.add(payload))
      },
      updateListItem: (list, id, payload) => {
        promises.push(this.sp.web.lists.getByTitle(list).items.getById(id).update(payload))
      },
      deleteListItem: (list, id) => {
        promises.push(this.sp.web.lists.getByTitle(list).items.getById(id).recycle())
      },
      deleteFile: (url) => {
        promises.push(this.sp.web.getFileByServerRelativePath(url).recycle())
      },
    }

    builder(proxy)
    this.logger.log(`Executing Batch...`)
    await Promise.all(promises)
  }

  // --------------------------------------------------------------------------
  // 3. LIST ITEMS & ATTACHMENTS
  // --------------------------------------------------------------------------

  /** Creates a new item in a list */
  async createListItem<T>(
    listTitle: string,
    payload: Record<string, unknown>,
    _abortSignal?: AbortSignal
  ): Promise<T> {
    const r = await this.sp.web.lists.getByTitle(listTitle).items.add(payload)
    // eslint-disable-next-line @typescript-eslint/no-unnecessary-type-assertion
    return r.data as unknown as T
  }

  /** Updates an existing item in a list */
  async updateListItem(
    listTitle: string,
    id: number,
    payload: Record<string, unknown>,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(id)
      .update(payload)
  }

  /** Recycles a list item */
  async deleteListItem(listTitle: string, id: number, _abortSignal?: AbortSignal): Promise<void> {
    await this.sp.web.lists.getByTitle(listTitle).items.getById(id).recycle()
  }

  /** Retrieves a single list item by its ID */
  async getListItemById<T>(
    listTitle: string,
    id: number,
    select?: string[],
    expand?: string[],
    _abortSignal?: AbortSignal
  ): Promise<T> {
    let q = this.sp.web.lists.getByTitle(listTitle).items.getById(id)
    if (select && select.length > 0) {
      q = q.select(...select)
    }
    if (expand && expand.length > 0) {
      q = q.expand(...expand)
    }
    const result = (await q()) as unknown as T
    return result
  }

  /** Retrieves multiple items from a list */
  async getListItems<T>(
    listTitle: string,
    options?: ListItemQueryOptions,
    _abortSignal?: AbortSignal
  ): Promise<T[]> {
    let q = this.sp.web.lists.getByTitle(listTitle).items

    if (options?.select && options.select.length > 0) {
      q = q.select(...options.select)
    }
    if (options?.expand && options.expand.length > 0) {
      q = q.expand(...options.expand)
    }
    if (options?.filter) {
      q = q.filter(options.filter)
    }
    if (options?.top) {
      q = q.top(options.top)
    }
    if (options?.orderBy) {
      q = q.orderBy(options.orderBy, options.ascending ?? true)
    }

    if (options?.getAll) {
      let results: T[] = []
      // Use getPaged() for reliable paging
      let paged = await q.getPaged<T>()
      results = results.concat(paged.results)
      if (options?.onProgress) {
        options.onProgress(paged.results as unknown as Record<string, unknown>[])
      }
      while (paged.hasNext) {
          const nextPaged = await paged.getNext()
          if (!nextPaged) break
          paged = nextPaged
          results = results.concat(paged.results)
          if (options?.onProgress) {
            options.onProgress(paged.results as unknown as Record<string, unknown>[])
          }
      }
      return results
    }

    const results = (await q()) as unknown as T[]
    return results
  }

  /** Gets attachment files for a list item */
  async getItemAttachments(
    listTitle: string,
    itemId: number,
    _abortSignal?: AbortSignal
  ): Promise<AttachmentInfo[]> {
    const attachments = (await this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .attachmentFiles()) as unknown as Array<{ FileName: string, ServerRelativeUrl: string }>
    return attachments.map((a) => ({
      FileName: a.FileName,
      ServerRelativeUrl: a.ServerRelativeUrl,
    }))
  }

  /** Adds an attachment to a list item */
  async addAttachment(
    listTitle: string,
    itemId: number,
    fileName: string,
    file: Blob | ArrayBuffer,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .attachmentFiles.add(fileName, file)
  }

  /** Deletes an attachment from a list item */
  async deleteAttachment(
    listTitle: string,
    itemId: number,
    fileName: string,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .attachmentFiles.getByName(fileName)
      .delete()
  }

  // --------------------------------------------------------------------------
  // 4. FILES & FOLDERS
  // --------------------------------------------------------------------------

  /** Uploads a file to a specific folder */
  async uploadFile(
    serverRelativeUrl: string,
    fileName: string,
    file: Blob | ArrayBuffer,
    _abortSignal?: AbortSignal
  ): Promise<string> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const r = await this.sp.web
      .getFolderByServerRelativePath(fullUrl)
      .files.addUsingPath(fileName, file, { Overwrite: true })

    const data = r.data as unknown as { ServerRelativeUrl: string }
    return data.ServerRelativeUrl
  }

  /** Downloads a file as a Blob */
  async downloadFile(serverRelativeUrl: string, _abortSignal?: AbortSignal): Promise<Blob> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    return await this.sp.web.getFileByServerRelativePath(fullUrl).getBlob()
  }

  /** Updates metadata for a file */
  async updateFileMetadata(
    serverRelativeUrl: string,
    payload: Record<string, unknown>,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const file = this.sp.web.getFileByServerRelativePath(fullUrl)

    const itemInfo = (await file.listItemAllFields
      .select('Id', 'ParentList/Title')
      .expand('ParentList')()) as unknown as { Id: number; ParentList?: { Title: string } }

    if (!itemInfo || !itemInfo.ParentList || typeof itemInfo.ParentList.Title !== 'string') {
      throw new Error(`Could not determine list title for file: ${fullUrl}`)
    }

    const listTitle = itemInfo.ParentList.Title

    const formValues = await adaptFileMetadata(this, listTitle, payload)

    const item = file.listItemAllFields
    // Use validateUpdateListItem to bypass some permission issues and handle complex fields
    const validator = item as unknown as {
      validateUpdateListItem: (vals: unknown) => Promise<
        Array<{ ErrorCode: number; FieldName: string; ErrorMessage: string }>
      >
    }
    const results = await validator.validateUpdateListItem(formValues)
    const errors = results.filter((r) => r.ErrorCode !== 0)
    if (errors.length > 0) {
      const msg = errors.map((e) => `${e.FieldName}: ${e.ErrorMessage}`).join('; ')
      throw new Error(`UpdateFileMetadata failed: ${msg}`)
    }
  }

  /** Recycles a file */
  async deleteFile(serverRelativeUrl: string, _abortSignal?: AbortSignal): Promise<void> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    await this.sp.web.getFileByServerRelativePath(fullUrl).recycle()
  }

  /** Creates a folder at the specified path */
  async createFolder(serverRelativeUrl: string, _abortSignal?: AbortSignal): Promise<void> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    await this.sp.web.folders.addUsingPath(fullUrl)
  }

  // --------------------------------------------------------------------------
  // 5. WEBS & LISTS (STRUCTURE)
  // --------------------------------------------------------------------------

  /** Gets information about the current web */
  async getWebInfo(_abortSignal?: AbortSignal): Promise<WebInfo> {
    const w = (await this.sp.web.select('Id', 'Title', 'Url', 'Description')()) as unknown as WebInfo
    return {
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }
  }

  /** Gets subwebs of the current web */
  async getSubwebs(_abortSignal?: AbortSignal): Promise<WebInfo[]> {
    const webs = (await this.sp.web.webs.select('Id', 'Title', 'Url', 'Description')()) as unknown as WebInfo[]
    return webs.map((w) => ({
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description,
    }))
  }

  /** Gets all lists in the current web */
  async getLists(_abortSignal?: AbortSignal): Promise<ListInfo[]> {
    const lists = (await this.sp.web.lists.select('Id', 'Title', 'Description', 'ItemCount', 'Hidden', 'ImageUrl')()) as unknown as PnPList[]
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
  async getList(listTitle: string, _abortSignal?: AbortSignal): Promise<ListInfo> {
    const l = (await this.sp.web.lists.getByTitle(listTitle).select('Id', 'Title', 'Description', 'ItemCount', 'Hidden', 'ImageUrl')()) as unknown as PnPList
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
    _abortSignal?: AbortSignal
  ): Promise<ListInfo> {
    const r = await this.sp.web.lists.add(title, description, template)
    const l = r.data as unknown as PnPList
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
  async deleteList(title: string, _abortSignal?: AbortSignal): Promise<void> {
    await this.sp.web.lists.getByTitle(title).delete()
  }

  // --------------------------------------------------------------------------
  // 6. USERS & GROUPS
  // --------------------------------------------------------------------------

  /** Gets the currently logged in user */
  async getCurrentUser(_abortSignal?: AbortSignal): Promise<UserInfo> {
    const user = (await this.sp.web.currentUser()) as unknown as PnPUser
    return {
        Id: user.Id,
        Title: user.Title,
        Email: user.Email,
        LoginName: user.LoginName
    }
  }

  /** Gets all users on the site */
  async getSiteUsers(_abortSignal?: AbortSignal): Promise<UserInfo[]> {
    const users = (await this.sp.web.siteUsers()) as unknown as PnPUser[]
    return users.map(user => ({
        Id: user.Id,
        Title: user.Title,
        Email: user.Email,
        LoginName: user.LoginName
    }))
  }

  /** Searches for users by title or email */
  async searchUsers(query: string, _abortSignal?: AbortSignal): Promise<UserInfo[]> {
    if (!query) return []

    const users = (await this.sp.web.siteUsers.filter(
      `substringof('${query}', Title) or substringof('${query}', Email)`
    )()) as unknown as PnPUser[]
    return users.map(user => ({
        Id: user.Id,
        Title: user.Title,
        Email: user.Email,
        LoginName: user.LoginName
    }))
  }

  /** Ensures a user exists on the site */
  async ensureUser(loginName: string, _abortSignal?: AbortSignal): Promise<UserInfo> {
    const result = await this.sp.web.ensureUser(loginName)
    const user = result.data as unknown as PnPUser
    return {
        Id: user.Id,
        Title: user.Title,
        Email: user.Email,
        LoginName: user.LoginName
    }
  }

  /** Gets groups for a specific user or current user */
  async getUserGroups(email?: string, _abortSignal?: AbortSignal): Promise<SiteGroup[]> {
    let groups: Array<{ Id: number, Title: string, Description: string, LoginName: string }>;
    if (email) {
      groups = await this.sp.web.siteUsers.getByEmail(email).groups()
    } else {
      groups = await this.sp.web.currentUser.groups()
    }
    return groups.map(g => ({
        Id: g.Id,
        Title: g.Title,
        Description: g.Description,
        LoginName: g.LoginName
    }))
  }

  /** Adds a user to a SharePoint group */
  async addUserToGroup(groupName: string, loginName: string, _abortSignal?: AbortSignal): Promise<void> {
    const userResult = await this.sp.web.ensureUser(loginName)
    const user = userResult.data as unknown as PnPUser
    await this.sp.web.siteGroups
      .getByName(groupName)
      .users.add(user.LoginName)
  }

  /** Removes a user from a SharePoint group */
  async removeUserFromGroup(
    groupName: string,
    loginName: string,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    await this.sp.web.siteGroups
      .getByName(groupName)
      .users.removeByLoginName(loginName)
  }

  /** Creates a new SharePoint group */
  async createGroup(
    groupName: string,
    description?: string,
    _abortSignal?: AbortSignal
  ): Promise<SiteGroup> {
    const r = await this.sp.web.siteGroups.add({
      Title: groupName,
      Description: description,
    })
    const g = r.data as unknown as PnPUser & { Description: string }
    return {
        Id: g.Id,
        Title: g.Title,
        Description: g.Description,
        LoginName: g.LoginName
    }
  }

  /** Gets effective permissions for a user */
  async getUserEffectivePermissions(
    email?: string,
    _abortSignal?: AbortSignal
  ): Promise<SPBasePermissions> {
    if (email) {
      const user = (await this.sp.web.siteUsers.getByEmail(email)()) as unknown as PnPUser
      return await this.sp.web.getUserEffectivePermissions(user.LoginName)
    }
    return await this.sp.web.getCurrentUserEffectivePermissions()
  }

  // --------------------------------------------------------------------------
  // 7. FIELDS & METADATA
  // --------------------------------------------------------------------------

  /** Gets all visible fields for a list */
  async getListFields(
    listTitle: string,
    _webUrl?: string,
    _abortSignal?: AbortSignal
  ): Promise<FieldDefinition[]> {
    const r = (await this.sp.web.lists
      .getByTitle(listTitle)
      .fields.filter('Hidden eq false')()) as unknown as Array<{
      InternalName: string
      Title: string
      TypeAsString: string
      Hidden: boolean
      Choices?: string[]
      TermSetId?: string
    }>

    return r.map((f) => ({
      InternalName: f.InternalName,
      Title: f.Title,
      TypeAsString: f.TypeAsString,
      Hidden: f.Hidden,
      Choices: f.Choices || [],
      TermSetId: f.TermSetId || undefined,
    }))
  }

  /** Gets choices for a Choice field */
  async getFieldChoices(
    listTitle: string,
    fieldInternalName: string,
    _abortSignal?: AbortSignal
  ): Promise<string[]> {
    const f = (await this.sp.web.lists
      .getByTitle(listTitle)
      .fields.getByInternalNameOrTitle(fieldInternalName)
      .select('Choices')()) as unknown as { Choices?: string[] }
    return f.Choices || []
  }

  /** Searches for a taxonomy term by label */
  async searchTerm(
    termSetId: string,
    label: string,
    abortSignal?: AbortSignal
  ): Promise<{ Label: string; TermGuid: string } | null> {
    try {
      const safeLabel = label.replace(/'/g, "''")
      const endpoint = `${this.baseUrl}/_api/v2.1/termStore/termSets/${termSetId}/terms?$filter=labels/any(l:l/name eq '${safeLabel}') or name eq '${safeLabel}'&$select=id,name`

      const res = await fetch(endpoint, {
        headers: {
            'Accept': 'application/json;odata=verbose'
        },
        signal: abortSignal
      })

      if (!res.ok) return null

      const data = (await res.json()) as { value: Array<{ id: string, name: string, names?: Array<{ name: string }> }> }
      const terms = data.value

      if (terms && terms.length > 0) {
        const t = terms[0]
        return {
             Label: t.names?.[0]?.name || t.name || label,
             TermGuid: t.id
        }
      }
    } catch (e: unknown) {
      this.logger.warn(`SearchTerm failed (PnP Client):`, [e])
    }
    return null
  }

  // --------------------------------------------------------------------------
  // 8. VERSION HISTORY
  // --------------------------------------------------------------------------

  /** Gets version history for a file */
  async getFileVersions(
    serverRelativeUrl: string,
    _abortSignal?: AbortSignal
  ): Promise<FileVersion[]> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const versions = (await this.sp.web
      .getFileByServerRelativePath(fullUrl)
      .versions.expand('CreatedBy')()) as unknown as PnPFileVersion[]

    return versions.map((v) => ({
      VersionLabel: v.VersionLabel,
      Created: v.Created,
      CheckInComment: v.CheckInComment || '',
      IsCurrentVersion: v.IsCurrentVersion,
      Size: v.Size,
      Url: v.Url,
      CreatedBy: {
        Id: v.CreatedBy?.Id,
        Title: v.CreatedBy?.Title,
        Email: v.CreatedBy?.Email,
      },
    }))
  }

  /** Gets a link to the version history page for a file */
  async getVersionHistoryLink(
    serverRelativeUrl: string,
    _abortSignal?: AbortSignal
  ): Promise<string> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const file = this.sp.web.getFileByServerRelativePath(fullUrl)
    const item = (await file.listItemAllFields
      .select('Id', 'ParentList/Id')
      .expand('ParentList')()) as unknown as { Id: number; ParentList?: { Id: string } }

    if (!item || !item.ParentList) {
      throw new Error(`Could not resolve List ID for file: ${fullUrl}`)
    }

    const listId = String(item.ParentList.Id)
    const itemId = item.Id

    return this.getVersionHistoryLinkByItem(listId, itemId)
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

  // --------------------------------------------------------------------------
  // PRIVATE HELPERS
  // --------------------------------------------------------------------------

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
          return `Path:"${origin}${safeS}*"`
        }
        return `Path:"${this.baseUrl}/${safeS}*"`
      })
      parts.push(`(${normalizedScopes.join(' OR ')})`)
    }

    if (opts.resultType === 'items') {
      parts.push(`(NOT ContentTypeId:0x0120*)`)
    } else if (opts.resultType === 'folders') {
      parts.push(`ContentTypeId:0x0120*`)
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
        const formatValue = (v: unknown) => {
          if (typeof v === 'object' && v !== null && !Array.isArray(v)) {
            const range = v as Record<string, unknown>
            const rangeParts: string[] = []
            if (range.gte !== undefined) rangeParts.push(`${mp}>=${range.gte}`)
            if (range.lte !== undefined) rangeParts.push(`${mp}<=${range.lte}`)
            if (range.gt !== undefined) rangeParts.push(`${mp}>${range.gt}`)
            if (range.lt !== undefined) rangeParts.push(`${mp}<${range.lt}`)
            if (rangeParts.length > 1) {
              return `(${rangeParts.join(' AND ')})`
            } else if (rangeParts.length === 1) {
              return rangeParts[0]
            }
            return ''
          }

          const val = String(v)
          if (/^(?:>=|<=|<|>|=)/.test(val)) {
            return `${mp}${val}`
          }
          return `${mp}:"${val}"`
        }

        if (Array.isArray(value)) {
          const formatted = value.map((v) => formatValue(v)).filter(Boolean)
          if (formatted.length > 0) {
            parts.push(`(${formatted.join(' OR ')})`)
          }
        } else {
          const formatted = formatValue(value)
          if (formatted) parts.push(formatted)
        }
      }
    }
    return parts.join(' AND ')
  }

  private generateSearchKey(opts: SearchRequestOptions): string {
    return `SP_SEARCH_${this.hashCode(JSON.stringify(opts))}`
  }

  private readFromCache<T>(key: string): T | null {
    try {
      const item = sessionStorage.getItem(key)
      if (!item) return null
      const parsed = JSON.parse(item) as SearchCacheEntry<unknown>
      if (Date.now() > parsed.expiry) {
        sessionStorage.removeItem(key)
        return null
      }
      return parsed.data as T
    } catch {
      return null
    }
  }

  private writeToCache<T>(key: string, data: T) {
    try {
      const payload = {
        data,
        expiry: Date.now() + 15 * 60 * 1000, // 15 Mins
      }
      sessionStorage.setItem(key, JSON.stringify(payload))
    } catch (_e: unknown) {
      /* ignore storage errors */
    }
  }

  private hashCode(str: string): number {
    let hash = 0
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i)
      hash = (hash << 5) - hash + char
      hash |= 0
    }
    return hash
  }

  private isSearchOnlyProp(field: string): boolean {
    const searchOnlyPattern =
      /^(Refinable|HitHighlighted|Path$|OriginalPath$|Rank$|DocId$|WorkId$|Piw)/i
    return searchOnlyPattern.test(field)
  }
}
