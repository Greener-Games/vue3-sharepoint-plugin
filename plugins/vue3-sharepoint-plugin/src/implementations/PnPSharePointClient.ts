import type {
  IBatch,
  ISharePointClient,
  ListItemQueryOptions,
  AdvancedSearchOptions,
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
import { Logger } from '../utils/debug'
import { spfi, SPFI } from '@pnp/sp'
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
import '@pnp/sp/search'

export class PnPSharePointClient implements ISharePointClient {
  private sp: SPFI
  private baseUrl: string
  private enableCache: boolean = false
  private logger: Logger

  constructor(options: SharePointConfig) {
    this.baseUrl = options.baseUrl
    this.enableCache = !!options.enableCache
    this.logger = new Logger(options.debug)

    // 1. Initialize PnPjs with Warning logging to prevent hangs
    // If debug is on, we might want to increase log level, but kept to Warning to avoid spam unless PnPLogging supports it well.
    // For now we use our own logger.
    this.sp = spfi(options.baseUrl).using(PnPLogging(2)) // 2 = LogLevel.Warning

    // 2. Enable PnPjs Caching for Standard CRUD (if enabled)
    // Note: Search uses its own manual cache below
    if (this.enableCache) {
      console.log('✅ PnP: Caching Enabled (Session Storage)')
      this.sp.using(
          Caching({
            store: 'session',
            keyFactory: (url) => `${url}`,
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

  // --------------------------------------------------------------------------
  // 1. SEARCH (Native Fetch + Manual Cache)
  // --------------------------------------------------------------------------

  async search<T = any>(
      options: SearchRequestOptions
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
      // Normalize selectFields
      let select: string[] = []
      let hydration: AdvancedSearchOptions | undefined

      if (Array.isArray(options.selectFields)) {
        select = options.selectFields
      } else if (options.selectFields) {
        // Object format -> Advanced/Hydration
        hydration = options.selectFields
        select = hydration.searchSelect || [
          'Title',
          'Path',
          'HitHighlightedSummary'
        ]
      } else {
        // Default
        select = ['Title', 'Path', 'HitHighlightedSummary']
      }

      // Add mapping keys to search selection
      if (options.mapping) {
        select = [...new Set([...select, ...Object.keys(options.mapping)])]
      }

      // Ensure we fetch identifiers if hydration is requested
      if (hydration) {
        if (!select.includes('ListId')) select.push('ListId')
        if (!select.includes('ListItemId')) select.push('ListItemId')
        if (!select.includes('UniqueId')) select.push('UniqueId')
      }

      const searchResults = await this.sp.search({
        Querytext: kql,
        RowLimit: options.rowLimit || 10,
        StartRow: options.startRow || 0,
        SelectProperties: select,
        TrimDuplicates: false,
        // Explicitly requesting highlights for content
        HitHighlightedProperties: ['Contents', 'Title'],
        SummaryLength: 250,
        EnableStemming: true,
        ClientType: 'ContentSearchRegular',
      })

      this.logger.log(`Search Results:`, searchResults)

      // @ts-ignore - PnPjs types might be slightly off for search results structure in strict mode
      const relevantResults = searchResults.PrimaryQueryResult?.RelevantResults || searchResults.Raw?.PrimaryQueryResult?.RelevantResults

      if (!relevantResults) {
        return { items: [], totalHits: 0, startRow: options.startRow || 0 }
      }

      const rawRows = relevantResults.Table?.Rows || []

      // E. Map (Phase 1: Raw to Flat Object)
      let items = rawRows.map((row: any) => {
        const map: any = {}
        if (row.Cells) {
          row.Cells.forEach((c: any) => (map[c.Key] = c.Value))
        } else {
          Object.assign(map, row)
        }

        if (options.includeRelativePath && map.Path) {
          try {
            map.relativePath = decodeURIComponent(new URL(map.Path).pathname)
          } catch {
            map.relativePath = map.Path
          }
        }
        return map
      })

      // F. Hydration (Optional)
      if (hydration && items.length > 0) {
        // We need to fetch the actual list item for each result
        // We can do this efficiently by batching requests
        const [batchedWeb, execute] = this.sp.web.batched()
        const hydrationMap = new Map<string, any>() // Key: ListId:ListItemId (or UniqueId), Value: Item

        items.forEach((item: any) => {
          // Identify Item
          // Strategy 1: ListId + ListItemId
          if (item.ListId && item.ListItemId) {
             let q = batchedWeb.lists.getById(item.ListId).items.getById(parseInt(item.ListItemId, 10))
             if (hydration!.select) q = q.select(...hydration!.select)
             if (hydration!.expand) q = q.expand(...hydration!.expand)

             q().then(r => {
               // Store result mapped by ID
               hydrationMap.set(`${item.ListId}:${item.ListItemId}`, r)
             }).catch(() => { /* ignore missing items */ })
          }
          // Strategy 2: UniqueId (DocId) -> GetFileById? But that is for Files.
          // ListId+ListItemId is safer for general items if available.
        })

        await execute()

        // Merge Hydrated Data
        items = items.map((item: any) => {
           if (item.ListId && item.ListItemId) {
             const hydrated = hydrationMap.get(`${item.ListId}:${item.ListItemId}`)
             if (hydrated) {
               return { ...item, ...hydrated } // Hydrated data overrides search data (e.g. Author string -> Author object)
             }
           }
           return item
        })
      }

      // G. Final Mapping
      if (options.mapping) {
        items = items.map((map: any) => {
          const out: any = {}

          Object.entries(options.mapping!).forEach(([k, v]) => {
            // 1. Get Value from Source (support dot notation e.g. "Author.Title")
            let val = map
            if (k.includes('.')) {
              const parts = k.split('.')
              for (const p of parts) {
                val = val ? val[p] : null
              }
            } else {
              val = map[k]
            }

            // 2. Assign Value to Destination (support dot notation e.g. "owner.name")
            if (v.includes('.')) {
              const parts = v.split('.')
              let current = out
              for (let i = 0; i < parts.length - 1; i++) {
                const part = parts[i]
                if (!current[part]) current[part] = {}
                current = current[part]
              }
              current[parts[parts.length - 1]] = val
            } else {
              out[v] = val
            }
          })

          if (!out.url) out.url = map.Path
          if (options.includeRelativePath) out.relativePath = map.relativePath
          return out
        })
      }

      const result: SearchResult<T> = {
        items: items as T[],
        totalHits: relevantResults.TotalRows || 0,
        startRow: options.startRow || 0,
      }

      // F. Cache Result
      if (this.enableCache) {
        this.writeToCache(cacheKey, result)
      }

      return result
    } catch (error) {
      this.logger.error('Search Error:', error)
      throw error
    }
  }

  // --------------------------------------------------------------------------
  // 2. BATCHING
  // --------------------------------------------------------------------------

  async executeBatch(builder: (batch: IBatch) => void): Promise<void> {
    const [batchedWeb, execute] = this.sp.web.batched()

    const proxy: IBatch = {
      createListItem: (list, payload) =>
          batchedWeb.lists.getByTitle(list).items.add(payload),
      updateListItem: (list, id, payload) =>
          batchedWeb.lists.getByTitle(list).items.getById(id).update(payload),
      deleteListItem: (list, id) =>
          batchedWeb.lists.getByTitle(list).items.getById(id).recycle(),
      deleteFile: (url) =>
          batchedWeb.getFileByServerRelativePath(url).recycle(),
    }

    builder(proxy)
    this.logger.log(`Executing Batch...`)
    await execute()
  }

  // --------------------------------------------------------------------------
  // 3. LIST ITEMS & ATTACHMENTS
  // --------------------------------------------------------------------------

  async createListItem<T = any>(
      listTitle: string,
      payload: Record<string, any>
  ): Promise<T> {
    const r = await this.sp.web.lists.getByTitle(listTitle).items.add(payload)
    return r.data as T
  }

  async updateListItem(
      listTitle: string,
      id: number,
      payload: Record<string, any>
  ): Promise<void> {
    await this.sp.web.lists
        .getByTitle(listTitle)
        .items.getById(id)
        .update(payload)
  }

  async deleteListItem(listTitle: string, id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(listTitle).items.getById(id).recycle()
  }

  async getListItemById<T = any>(
      listTitle: string,
      id: number,
    select?: string[],
    expand?: string[]
  ): Promise<T> {
    let q = this.sp.web.lists.getByTitle(listTitle).items.getById(id)
    if (select && select.length > 0) {
      q = q.select(...select)
    }
    if (expand && expand.length > 0) {
      q = q.expand(...expand)
    }
    return (await q()) as T
  }

  async getListItems<T = any>(
    listTitle: string,
    options?: ListItemQueryOptions
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

    return (await q()) as T[]
  }

  async getItemAttachments(listTitle: string, itemId: number): Promise<AttachmentInfo[]> {
    const attachments = await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).attachmentFiles()
    return attachments.map(a => ({
      FileName: a.FileName,
      ServerRelativeUrl: a.ServerRelativeUrl
    }))
  }

  async addAttachment(listTitle: string, itemId: number, fileName: string, file: Blob | ArrayBuffer): Promise<void> {
    await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).attachmentFiles.add(fileName, file)
  }

  async deleteAttachment(listTitle: string, itemId: number, fileName: string): Promise<void> {
    await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).attachmentFiles.getByName(fileName).delete()
  }

  // --------------------------------------------------------------------------
  // 4. FILES & FOLDERS
  // --------------------------------------------------------------------------

  async uploadFile(
      serverRelativeUrl: string,
      fileName: string,
      file: Blob | ArrayBuffer
  ): Promise<string> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const r = await this.sp.web
        .getFolderByServerRelativePath(fullUrl)
        .files.addUsingPath(fileName, file, { Overwrite: true })

    // FIX: Cast to 'any' to avoid TS2339 when TS infers return type as IFileInfo
    // Runtime returns { data: ..., file: ... }
    return (r as any).data.ServerRelativeUrl
  }

  async downloadFile(serverRelativeUrl: string): Promise<Blob> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    return await this.sp.web
        .getFileByServerRelativePath(fullUrl)
        .getBlob()
  }

  async updateFileMetadata(
      serverRelativeUrl: string,
      payload: Record<string, any>
  ): Promise<void> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const item = await this.sp.web
        .getFileByServerRelativePath(fullUrl)
        .getItem()
    await item.update(payload)
  }

  async deleteFile(serverRelativeUrl: string): Promise<void> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    await this.sp.web.getFileByServerRelativePath(fullUrl).recycle()
  }

  async createFolder(serverRelativeUrl: string): Promise<void> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    await this.sp.web.folders.addUsingPath(fullUrl)
  }

  // --------------------------------------------------------------------------
  // 5. WEBS & LISTS (STRUCTURE)
  // --------------------------------------------------------------------------

  async getWebInfo(): Promise<WebInfo> {
    const w = await this.sp.web()
    return {
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description
    }
  }

  async getSubwebs(): Promise<WebInfo[]> {
    const webs = await this.sp.web.webs()
    return webs.map(w => ({
      Id: w.Id,
      Title: w.Title,
      Url: w.Url,
      Description: w.Description
    }))
  }

  async getLists(): Promise<ListInfo[]> {
    const lists = await this.sp.web.lists()
    return lists.map(l => ({
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl
    }))
  }

  async getList(listTitle: string): Promise<ListInfo> {
    const l = await this.sp.web.lists.getByTitle(listTitle)()
    return {
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl
    }
  }

  async createList(title: string, description?: string, template = 100): Promise<ListInfo> {
    const r = await this.sp.web.lists.add(title, description, template)
    // @ts-ignore - PnPjs add() returns { data: ..., list: ... }
    const l = r.data
    return {
      Id: l.Id,
      Title: l.Title,
      Description: l.Description,
      ItemCount: l.ItemCount,
      Hidden: l.Hidden,
      ImageUrl: l.ImageUrl
    }
  }

  async deleteList(title: string): Promise<void> {
    await this.sp.web.lists.getByTitle(title).delete()
  }

  // --------------------------------------------------------------------------
  // 6. USERS & GROUPS
  // --------------------------------------------------------------------------

  async getCurrentUser(): Promise<UserInfo> {
    return await this.sp.web.currentUser()
  }

  async getSiteUsers(): Promise<UserInfo[]> {
    return await this.sp.web.siteUsers()
  }

  async ensureUser(loginName: string): Promise<UserInfo> {
    const result = await this.sp.web.ensureUser(loginName)
    // @ts-ignore
    return result.data
  }

  async getUserGroups(email?: string): Promise<SiteGroup[]> {
    if (email) {
      return await this.sp.web.siteUsers.getByEmail(email).groups()
    }
    return await this.sp.web.currentUser.groups()
  }

  async addUserToGroup(groupName: string, loginName: string): Promise<void> {
    const user = await this.sp.web.ensureUser(loginName)
    // @ts-ignore
    await this.sp.web.siteGroups.getByName(groupName).users.add(user.data.LoginName)
  }

  async removeUserFromGroup(groupName: string, loginName: string): Promise<void> {
    // We need User ID to remove from group in some versions, but LoginName usually works if using proper method.
    // PnPjs 'removeByLoginName' exists on users collection
    await this.sp.web.siteGroups.getByName(groupName).users.removeByLoginName(loginName)
  }

  async createGroup(groupName: string, description?: string): Promise<SiteGroup> {
    const r = await this.sp.web.siteGroups.add({ Title: groupName, Description: description })
    // @ts-ignore
    return r.data
  }

  async getUserEffectivePermissions(
      email?: string
  ): Promise<SPBasePermissions> {
    if (email) {
      // Resolve user first to get LoginName
      const user = await this.sp.web.siteUsers.getByEmail(email)()
      return await this.sp.web.getUserEffectivePermissions(user.LoginName)
    }
    return await this.sp.web.getCurrentUserEffectivePermissions()
  }

  // --------------------------------------------------------------------------
  // 7. FIELDS & METADATA
  // --------------------------------------------------------------------------

  async getListFields(listTitle: string): Promise<FieldDefinition[]> {
    const r = await this.sp.web.lists
        .getByTitle(listTitle)
        .fields.filter('Hidden eq false')()

    return r.map((f: any) => ({
      InternalName: f.InternalName,
      Title: f.Title,
      TypeAsString: f.TypeAsString,
      Hidden: f.Hidden,
      Choices: f.Choices || [],
    }))
  }

  async getFieldChoices(
      listTitle: string,
      fieldInternalName: string
  ): Promise<string[]> {
    const f: any = await this.sp.web.lists
        .getByTitle(listTitle)
        .fields.getByInternalNameOrTitle(fieldInternalName)()
    return f.Choices || []
  }

  // --------------------------------------------------------------------------
  // 8. VERSION HISTORY
  // --------------------------------------------------------------------------

  async getFileVersions(serverRelativeUrl: string): Promise<FileVersion[]> {
    const fullUrl = getServerRelativePath(serverRelativeUrl, this.baseUrl)
    const versions = await this.sp.web
        .getFileByServerRelativePath(fullUrl)
        .versions.expand('CreatedBy')()

    return versions.map((v: any) => ({
      VersionLabel: v.VersionLabel,
      Created: v.Created,
      CheckInComment: v.CheckInComment,
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

  getVersionHistoryLink(serverRelativeUrl: string): string {
    return `${
        this.baseUrl || ''
    }/_layouts/15/Versions.aspx?FileName=${encodeURIComponent(
        serverRelativeUrl
    )}`
  }

  // --------------------------------------------------------------------------
  // PRIVATE HELPERS
  // --------------------------------------------------------------------------

  private buildKql(opts: SearchRequestOptions): string {
    const parts: string[] = []

    // 1. Query Text
    const txt = opts.query?.trim() || '*'
    if (txt !== '*') {
      const escapedTxt = txt.replace(/"/g, '""')
      if (opts.searchTitleOnly) {
        // Logic for "Name Only" search
        parts.push(`(Title:"${escapedTxt}*" OR Filename:"${escapedTxt}*")`)
      } else {
        // Logic for "Everything" search (Title, Name, and Content)
        parts.push(`(Title:"${escapedTxt}*" OR Filename:"${escapedTxt}*" OR ${escapedTxt}*)`)
      }
    } else {
      parts.push('*')
    }

    // 1.5 File Types (Include / Exclude)
    if (opts.fileTypes?.include?.length) {
      // Inclusion
      parts.push(
          `(${opts.fileTypes.include.map((e) => `FileExtension:${e}`).join(' OR ')})`
      )
    }
    if (opts.fileTypes?.exclude?.length) {
      // Exclusion
      parts.push(
          `(${opts.fileTypes.exclude.map((e) => `NOT FileExtension:${e}`).join(' AND ')})`
      )
    }

    // 2. Scope (Path)
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

    // 3. Result Type
    if (opts.resultType === 'items') {
      parts.push(`(NOT ContentTypeId:0x0120*)`)
    } else if (opts.resultType === 'folders') {
      parts.push(`ContentTypeId:0x0120*`)
    }

    // 4. Filters
    if (opts.filters) {
      for (const [key, value] of Object.entries(opts.filters)) {
        if (!value || (Array.isArray(value) && value.length === 0)) continue
        let mp = key
        // Resolve mapped property if alias is used
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

  // --- Manual Search Cache Helpers ---
  private generateSearchKey(opts: SearchRequestOptions): string {
    return `SP_SEARCH_${this.hashCode(JSON.stringify(opts))}`
  }

  private readFromCache<T>(key: string): T | null {
    try {
      const item = sessionStorage.getItem(key)
      if (!item) return null
      const parsed = JSON.parse(item)
      if (Date.now() > parsed.expiry) {
        sessionStorage.removeItem(key)
        return null
      }
      return parsed.data as T
    } catch {
      return null
    }
  }

  private writeToCache(key: string, data: any) {
    try {
      const payload = {
        data,
        expiry: Date.now() + 15 * 60 * 1000, // 15 Mins
      }
      sessionStorage.setItem(key, JSON.stringify(payload))
    } catch (e) {
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
}
