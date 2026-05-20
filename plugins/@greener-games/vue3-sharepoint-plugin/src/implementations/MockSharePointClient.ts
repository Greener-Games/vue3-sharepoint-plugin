import type {
  EmailProperties,
  FileVersion,
  IBatch,
  ISharePointClient,
  ListItemQueryOptions,
  SearchRequestOptions,
  SearchResult,
  SiteGroup,
  SPBasePermissions,
  UserInfo,
  WebInfo,
  ListInfo,
  AttachmentInfo,
  FieldDefinition,
} from '../types'
import { Logger } from '../utils/debug'

export interface MockData {
  /** Current Logged in User */
  currentUser?: UserInfo

  /** Simulated Network Latency (ms) */
  delay?: number

  /**
   * Searchable Lists.
   * Key: List Title, Value: Array of Items
   */
  lists?: Record<string, Record<string, unknown>[]>

  /**
   * Files (for download/upload simulation).
   * Key: ServerRelativeUrl, Value: Blob
   */
  files?: Record<string, Blob | string>

  /**
   * All Users in the "Site" (for People Pickers).
   */
  siteUsers?: UserInfo[]

  /**
   * Groups a user belongs to.
   * Key: User Email, Value: Array of Groups
   */
  userGroups?: Record<string, SiteGroup[]>

  /**
   * All Groups in the "Site"
   */
  siteGroups?: SiteGroup[]

  /**
   * Version History for files.
   * Key: File ServerRelativeUrl, Value: Array of Versions
   */
  fileVersions?: Record<string, FileVersion[]>

  /**
   * Web Info
   */
  webInfo?: WebInfo

  /**
   * Subwebs
   */
  subwebs?: WebInfo[]

  /**
   * List Info
   */
  listsInfo?: ListInfo[]
  /**
   * Debug flag
   */
  debug?: boolean
}

export class MockSharePointClient implements ISharePointClient {
  private data: MockData
  private logger: Logger

  constructor(data: MockData = {}) {
    // Set Defaults if not provided
    this.data = {
      lists: {},
      files: {},
      userGroups: {},
      fileVersions: {},
      siteUsers: [],
      siteGroups: [],
      subwebs: [],
      listsInfo: [],
      delay: 500,
      currentUser: {
        Id: 1,
        Title: 'Mock User',
        Email: 'mock@local',
        LoginName: 'i:0#.f|mock',
      },
      webInfo: {
        Id: 'mock-web-id',
        Title: 'Mock Web',
        Url: '/sites/mock',
        Description: 'This is a mock web',
      },
      ...data,
    }
    this.logger = new Logger(data.debug)
  }

  /** Gets the configured base URL */
  public getBaseUrl(): string {
    return this.data.webInfo ? this.data.webInfo.Url : '/sites/mock'
  }

  /** Performs a raw request using native fetch */
  public async request<T>(_endpoint: string, _options: RequestInit = {}): Promise<T> {
    await this.wait()
    return null as unknown as T
  }

  private async wait() {
    return new Promise((resolve) => setTimeout(resolve, this.data.delay))
  }

  // --- 1. SEARCH ---
  /** Executes a search using SharePoint Search API */
  async search<T>(opts: SearchRequestOptions, _abortSignal?: AbortSignal): Promise<SearchResult<T>> {
    await this.wait()
    this.logger.log('Search Options:', opts)

    let allItems: Record<string, unknown>[] = []

    // --- STEP 1: Smart Aggregation based on Scope ---
    // If scope is provided, try to find which Mock List it refers to.
    const definedListNames = Object.keys(this.data.lists || {})
    let targetListKeys: string[] = []

    if (opts.scope) {
      const scopes = Array.isArray(opts.scope) ? opts.scope : [opts.scope]

      // Check if any defined List Name appears inside the requested Scope URL
      // e.g. Scope: "/sites/V2/Shared Documents", List Key: "Shared Documents" -> MATCH
      targetListKeys = definedListNames.filter((listName) => {
        // Normalize to lowercase for comparison
        return scopes.some((s) =>
          s.toLowerCase().includes(listName.toLowerCase())
        )
      })
    }

    // If we found specific lists matching the scope, only load those.
    // Otherwise (or if scope is empty), load ALL lists.
    const listsToSearch =
      targetListKeys.length > 0 ? targetListKeys : definedListNames

    listsToSearch.forEach((key) => {
      const listItems = this.data.lists ? this.data.lists[key] : []
      if (listItems) {
          allItems = [...allItems, ...listItems]
      }
    })
    // ------------------------------------------------

    // --- STEP 2: Path Filtering (Sub-folders) ---
    // Even if we selected the right list, the scope might be a sub-folder.
    if (opts.scope) {
      const scopes = Array.isArray(opts.scope) ? opts.scope : [opts.scope]

      allItems = allItems.filter((item: Record<string, unknown>) => {
        // If item has no path, we can't verify sub-folder, so exclude safely
        if (!item.Path && !item.url) return false

        const itemPath = String(item.Path as string || item.url as string || '').toLowerCase()
        // Does the item path start with the scope?
        return scopes.some((s) => itemPath.includes(s.toLowerCase()))
      })
    }

    // --- STEP 3: Text Search ---
    const qRaw = opts.query || ''
    if (qRaw && qRaw !== '*') {
      const q = qRaw.toLowerCase()
      allItems = allItems.filter((item: Record<string, unknown>) => {
        const titleMatch = (typeof item.Title === 'string' ? item.Title : '').toLowerCase().includes(q)
        // In SP, Filename is often FileLeafRef or derived from Path
        const filenameMatch = String(item.FileLeafRef as string || item.Name as string || '')
          .toLowerCase()
          .includes(q)
        const pathMatch = String(item.Path as string || item.url as string || '')
          .toLowerCase()
          .includes(q)

        const summaryMatch =
          String(item.HitHighlightedSummary as string || '').toLowerCase().includes(q)

        return titleMatch || filenameMatch || pathMatch || summaryMatch
      })
    }

    // --- STEP 3.4: File Types (Inclusion / Exclusion) (Mock) ---
    if (opts.fileTypes?.include?.length) {
      const includes = opts.fileTypes.include.map((e) => e.toLowerCase())
      allItems = allItems.filter((item: Record<string, unknown>) => {
        const path = String(item.Path as string || item.url as string || '').toLowerCase()
        return includes.some((ext) => path.endsWith(`.${ext}`))
      })
    }

    if (opts.fileTypes?.exclude?.length) {
      const excludes = opts.fileTypes.exclude.map((e) => e.toLowerCase())
      allItems = allItems.filter((item: Record<string, unknown>) => {
        const path = String(item.Path as string || item.url as string || '').toLowerCase()
        return !excludes.some((ext) => path.endsWith(`.${ext}`))
      })
    }

    // --- STEP 3.5: Result Type Filtering (Mock) ---
    if (opts.resultType === 'folders') {
      // In Mock, we assume items in `data.lists` are File/Items unless they have FSObjType=1
      // If no FSObjType property exists, we assume it's NOT a folder.
      allItems = allItems.filter(
        (item) => item.FSObjType === 1 || item.FileSystemObjectType === 1
      )
    } else if (opts.resultType === 'items') {
      // Exclude folders
      allItems = allItems.filter(
        (item) => item.FSObjType !== 1 && item.FileSystemObjectType !== 1
      )
    }

    // --- STEP 4: Metadata Filters ---
    if (opts.filters) {
      Object.entries(opts.filters).forEach(([key, value]) => {
        if (!value || (Array.isArray(value) && value.length === 0)) return

        let dataKey = key
        if (opts.mapping) {
          const mapping = opts.mapping
          const found = Object.keys(mapping).find(
            (k) => mapping[k] === key
          )
          if (found) dataKey = found
        }

        allItems = allItems.filter((item) => {
          const itemVal = item[dataKey]
          if (itemVal === undefined || itemVal === null) return false

          // eslint-disable-next-line @typescript-eslint/no-base-to-string
          const normalizedItemVal = (typeof itemVal === 'object' && itemVal !== null) ? JSON.stringify(itemVal).toLowerCase() : String(itemVal).toLowerCase()

          if (Array.isArray(value)) {
            return value.some((v) =>
              normalizedItemVal.includes(String(v).toLowerCase())
            )
          }

          return normalizedItemVal.includes(String(value).toLowerCase())
        })
      })
    }

    // --- STEP 4.5: Sorting (Mock) ---
    const sortList = opts.sortList
    if (sortList && sortList.length > 0) {
      allItems.sort((a, b) => {
        for (const sort of sortList) {
          const valA = a[sort.property]
          const valB = b[sort.property]

          if (valA === valB) continue
          if ((typeof valA !== 'string' && typeof valA !== 'number') || (typeof valB !== 'string' && typeof valB !== 'number')) continue

          const comparison = (valA) > (valB) ? 1 : -1
          return sort.direction === 'ascending' ? comparison : -comparison
        }
        return 0
      })
    }

    // --- STEP 5: Pagination & Hydration & Mapping ---
    const totalHits = allItems.length
    const start = opts.startRow || 0
    const limit = opts.rowLimit || 10

    let itemsSlice = allItems.slice(start, start + limit).map((item) => {
      const map = { ...item }

      if (opts.includeRelativePath && typeof map.Path === 'string') {
        try {
          map.relativePath = decodeURIComponent(new URL(map.Path).pathname)
        } catch {
          map.relativePath = map.Path
        }
      }
      return map
    })

    // Hydration (Mock - just merging, as we already have full objects in Mock lists usually)
    if (opts.expandFields) {
      // Pass (No-op for Mock unless we want to simulate partial objects first)
    }

    // Mapping
    const mapping = opts.mapping
    if (mapping) {
      itemsSlice = itemsSlice.map((map) => {
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

    return { items: itemsSlice as unknown as T[], totalHits, startRow: start }
  }

  // --- 2. UTILITIES & PERMISSIONS ---

  async getCurrentUser(_abortSignal?: AbortSignal): Promise<UserInfo> {
    await this.wait()
    return this.data.currentUser!
  }

  async getSiteUsers(_abortSignal?: AbortSignal): Promise<UserInfo[]> {
    await this.wait()
    return this.data.siteUsers || [this.data.currentUser!]
  }

  async searchUsers(query: string, _abortSignal?: AbortSignal): Promise<UserInfo[]> {
    await this.wait()
    if (!query) return []

    const q = query.toLowerCase()
    const allUsers = this.data.siteUsers || [this.data.currentUser!]

    return allUsers.filter(
      (u) =>
        u.Title.toLowerCase().includes(q) || u.Email.toLowerCase().includes(q)
    )
  }

  async ensureUser(loginName: string, _abortSignal?: AbortSignal): Promise<UserInfo> {
    await this.wait()
    const siteUsers = this.data.siteUsers || []
    const user = siteUsers.find(
      (u) => u.LoginName === loginName || u.Email === loginName
    )
    if (user) return user

    // Create new mock user
    const newUser = {
      Id: Math.floor(Math.random() * 10000),
      Title: 'Ensured User',
      Email: loginName,
      LoginName: loginName,
    }
    if (this.data.siteUsers) {
        this.data.siteUsers.push(newUser)
    } else {
        this.data.siteUsers = [newUser]
    }
    return newUser
  }

  async getUserGroups(email?: string, _abortSignal?: AbortSignal): Promise<SiteGroup[]> {
    await this.wait()
    const targetEmail = email || (this.data.currentUser!).Email
    return this.data.userGroups?.[targetEmail] || []
  }

  async addUserToGroup(groupName: string, loginName: string, _abortSignal?: AbortSignal): Promise<void> {
    await this.wait()
    // Find or create user
    const user = await this.ensureUser(loginName)

    // Find Group
    let group = this.data.siteGroups?.find((g) => g.Title === groupName)
    if (!group) {
      // Mock creating group if not exists in site groups for simplicity
      group = { Id: Math.floor(Math.random() * 1000), Title: groupName }
      if (!this.data.siteGroups) this.data.siteGroups = []
      this.data.siteGroups.push(group)
    }

    // Add to userGroups map
    const groupsMap = this.data.userGroups || (this.data.userGroups = {})
    if (!groupsMap[user.Email]) {
      groupsMap[user.Email] = []
    }
    if (!groupsMap[user.Email].find((g) => g.Title === groupName)) {
      groupsMap[user.Email].push(group)
    }
    this.logger.log(`Added ${loginName} to group ${groupName}`)
  }

  async removeUserFromGroup(
    groupName: string,
    loginName: string,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    await this.wait()
    // Find user (by email approx)
    const user = this.data.siteUsers?.find(
      (u) => u.LoginName === loginName || u.Email === loginName
    )
    if (!user) return

    const userGroups = this.data.userGroups
    if (userGroups) {
      const groups = userGroups[user.Email]
      if (groups) {
        userGroups[user.Email] = groups.filter((g) => g.Title !== groupName)
      }
    }
    this.logger.log(`Removed ${loginName} from group ${groupName}`)
  }

  async createGroup(
    groupName: string,
    description?: string,
    _abortSignal?: AbortSignal
  ): Promise<SiteGroup> {
    await this.wait()
    const newGroup = {
      Id: Math.floor(Math.random() * 1000),
      Title: groupName,
      Description: description,
    }
    if (!this.data.siteGroups) this.data.siteGroups = []
    this.data.siteGroups.push(newGroup)
    this.logger.log(`Created group ${groupName}`)
    return newGroup
  }

  async getUserEffectivePermissions(
    email?: string,
    _abortSignal?: AbortSignal
  ): Promise<SPBasePermissions> {
    await this.wait()
    const targetEmail = email || (this.data.currentUser!).Email
    const groups = this.data.userGroups?.[targetEmail] || []

    // Logic: If in 'Owners', return Full Control. Else Read.
    if (groups.some((g) => g.Title.includes('Owner'))) {
      return { High: 2147483647, Low: 4294967295 }
    }
    return { High: 0, Low: 1 } // View Only
  }

  async sendEmail(props: EmailProperties): Promise<void> {
    await this.wait()
    this.logger.log(
      `📧 Sent Email to [${props.To.join(', ')}]: ${props.Subject}`
    )
  }

  // --- 3. FILES & VERSIONS ---

  async getFileVersions(url: string, _abortSignal?: AbortSignal): Promise<FileVersion[]> {
    await this.wait()
    return this.data.fileVersions?.[url] || []
  }

  async getVersionHistoryLink(url: string, _abortSignal?: AbortSignal): Promise<string> {
    await this.wait()
    return `#mock-history/${url}`
  }

  getVersionHistoryLinkByItem(
    listId: string,
    itemId: number,
    _webUrl?: string
  ): string {
    return `#mock-history/${listId}/${itemId}`
  }

  async uploadFile(
    url: string,
    name: string,
    file: Blob | ArrayBuffer,
    _abortSignal?: AbortSignal
  ): Promise<string> {
    await this.wait()
    // Normalize url if it's site relative
    let targetUrl = url
    if (!targetUrl.startsWith('/') && !targetUrl.startsWith('http')) {
      const webUrl = this.data.webInfo ? this.data.webInfo.Url : '/sites/mock';
      targetUrl = `${webUrl}/${url}`
    }

    const fullPath = `${targetUrl}/${name}`
    if (!this.data.files) this.data.files = {}
    this.data.files[fullPath] =
      file instanceof ArrayBuffer ? new Blob([file]) : file
    this.logger.log(`Uploaded ${fullPath}`)
    return fullPath
  }

  async downloadFile(url: string, _abortSignal?: AbortSignal): Promise<Blob> {
    await this.wait()
    let targetUrl = url
    if (!targetUrl.startsWith('/') && !targetUrl.startsWith('http')) {
        const webUrl = this.data.webInfo ? this.data.webInfo.Url : '/sites/mock';
        targetUrl = `${webUrl}/${url}`
    }

    const f = this.data.files?.[targetUrl]
    if (!f) throw new Error('File not found')
    return typeof f === 'string' ? new Blob([f]) : f
  }

  // --- 4. CRUD & BATCH (Stubs) ---
  async executeBatch(builder: (batch: IBatch) => void, _abortSignal?: AbortSignal): Promise<void> {
    builder({
        createListItem: (list, payload) => { void this.createListItem(list, payload) },
        updateListItem: (list, id, payload) => { void this.updateListItem(list, id, payload) },
        deleteListItem: (list, id) => { void this.deleteListItem(list, id) },
        deleteFile: (url) => { void this.deleteFile(url) }
    })
    await Promise.resolve()
  }
  async createListItem<T>(list: string, payload: Record<string, unknown>, _abortSignal?: AbortSignal): Promise<T> {
    await this.wait()
    const allLists = this.data.lists || (this.data.lists = {})
    if (!allLists[list]) allLists[list] = []
    const newItem = { Id: Math.floor(Math.random() * 1000), ...payload }
    allLists[list].push(newItem)
    return newItem as unknown as T
  }
  async updateListItem(list: string, id: number, payload: Record<string, unknown>, _abortSignal?: AbortSignal): Promise<void> {
    await this.wait()
    this.logger.log(`Update ${list} #${id}`, payload)
    const item = this.data.lists?.[list]?.find((i) => i.Id === id)
    if (item) Object.assign(item, payload)
  }
  async deleteListItem(list: string, id: number, _abortSignal?: AbortSignal): Promise<void> {
    await this.wait()
    this.logger.log(`Delete ${list} #${id}`)
    const lists = this.data.lists
    if (lists && lists[list]) {
      lists[list] = lists[list].filter((i) => i.Id !== id)
    }
  }
  async getListItemById<T>(
    list: string,
    id: number,
    _select?: string[],
    _expand?: string[],
    _abortSignal?: AbortSignal
  ): Promise<T> {
    await this.wait()
    const item = this.data.lists?.[list]?.find((i) => i.Id === id)
    return item as unknown as T
  }

  async getListItems<T>(
    listTitle: string,
    options?: ListItemQueryOptions,
    _abortSignal?: AbortSignal
  ): Promise<T[]> {
    await this.wait()
    const allItems = this.data.lists?.[listTitle] || []
    const result = [...allItems]

    // Order By
    const orderByField = options?.orderBy
    if (orderByField) {
      result.sort((a, b) => {
        const valA = a[orderByField]
        const valB = b[orderByField]
        if (typeof valA !== 'string' && typeof valA !== 'number') return 0
        if (typeof valB !== 'string' && typeof valB !== 'number') return 0
        if (valA < valB) return options?.ascending === false ? 1 : -1
        if (valA > valB) return options?.ascending === false ? -1 : 1
        return 0
      })
    }

    let finalResult = result
    // Top
    if (options?.top && !options.getAll) {
      finalResult = result.slice(0, options.top)
    }

    if (options?.getAll && options?.onProgress) {
      options.onProgress(finalResult)
    }

    return finalResult as unknown as T[]
  }
  async updateFileMetadata(url: string, payload: Record<string, unknown>, _abortSignal?: AbortSignal): Promise<void> {
    await this.wait()
    this.logger.log(`Update Meta ${url}`, payload)
  }
  async deleteFile(url: string, _abortSignal?: AbortSignal): Promise<void> {
    await this.wait()
    if (this.data.files) {
      delete this.data.files[url]
    }
  }
  async createFolder(url: string, _abortSignal?: AbortSignal): Promise<void> {
    await this.wait()
    this.logger.log(`Create Folder ${url}`)
  }
  async restoreItem(id: string | number): Promise<void> {
    await this.wait()
    this.logger.log(`Restore ${id}`)
  }
  async getListFields(_listTitle: string, _webUrl?: string, _abortSignal?: AbortSignal): Promise<FieldDefinition[]> {
    await this.wait()
    return []
  }
  async getFieldChoices(): Promise<string[]> {
    await this.wait()
    return []
  }

  async searchTerm(
    _termSetId: string,
    label: string,
    _abortSignal?: AbortSignal
  ): Promise<{ Label: string; TermGuid: string } | null> {
    await this.wait()
    return {
      Label: label,
      TermGuid: '00000000-0000-0000-0000-000000000000'
    }
  }

  // --- 5. WEBS & LISTS ---
  async getWebInfo(_abortSignal?: AbortSignal): Promise<WebInfo> {
    await this.wait()
    return this.data.webInfo!
  }

  async getSubwebs(_abortSignal?: AbortSignal): Promise<WebInfo[]> {
    await this.wait()
    return this.data.subwebs || []
  }

  async getLists(_abortSignal?: AbortSignal): Promise<ListInfo[]> {
    await this.wait()

    // 1. Get explicitly defined listsInfo
    const explicitInfos = this.data.listsInfo || []

    // 2. Get lists that exist in data.lists but NOT in explicit listsInfo
    const knownTitles = new Set(explicitInfos.map((l) => l.Title))
    const derivedInfos = Object.keys(this.data.lists || {})
      .filter((k) => !knownTitles.has(k))
      .map((k, i) => ({
        Id: `mock-list-${i}`,
        Title: k,
        Description: 'Mock List',
        ItemCount: (this.data.lists?.[k] || []).length,
        Hidden: false,
        ImageUrl: '',
      }))

    return [...explicitInfos, ...derivedInfos]
  }

  async getList(listTitle: string, _abortSignal?: AbortSignal): Promise<ListInfo> {
    await this.wait()
    const lists = await this.getLists()
    const found = lists.find((l) => l.Title === listTitle)
    if (!found) throw new Error(`List ${listTitle} not found`)
    const items = this.data.lists?.[listTitle] || []
    found.ItemCount = items.length
    return found
  }

  async createList(
    title: string,
    description?: string,
    _template?: number,
    _abortSignal?: AbortSignal
  ): Promise<ListInfo> {
    await this.wait()
    if (!this.data.lists) {
      this.data.lists = {}
    }
    const lists = this.data.lists
    lists[title] = []

    const info: ListInfo = {
      Id: `mock-list-${Date.now()}`,
      Title: title,
      Description: description || '',
      ItemCount: 0,
      Hidden: false,
      ImageUrl: '',
    }
    if (!this.data.listsInfo) this.data.listsInfo = []
    this.data.listsInfo.push(info)
    return info
  }

  async deleteList(title: string, _abortSignal?: AbortSignal): Promise<void> {
    await this.wait()
    if (this.data.lists) {
      delete this.data.lists[title]
    }
    if (this.data.listsInfo) {
      this.data.listsInfo = this.data.listsInfo.filter((l) => l.Title !== title)
    }
  }

  // --- 6. ATTACHMENTS ---
  async getItemAttachments(
    _listTitle: string,
    _itemId: number,
    _abortSignal?: AbortSignal
  ): Promise<AttachmentInfo[]> {
    await this.wait()
    // Mock implementation: return empty or fake attachments
    return []
  }

  async addAttachment(
    listTitle: string,
    itemId: number,
    fileName: string,
    _file: Blob | ArrayBuffer,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    await this.wait()
    this.logger.log(
      `Added attachment ${fileName} to ${listTitle} item ${itemId}`
    )
  }

  async deleteAttachment(
    listTitle: string,
    itemId: number,
    fileName: string,
    _abortSignal?: AbortSignal
  ): Promise<void> {
    await this.wait()
    this.logger.log(
      `Deleted attachment ${fileName} from ${listTitle} item ${itemId}`
    )
  }
}