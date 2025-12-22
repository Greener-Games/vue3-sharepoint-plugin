import {
  EmailProperties,
  FieldDefinition,
  FileVersion,
  IBatch,
  ISharePointClient,
  SearchRequestOptions,
  SearchResult,
  SiteGroup,
  SPBasePermissions,
  UserInfo,
  WebInfo,
  ListInfo,
  AttachmentInfo,
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
  lists?: Record<string, any[]>

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
        Description: 'This is a mock web'
      },
      ...data,
    }
    this.logger = new Logger(data.debug)
  }

  private async wait() {
    return new Promise((resolve) => setTimeout(resolve, this.data.delay))
  }

  // --- 1. SEARCH ---
  async search<T = any>(opts: SearchRequestOptions): Promise<SearchResult<T>> {
    await this.wait()
    this.logger.log('Search Options:', opts)

    let allItems: any[] = []

    // --- STEP 1: Smart Aggregation based on Scope ---
    // If scope is provided, try to find which Mock List it refers to.
    const definedListNames = Object.keys(this.data.lists || {})
    let targetListKeys: string[] = []

    if (opts.scope) {
      const scopes = Array.isArray(opts.scope) ? opts.scope : [opts.scope]

      // Check if any defined List Name appears inside the requested Scope URL
      // e.g. Scope: "/sites/V2/Shared Documents", List Key: "Shared Documents" -> MATCH
      targetListKeys = definedListNames.filter(listName => {
        // Normalize to lowercase for comparison
        return scopes.some(s => s.toLowerCase().includes(listName.toLowerCase()))
      })
    }

    // If we found specific lists matching the scope, only load those.
    // Otherwise (or if scope is empty), load ALL lists.
    const listsToSearch = targetListKeys.length > 0 ? targetListKeys : definedListNames

    listsToSearch.forEach(key => {
      const listItems = this.data.lists![key] || []
      allItems = [...allItems, ...listItems]
    })
    // ------------------------------------------------


    // --- STEP 2: Path Filtering (Sub-folders) ---
    // Even if we selected the right list, the scope might be a sub-folder.
    if (opts.scope) {
      const scopes = Array.isArray(opts.scope) ? opts.scope : [opts.scope]

      allItems = allItems.filter(item => {
        // If item has no path, we can't verify sub-folder, so exclude safely
        if (!item.Path && !item.url) return false

        const itemPath = (item.Path || item.url).toLowerCase()
        // Does the item path start with the scope?
        return scopes.some(s => itemPath.includes(s.toLowerCase()))
      })
    }

    // --- STEP 3: Text Search ---
    const qRaw = opts.query || ''
    if (qRaw && qRaw !== '*') {
      const q = qRaw.toLowerCase()
      allItems = allItems.filter(item => {
        const titleMatch = item.Title?.toLowerCase().includes(q)
        let summaryMatch = false
        if (item.HitHighlightedSummary?.toLowerCase().includes(q)) {
          summaryMatch = true
          const regex = new RegExp(`(${q})`, 'gi')
          item.HitHighlightedSummary = item.HitHighlightedSummary.replace(regex, '<c0>$1</c0>')
        }
        return titleMatch || summaryMatch
      })
    }

    // --- STEP 3.4: File Types (Inclusion / Exclusion) (Mock) ---
    if (opts.fileTypes?.include?.length) {
      const includes = opts.fileTypes.include.map(e => e.toLowerCase())
      allItems = allItems.filter(item => {
        const path = (item.Path || item.url || '').toLowerCase()
        return includes.some(ext => path.endsWith(`.${ext}`))
      })
    }

    if (opts.fileTypes?.exclude?.length) {
      const excludes = opts.fileTypes.exclude.map(e => e.toLowerCase())
      allItems = allItems.filter(item => {
        const path = (item.Path || item.url || '').toLowerCase()
        return !excludes.some(ext => path.endsWith(`.${ext}`))
      })
    }

    // --- STEP 3.5: Result Type Filtering (Mock) ---
    if (opts.resultType === 'folders') {
      // In Mock, we assume items in `data.lists` are File/Items unless they have FSObjType=1
      // If no FSObjType property exists, we assume it's NOT a folder.
      allItems = allItems.filter(item => item.FSObjType === 1 || item.FileSystemObjectType === 1)
    } else if (opts.resultType === 'items') {
      // Exclude folders
      allItems = allItems.filter(item => item.FSObjType !== 1 && item.FileSystemObjectType !== 1)
    }

    // --- STEP 4: Metadata Filters ---
    if (opts.filters) {
      Object.entries(opts.filters).forEach(([key, value]) => {
        if (!value || (Array.isArray(value) && value.length === 0)) return

        let dataKey = key
        if (opts.mapping) {
          const found = Object.keys(opts.mapping).find(k => opts.mapping![k] === key)
          if(found) dataKey = found
        }

        allItems = allItems.filter(item => {
          const itemVal = item[dataKey]
          if (!itemVal) return false
          if (Array.isArray(value)) return value.some(v => String(v) === String(itemVal))
          return String(itemVal) === String(value)
        })
      })
    }

    // --- STEP 5: Pagination & Mapping ---
    const totalHits = allItems.length
    const start = opts.startRow || 0
    const limit = opts.rowLimit || 10

    const mappedItems = allItems.slice(start, start + limit).map(item => {
      const map: any = { ...item }

      if (opts.includeRelativePath && map.Path) {
        try {
          map.relativePath = decodeURIComponent(new URL(map.Path).pathname)
        } catch {
          map.relativePath = map.Path
        }
      }

      if (!opts.mapping) return map

      const newItem: any = { ...item }
      Object.entries(opts.mapping).forEach(([spKey, friendly]) => {
        newItem[friendly] = item[spKey]
      })
      if (!newItem.url) newItem.url = item.Path
      if (opts.includeRelativePath) newItem.relativePath = map.relativePath
      return newItem
    })

    return { items: mappedItems as T[], totalHits, startRow: start }
  }

  // --- 2. UTILITIES & PERMISSIONS ---

  async getCurrentUser(): Promise<UserInfo> {
    await this.wait()
    return this.data.currentUser!
  }

  async getSiteUsers(): Promise<UserInfo[]> {
    await this.wait()
    return this.data.siteUsers || [this.data.currentUser!]
  }

  async ensureUser(loginName: string): Promise<UserInfo> {
    await this.wait()
    const user = this.data.siteUsers?.find(
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
    this.data.siteUsers?.push(newUser)
    return newUser
  }

  async getUserGroups(email?: string): Promise<SiteGroup[]> {
    await this.wait()
    const targetEmail = email || this.data.currentUser!.Email
    return this.data.userGroups?.[targetEmail] || []
  }

  async addUserToGroup(groupName: string, loginName: string): Promise<void> {
    await this.wait()
    // Find or create user
    const user = await this.ensureUser(loginName)

    // Find Group
    let group = this.data.siteGroups?.find(g => g.Title === groupName)
    if (!group) {
      // Mock creating group if not exists in site groups for simplicity
      group = { Id: Math.floor(Math.random() * 1000), Title: groupName }
      this.data.siteGroups?.push(group)
    }

    // Add to userGroups map
    if (!this.data.userGroups![user.Email]) {
      this.data.userGroups![user.Email] = []
    }
    if (!this.data.userGroups![user.Email].find(g => g.Title === groupName)) {
      this.data.userGroups![user.Email].push(group)
    }
    this.logger.log(`Added ${loginName} to group ${groupName}`)
  }

  async removeUserFromGroup(groupName: string, loginName: string): Promise<void> {
    await this.wait()
    // Find user (by email approx)
    const user = this.data.siteUsers?.find(u => u.LoginName === loginName || u.Email === loginName)
    if (!user) return

    const groups = this.data.userGroups![user.Email]
    if (groups) {
      this.data.userGroups![user.Email] = groups.filter(g => g.Title !== groupName)
    }
    this.logger.log(`Removed ${loginName} from group ${groupName}`)
  }

  async createGroup(groupName: string, description?: string): Promise<SiteGroup> {
    await this.wait()
    const newGroup = { Id: Math.floor(Math.random() * 1000), Title: groupName, Description: description }
    this.data.siteGroups?.push(newGroup)
    this.logger.log(`Created group ${groupName}`)
    return newGroup
  }

  async getUserEffectivePermissions(
    email?: string
  ): Promise<SPBasePermissions> {
    await this.wait()
    const targetEmail = email || this.data.currentUser!.Email
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

  async getFileVersions(url: string): Promise<FileVersion[]> {
    await this.wait()
    return this.data.fileVersions?.[url] || []
  }

  getVersionHistoryLink(url: string): string {
    return `#mock-history/${url}`
  }

  async uploadFile(
    url: string,
    name: string,
    file: Blob | ArrayBuffer
  ): Promise<string> {
    await this.wait()
    // Normalize url if it's site relative
    let targetUrl = url
    if (!targetUrl.startsWith('/') && !targetUrl.startsWith('http')) {
        // Mock doesn't strictly have a "base URL" context in the same way, but usually keys are full server relative.
        // However, we passed 'webInfo.Url' in constructor default.
        // Let's assume we prepend the webUrl if it's missing slash.
        targetUrl = `${this.data.webInfo!.Url}/${url}`
    }

    const fullPath = `${targetUrl}/${name}`
    this.data.files![fullPath] =
      file instanceof ArrayBuffer ? new Blob([file]) : file
    this.logger.log(`Uploaded ${fullPath}`)
    return fullPath
  }

  async downloadFile(url: string): Promise<Blob> {
    await this.wait()
    let targetUrl = url
    if (!targetUrl.startsWith('/') && !targetUrl.startsWith('http')) {
        targetUrl = `${this.data.webInfo!.Url}/${url}`
    }

    const f = this.data.files![targetUrl]
    if (!f) throw new Error('File not found')
    return typeof f === 'string' ? new Blob([f]) : f
  }

  // --- 4. CRUD & BATCH (Stubs) ---
  async executeBatch(builder: (batch: IBatch) => void) {
    builder({} as any)
  }
  async createListItem(list: string, payload: any) {
    if (!this.data.lists![list]) this.data.lists![list] = []
    const newItem = { Id: Math.floor(Math.random() * 1000), ...payload }
    this.data.lists![list].push(newItem)
    return newItem
  }
  async updateListItem(list: string, id: number, payload: any) {
    this.logger.log(`Update ${list} #${id}`, payload)
    const item = this.data.lists![list]?.find(i => i.Id === id)
    if (item) Object.assign(item, payload)
  }
  async deleteListItem(list: string, id: number) {
    this.logger.log(`Delete ${list} #${id}`)
    if (this.data.lists![list]) {
      this.data.lists![list] = this.data.lists![list].filter(i => i.Id !== id)
    }
  }
  async getListItemById(list: string, id: number) {
    return this.data.lists![list]?.find((i) => i.Id === id)
  }
  async getListItems<T = any>(
    list: string,
    query?: string,
    select?: string[],
    expand?: string[]
  ): Promise<T[]> {
    // Basic Mock Filtering (only supports exact match key=value or contains for now if we want to get fancy,
    // but for now return all)
    // TODO: Implement basic OData filter parsing for mock if needed.
    return (this.data.lists![list] || []) as T[]
  }
  async updateFileMetadata(url: string, payload: any) {
    this.logger.log(`Update Meta ${url}`, payload)
  }
  async deleteFile(url: string) {
    delete this.data.files![url]
  }
  async createFolder(url: string) {
    this.logger.log(`Create Folder ${url}`)
  }
  async restoreItem(id: string | number) {
    this.logger.log(`Restore ${id}`)
  }
  async getListFields() {
    return []
  }
  async getFieldChoices() {
    return []
  }

  // --- 5. WEBS & LISTS ---
  async getWebInfo(): Promise<WebInfo> {
    await this.wait()
    return this.data.webInfo!
  }

  async getSubwebs(): Promise<WebInfo[]> {
    await this.wait()
    return this.data.subwebs || []
  }

  async getLists(): Promise<ListInfo[]> {
    await this.wait()

    // 1. Get explicitly defined listsInfo
    const explicitInfos = this.data.listsInfo || []

    // 2. Get lists that exist in data.lists but NOT in explicit listsInfo
    const knownTitles = new Set(explicitInfos.map(l => l.Title))
    const derivedInfos = Object.keys(this.data.lists || {})
      .filter(k => !knownTitles.has(k))
      .map((k, i) => ({
        Id: `mock-list-${i}`,
        Title: k,
        Description: 'Mock List',
        ItemCount: this.data.lists![k].length,
        Hidden: false,
        ImageUrl: ''
      }))

    return [...explicitInfos, ...derivedInfos]
  }

  async getList(listTitle: string): Promise<ListInfo> {
    await this.wait()
    const lists = await this.getLists()
    const found = lists.find(l => l.Title === listTitle)
    if (!found) throw new Error(`List ${listTitle} not found`)
    return found
  }

  async createList(title: string, description?: string, template?: number): Promise<ListInfo> {
    await this.wait()
    this.data.lists![title] = []
    const info = {
      Id: `mock-list-${Date.now()}`,
      Title: title,
      Description: description || '',
      ItemCount: 0,
      Hidden: false,
      ImageUrl: ''
    }
    if (!this.data.listsInfo) this.data.listsInfo = []
    this.data.listsInfo.push(info)
    return info
  }

  async deleteList(title: string): Promise<void> {
    await this.wait()
    delete this.data.lists![title]
    if (this.data.listsInfo) {
      this.data.listsInfo = this.data.listsInfo.filter(l => l.Title !== title)
    }
  }

  // --- 6. ATTACHMENTS ---
  async getItemAttachments(listTitle: string, itemId: number): Promise<AttachmentInfo[]> {
    // Mock implementation: return empty or fake attachments
    return []
  }

  async addAttachment(listTitle: string, itemId: number, fileName: string, file: Blob | ArrayBuffer): Promise<void> {
    this.logger.log(`Added attachment ${fileName} to ${listTitle} item ${itemId}`)
  }

  async deleteAttachment(listTitle: string, itemId: number, fileName: string): Promise<void> {
     this.logger.log(`Deleted attachment ${fileName} from ${listTitle} item ${itemId}`)
  }
}
