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
} from '../types'

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
   * Version History for files.
   * Key: File ServerRelativeUrl, Value: Array of Versions
   */
  fileVersions?: Record<string, FileVersion[]>
}

export class MockSharePointClient implements ISharePointClient {
  private data: MockData

  constructor(data: MockData = {}) {
    // Set Defaults if not provided
    this.data = {
      lists: {},
      files: {},
      userGroups: {},
      fileVersions: {},
      siteUsers: [],
      delay: 500,
      currentUser: {
        Id: 1,
        Title: 'Mock User',
        Email: 'mock@local',
        LoginName: 'i:0#.f|mock',
      },
      ...data,
    }
  }

  private async wait() {
    return new Promise((resolve) => setTimeout(resolve, this.data.delay))
  }

  // --- 1. SEARCH ---
// --- 1. SEARCH ---
  async search<T = any>(opts: SearchRequestOptions): Promise<SearchResult<T>> {
    await this.wait()
    console.log('[Mock] Search Options:', opts)

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
      if (!opts.mapping) return item
      const newItem: any = { ...item }
      Object.entries(opts.mapping).forEach(([spKey, friendly]) => {
        newItem[friendly] = item[spKey]
      })
      if (!newItem.url) newItem.url = item.Path
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
    // Simulate creating if not found
    return {
      Id: 999,
      Title: 'Ensured User',
      Email: loginName,
      LoginName: loginName,
    }
  }

  async getUserGroups(email?: string): Promise<SiteGroup[]> {
    await this.wait()
    const targetEmail = email || this.data.currentUser!.Email
    return this.data.userGroups?.[targetEmail] || []
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
    console.log(
      `[Mock] 📧 Sent Email to [${props.To.join(', ')}]: ${props.Subject}`
    )
  }

  // --- 3. FILES & VERSIONS ---

  async getFileVersions(url: string): Promise<FileVersion[]> {
    await this.wait()
    return this.data.fileVersions?.[url] || []
  }

  async getVersionHistoryLink(url: string): Promise<string> {
    return `#mock-history/${url}`
  }

  async uploadFile(
    url: string,
    name: string,
    file: Blob | ArrayBuffer
  ): Promise<string> {
    await this.wait()
    const fullPath = `${url}/${name}`
    this.data.files![fullPath] =
      file instanceof ArrayBuffer ? new Blob([file]) : file
    console.log(`[Mock] Uploaded ${fullPath}`)
    return fullPath
  }

  async downloadFile(url: string): Promise<Blob> {
    await this.wait()
    const f = this.data.files![url]
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
    console.log(`[Mock] Update ${list} #${id}`, payload)
  }
  async deleteListItem(list: string, id: number) {
    console.log(`[Mock] Delete ${list} #${id}`)
  }
  async getListItemById(list: string, id: number) {
    return this.data.lists![list]?.find((i) => i.Id === id)
  }
  async updateFileMetadata(url: string, payload: any) {
    console.log(`[Mock] Update Meta ${url}`, payload)
  }
  async deleteFile(url: string) {
    delete this.data.files![url]
  }
  async createFolder(url: string) {
    console.log(`[Mock] Create Folder ${url}`)
  }
  async restoreItem(id: string | number) {
    console.log(`[Mock] Restore ${id}`)
  }
  async getListFields() {
    return []
  }
  async getFieldChoices() {
    return []
  }
}
