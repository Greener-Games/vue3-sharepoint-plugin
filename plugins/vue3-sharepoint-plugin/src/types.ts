export interface UserInfo {
  Id: number
  Title: string
  Email: string
  LoginName: string
}

export interface FieldDefinition {
  InternalName: string
  Title: string
  TypeAsString: string
  Hidden: boolean
  Choices?: string[]
}

export type FilterValue =
  | string
  | number
  | boolean
  | Array<string | number | boolean>

export interface SearchRequestOptions {
  query?: string
  searchTitleOnly?: boolean
  scope?: string | string[]
  fileTypes?: {
    include?: string[]
    exclude?: string[]
  }
  filters?: Record<string, FilterValue>
  rowLimit?: number
  startRow?: number
  mapping?: Record<string, string>
  selectFields?: string[]
  /**
   * Filter search results by type.
   * - 'items': Returns files and list items (excludes folders).
   * - 'folders': Returns only folders.
   * - 'all': Returns everything (default).
   */
  resultType?: 'items' | 'folders' | 'all'
  /**
   * If true, includes a 'relativePath' field in the returned items.
   * Calculated from the 'Path' managed property.
   */
  includeRelativePath?: boolean
}

export interface SearchResult<T> {
  items: T[]
  totalHits: number
  startRow: number
}

export interface SiteGroup {
  Id: number
  Title: string
  Description?: string
  LoginName?: string
}

export interface SPBasePermissions {
  High: number
  Low: number
}

// Helper Enum for common permission checks (optional but recommended)
export enum PermissionKind {
  EmptyMask = 0,
  ViewListItems = 1,
  AddListItems = 2,
  EditListItems = 3,
  DeleteListItems = 4,
  ApproveItems = 5,
  OpenItems = 6,
  ViewVersions = 7,
  DeleteVersions = 8,
  CancelCheckout = 9,
  ManagePersonalViews = 10,
  ManageLists = 12,
  ViewFormPages = 13,
  Open = 17,
  ViewPages = 18,
  ManagePermissions = 19, // "Full Control" often implies this
  BrowseDirectories = 24,
  BrowseUserInfo = 25,
  AddDelPrivateWebParts = 26,
  UpdatePersonalWebParts = 27,
  ManageWeb = 28, // Admin
  UseClientIntegration = 37,
  UseRemoteAPIs = 38,
  ManageAlerts = 39,
  CreateAlerts = 40,
  EditMyUserInfo = 41,
  EnumeratePermissions = 63
}

export interface FileVersion {
  VersionLabel: string
  Created: string
  CheckInComment: string
  IsCurrentVersion: boolean
  Size: number
  Url: string
  CreatedBy: Partial<UserInfo>
}

export interface EmailProperties {
  To: string[]
  CC?: string[]
  Subject: string
  Body: string
}

export interface WebInfo {
  Title: string
  Url: string
  Id: string
  Description: string
}

export interface ListInfo {
  Id: string
  Title: string
  Description: string
  ItemCount: number
  Hidden: boolean
  ImageUrl: string
}

export interface AttachmentInfo {
  FileName: string
  ServerRelativeUrl: string
}

// Builder Interface
export interface IBatch {
  createListItem(listTitle: string, payload: Record<string, any>): void
  updateListItem(
    listTitle: string,
    id: number,
    payload: Record<string, any>
  ): void
  deleteListItem(listTitle: string, id: number): void
  deleteFile(serverRelativeUrl: string): void
}

export interface ISharePointClient {
  search<T = any>(options: SearchRequestOptions): Promise<SearchResult<T>>
  executeBatch(builder: (batch: IBatch) => void): Promise<void>
  createListItem<T = any>(
    listTitle: string,
    payload: Record<string, any>
  ): Promise<T>
  updateListItem(
    listTitle: string,
    id: number,
    payload: Record<string, any>
  ): Promise<void>
  deleteListItem(listTitle: string, id: number): Promise<void>
  getListItemById<T = any>(
    listTitle: string,
    id: number,
    select?: string[]
  ): Promise<T>

  // Attachments
  getItemAttachments(listTitle: string, itemId: number): Promise<AttachmentInfo[]>
  addAttachment(listTitle: string, itemId: number, fileName: string, file: Blob | ArrayBuffer): Promise<void>
  deleteAttachment(listTitle: string, itemId: number, fileName: string): Promise<void>

  // Files & Folders
  uploadFile(
    serverRelativeUrl: string,
    fileName: string,
    file: Blob | ArrayBuffer
  ): Promise<string>
  downloadFile(serverRelativeUrl: string): Promise<Blob>
  updateFileMetadata(
    serverRelativeUrl: string,
    payload: Record<string, any>
  ): Promise<void>
  deleteFile(serverRelativeUrl: string): Promise<void>
  createFolder(serverRelativeUrl: string): Promise<void>

  // Webs
  getWebInfo(): Promise<WebInfo>
  getSubwebs(): Promise<WebInfo[]>

  // Lists
  getLists(): Promise<ListInfo[]>
  getList(listTitle: string): Promise<ListInfo>
  createList(title: string, description?: string, template?: number): Promise<ListInfo>
  deleteList(title: string): Promise<void>

  // Users & Groups
  getCurrentUser(): Promise<UserInfo>
  getListFields(listTitle: string): Promise<FieldDefinition[]>
  getFieldChoices(
    listTitle: string,
    fieldInternalName: string
  ): Promise<string[]>

  /** Get all users on the site */
  getSiteUsers(): Promise<UserInfo[]>
  ensureUser(loginName: string): Promise<UserInfo>

  /**
   * Get the SharePoint Groups a user belongs to.
   * If email is omitted, returns groups for the Current User.
   */
  getUserGroups(email?: string): Promise<SiteGroup[]>
  addUserToGroup(groupName: string, loginName: string): Promise<void>
  removeUserFromGroup(groupName: string, loginName: string): Promise<void>
  createGroup(groupName: string, description?: string): Promise<SiteGroup>

  /**
   * Get the effective permission mask for a user.
   * If email is omitted, returns permissions for the Current User.
   */
  getUserEffectivePermissions(email?: string): Promise<SPBasePermissions>

  /**
   * Fetch previous versions of a file.
   * Note: This usually returns *past* versions. The current version is the live file.
   */
  getFileVersions(serverRelativeUrl: string): Promise<FileVersion[]>

  /**
   * Generates a link to the native SharePoint Version History page.
   */
  getVersionHistoryLink(serverRelativeUrl: string): string
}

export interface SharePointConfig {
  baseUrl: string
  authProvider?: () => Promise<Record<string, string>>
  /**
   * Enable caching for GET requests (User, Fields, Choices).
   * Search results are usually not cached to ensure freshness.
   */
  enableCache?: boolean
  /**
   * If true, enables verbose logging to the console for debugging purposes.
   */
  debug?: boolean
}
