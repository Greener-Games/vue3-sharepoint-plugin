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
  TermSetId?: string
}

export interface ListItemQueryOptions {
  select?: string[]
  expand?: string[]
  filter?: string
  top?: number
  orderBy?: string
  ascending?: boolean
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
  refiners?: string[]
  /**
   * Fields to expand in the hydration step (e.g. ['Author', 'Department']).
   * Presence of this field automatically triggers hydration.
   */
  expandFields?: string[]
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
  /**
   * Defines the sort order for the search results.
   * Array of objects specifying property and direction.
   */
  sortList?: Array<{
    property: string
    direction: 'ascending' | 'descending'
  }>
}

export interface SearchResult<T> {
  items: T[]
  totalHits: number
  startRow: number
  refiners?: RefinerResult[]
}

export interface RefinerValue {
  label: string
  count: number
  token: string
}

export interface RefinerResult {
  name: string
  values: RefinerValue[]
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
export const PermissionKind = {
  EmptyMask: 0,
  ViewListItems: 1,
  AddListItems: 2,
  EditListItems: 3,
  DeleteListItems: 4,
  ApproveItems: 5,
  OpenItems: 6,
  ViewVersions: 7,
  DeleteVersions: 8,
  CancelCheckout: 9,
  ManagePersonalViews: 10,
  ManageLists: 12,
  ViewFormPages: 13,
  Open: 17,
  ViewPages: 18,
  ManagePermissions: 19, // "Full Control" often implies this
  BrowseDirectories: 24,
  BrowseUserInfo: 25,
  AddDelPrivateWebParts: 26,
  UpdatePersonalWebParts: 27,
  ManageWeb: 28, // Admin
  UseClientIntegration: 37,
  UseRemoteAPIs: 38,
  ManageAlerts: 39,
  CreateAlerts: 40,
  EditMyUserInfo: 41,
  EnumeratePermissions: 63
} as const;

export type PermissionKind = typeof PermissionKind[keyof typeof PermissionKind];

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
  search<T = any>(options: SearchRequestOptions, abortSignal?: AbortSignal): Promise<SearchResult<T>>
  executeBatch(builder: (batch: IBatch) => void, abortSignal?: AbortSignal): Promise<void>
  createListItem<T = any>(
    listTitle: string,
    payload: Record<string, any>,
    abortSignal?: AbortSignal
  ): Promise<T>
  updateListItem(
    listTitle: string,
    id: number,
    payload: Record<string, any>,
    abortSignal?: AbortSignal
  ): Promise<void>
  deleteListItem(listTitle: string, id: number, abortSignal?: AbortSignal): Promise<void>
  getListItemById<T = any>(
    listTitle: string,
    id: number,
    select?: string[],
    expand?: string[],
    abortSignal?: AbortSignal
  ): Promise<T>

  getListItems<T = any>(
    listTitle: string,
    options?: ListItemQueryOptions,
    abortSignal?: AbortSignal
  ): Promise<T[]>

  // Attachments
  getItemAttachments(
    listTitle: string,
    itemId: number,
    abortSignal?: AbortSignal
  ): Promise<AttachmentInfo[]>
  addAttachment(
    listTitle: string,
    itemId: number,
    fileName: string,
    file: Blob | ArrayBuffer,
    abortSignal?: AbortSignal
  ): Promise<void>
  deleteAttachment(
    listTitle: string,
    itemId: number,
    fileName: string,
    abortSignal?: AbortSignal
  ): Promise<void>

  // Files & Folders
  uploadFile(
    serverRelativeUrl: string,
    fileName: string,
    file: Blob | ArrayBuffer,
    abortSignal?: AbortSignal
  ): Promise<string>
  downloadFile(serverRelativeUrl: string, abortSignal?: AbortSignal): Promise<Blob>
  updateFileMetadata(
    serverRelativeUrl: string,
    payload: Record<string, any>,
    abortSignal?: AbortSignal
  ): Promise<void>
  deleteFile(serverRelativeUrl: string, abortSignal?: AbortSignal): Promise<void>
  createFolder(serverRelativeUrl: string, abortSignal?: AbortSignal): Promise<void>

  // Webs
  getWebInfo(abortSignal?: AbortSignal): Promise<WebInfo>
  getSubwebs(abortSignal?: AbortSignal): Promise<WebInfo[]>

  // Lists
  getLists(abortSignal?: AbortSignal): Promise<ListInfo[]>
  getList(listTitle: string, abortSignal?: AbortSignal): Promise<ListInfo>
  createList(
    title: string,
    description?: string,
    template?: number,
    abortSignal?: AbortSignal
  ): Promise<ListInfo>
  deleteList(title: string, abortSignal?: AbortSignal): Promise<void>

  // Users & Groups
  getCurrentUser(abortSignal?: AbortSignal): Promise<UserInfo>
  getListFields(listTitle: string, webUrl?: string, abortSignal?: AbortSignal): Promise<FieldDefinition[]>
  getFieldChoices(
    listTitle: string,
    fieldInternalName: string,
    abortSignal?: AbortSignal
  ): Promise<string[]>

  /**
   * Search for a Taxonomy Term by Label in a specific Term Set.
   * Returns the first match with Label and GUID.
   */
  searchTerm(
    termSetId: string,
    label: string,
    abortSignal?: AbortSignal
  ): Promise<{ Label: string; TermGuid: string } | null>

  /** Get all users on the site */
  getSiteUsers(abortSignal?: AbortSignal): Promise<UserInfo[]>
  /** Search for users by name or email */
  searchUsers(query: string, abortSignal?: AbortSignal): Promise<UserInfo[]>
  ensureUser(loginName: string, abortSignal?: AbortSignal): Promise<UserInfo>

  /**
   * Get the SharePoint Groups a user belongs to.
   * If email is omitted, returns groups for the Current User.
   */
  getUserGroups(email?: string, abortSignal?: AbortSignal): Promise<SiteGroup[]>
  addUserToGroup(groupName: string, loginName: string, abortSignal?: AbortSignal): Promise<void>
  removeUserFromGroup(groupName: string, loginName: string, abortSignal?: AbortSignal): Promise<void>
  createGroup(groupName: string, description?: string, abortSignal?: AbortSignal): Promise<SiteGroup>

  /**
   * Get the effective permission mask for a user.
   * If email is omitted, returns permissions for the Current User.
   */
  getUserEffectivePermissions(email?: string, abortSignal?: AbortSignal): Promise<SPBasePermissions>

  /**
   * Fetch previous versions of a file.
   * Note: This usually returns *past* versions. The current version is the live file.
   */
  getFileVersions(serverRelativeUrl: string, abortSignal?: AbortSignal): Promise<FileVersion[]>

  /**
   * Generates a link to the native SharePoint Version History page.
   * Requires a server call to resolve List ID and Item ID.
   */
  getVersionHistoryLink(serverRelativeUrl: string, abortSignal?: AbortSignal): Promise<string>

  /**
   * Generates a link to the native SharePoint Version History page.
   * Uses provided List ID and Item ID to construct the link synchronously.
   */
  getVersionHistoryLinkByItem(
    listId: string,
    itemId: number,
    webUrl?: string
  ): string
}

export interface SharePointConfig {
  baseUrl: string
  /**
   * Optional: Dev endpoint to use when running on localhost.
   * If provided, the client will automatically switch to this URL
   * when location.hostname is 'localhost' or '127.0.0.1'.
   */
  devBaseUrl?: string
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
