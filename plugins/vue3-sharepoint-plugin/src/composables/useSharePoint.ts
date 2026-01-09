import { inject, ref } from 'vue'
import {
  type ISharePointClient,
  PermissionKind,
  type SPBasePermissions,
  type UserInfo,
} from '../types'
import { useSearch } from './useSearch'
import { SP_CLIENT_SYMBOL } from '../index'

let _cachedUser: UserInfo | null = null

export function useSharePoint() {
  const client = inject<ISharePointClient>(SP_CLIENT_SYMBOL)
  if (!client) throw new Error('SharePointPlugin not installed.')

  // --- User Caching ---
  const currentUser = ref<UserInfo | null>(_cachedUser)
  const loadUser = async () => {
    if (_cachedUser) {
      currentUser.value = _cachedUser
      return _cachedUser
    }
    const user = await client.getCurrentUser()
    _cachedUser = user
    currentUser.value = user
    return user
  }

  // --- File Download Helper ---
  const downloadAndSave = async (
    serverRelativeUrl: string,
    filename: string
  ) => {
    try {
      const blob = await client.downloadFile(serverRelativeUrl)
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.style.display = 'none'
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)
    } catch (e) {
      console.error('Download failed', e)
      throw e
    }
  }

  /**
   * Check if a specific permission bit exists in the High/Low mask.
   * This logic mimics SharePoint's bitwise check.
   */
  const hasPermission = (
    perms: SPBasePermissions,
    kind: PermissionKind
  ): boolean => {
    // 1. Handle Full Control (If High & Low are maxed out, return true immediately)
    // 2. Perform Bitwise check.
    // Note: JS Bitwise operators treat numbers as 32-bit signed integers.
    // SharePoint perms are 64-bit split into High/Low.

    // This is a simplified check. For production robustness with 64-bit:
    // PnP provides a helper class, but since we are generic, we do a basic check here.

    // Simple Logic:
    // If the permission kind is < 32, it's in the Low bits.
    // If the permission kind is >= 32, it's in the High bits.

    const bit = kind % 32
    const chunk = kind >= 32 ? perms.High : perms.Low

    // Check if bit is set
    return (chunk & (1 << bit)) !== 0
  }

  return {
    client, // Exposed for raw access if needed
    createSearch: () => useSearch(client),

    // Globals
    currentUser,
    loadUser,
    downloadAndSave,

    // Proxy Methods
    executeBatch: client.executeBatch.bind(client),

    // Items
    addItem: client.createListItem.bind(client),
    updateItem: client.updateListItem.bind(client),
    deleteItem: client.deleteListItem.bind(client),
    getItem: client.getListItemById.bind(client),
    getItems: client.getListItems.bind(client),

    // Attachments
    getAttachments: client.getItemAttachments.bind(client),
    addAttachment: client.addAttachment.bind(client),
    deleteAttachment: client.deleteAttachment.bind(client),

    // Files/Folders
    upload: client.uploadFile.bind(client),
    createFolder: client.createFolder.bind(client),
    updateFileMetadata: client.updateFileMetadata.bind(client),
    deleteFile: client.deleteFile.bind(client),

    // Webs
    getWebInfo: client.getWebInfo.bind(client),
    getSubwebs: client.getSubwebs.bind(client),

    // Lists
    getLists: client.getLists.bind(client),
    getList: client.getList.bind(client),
    createList: client.createList.bind(client),
    deleteList: client.deleteList.bind(client),

    // Fields
    getFields: client.getListFields.bind(client),
    getChoices: client.getFieldChoices.bind(client),

    // Users & Groups
    getSiteUsers: client.getSiteUsers.bind(client),
    searchUsers: client.searchUsers.bind(client), // Add this
    ensureUser: client.ensureUser.bind(client),
    getUserGroups: client.getUserGroups.bind(client),
    addUserToGroup: client.addUserToGroup.bind(client),
    removeUserFromGroup: client.removeUserFromGroup.bind(client),
    createGroup: client.createGroup.bind(client),
    getPermissions: client.getUserEffectivePermissions.bind(client),

    // Version History
    getVersions: client.getFileVersions.bind(client),
    historyLink: client.getVersionHistoryLink.bind(client),

    hasPermission,
    PermissionKind, // Expose Enum
  }
}
