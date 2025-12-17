# Vue 3 SharePoint Plugin

A robust, type-safe wrapper for SharePoint operations in Vue 3 applications. This plugin provides a unified interface for Search, CRUD, File operations, and Batching, supporting both **PnPjs** (Production) and an in-memory **Mock Engine** (Development/Testing).

## 📦 Installation

### 1. Install Dependencies
The plugin relies on PnPjs for the production implementation.
```bash
npm install @pnp/sp @pnp/logging @pnp/queryable
```

### 2. Link the Plugin (Local Development)
If you are developing the plugin locally alongside your app, update your `vite.config.ts` alias or install it via file path:
```bash
npm install ./plugins/vue3-sharepoint-plugin
```

---

## ⚙️ Configuration (main.ts)

Register the plugin in your main entry file (`main.ts`). You can switch between Real and Mock clients based on the environment.

### Option A: Quick Mock Setup (Recommended)
Use the built-in `createMockData` factory to instantly generate users, groups, and hundreds of dummy documents with rich metadata.

```typescript
import { createApp } from 'vue'
import App from './App.vue'
import {
  SharePointPlugin,
  PnPSharePointClient,
  MockSharePointClient,
  createMockData
} from 'vue3-sharepoint-plugin'

const app = createApp(App)
const isDev = import.meta.env.DEV

let client

if (isDev) {
  // --- DEVELOPMENT: Use Generated Mock Data ---
  // Generates 100 documents located at the specified site URL
  const mockData = createMockData(100, '/sites/Intranet')

  console.log('🚧 Running with Mock Data', mockData)
  client = new MockSharePointClient(mockData)
} else {
  // --- PRODUCTION: Use Real PnPjs ---
  client = new PnPSharePointClient({
    baseUrl: 'https://your-tenant.sharepoint.com/sites/Intranet',
    enableCache: true, // Caches User Info & List Fields in SessionStorage
  })
}

app.use(SharePointPlugin, { client })
app.mount('#app')
```

### Option B: Custom Mock Setup (Advanced)
If you need specific testing scenarios, you can define the `MockData` object manually or use the exposed helper functions.

```typescript
import { 
  MockSharePointClient, 
  createMockUser, 
  type MockData 
} from 'vue3-sharepoint-plugin'

const customMock: MockData = {
  currentUser: createMockUser(1, 'Tester', 'test@local'),
  // Define specific lists
  lists: {
    '/sites/dev/Lists/Projects': [
      { Id: 1, Title: 'Project Alpha', Status: 'Active' },
      { Id: 2, Title: 'Project Beta', Status: 'Closed' }
    ]
  },
  delay: 1000 // Simulate slow network (1s latency)
}

const client = new MockSharePointClient(customMock)
```

### 📘 Mock Factory API Reference

These utilities help generate data for the `MockSharePointClient`.

**`createMockData(docCount, baseUrl)`**
| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `docCount` | `number` | `60` | Number of documents to generate. |
| `baseUrl` | `string` | `'/sites/Intranet'` | The base URL for the simulated site. |

**`createMockUser(id, name, email)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `id` | `number` | Unique ID for the user. |
| `name` | `string` | Display name. |
| `email` | `string` | User email (used for login name generation). |

---

## 🔍 Search Engine

The search engine uses a **Factory Pattern**. You create an isolated search instance for every search interface.

### Usage Example

```typescript
<script setup lang="ts">
import { useSharePoint } from 'vue3-sharepoint-plugin'

// 1. Get the factory
const { createSearch } = useSharePoint()

// 2. Create an isolated instance
const searchCtx = createSearch()

const handleSearch = () => {
  searchCtx.execute({
    query: 'Engineering', 
    rowLimit: 10,
    scope: 'Shared Documents', // Automatically inferred as relative to baseUrl
    resultType: 'items', // Exclude folders
    includeRelativePath: true, // Returns a 'relativePath' field
    fileTypes: {
      include: ['docx', 'pdf'],
      exclude: ['aspx']
    },
    filters: {
      'RefinableString01': ['UK', 'US'], // Array = OR Logic
    },
    mapping: {
      'Title': 'title',
      'Path': 'url',
      'RefinableString01': 'region'
    }
  })
}
</script>
```

### 📘 Search API Reference

**`searchCtx.execute(options)`**
Performs a search query. This updates `searchCtx.results`.

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `query` | `string` | No | `'*'` | The text to search for. |
| `rowLimit` | `number` | No | `10` | Number of items to return per page. |
| `startRow` | `number` | No | `0` | The offset index (used for pagination). |
| `scope` | `string` \| `string[]` | No | - | Limit search to specific paths. <br>• **Site-Relative:** `'Shared Documents'` (appended to baseUrl)<br>• **Server-Relative:** `'/sites/Other/Docs'` (appended to origin)<br>• **Absolute:** `'https://...'` |
| `resultType` | `'items'` \| `'folders'` \| `'all'` | No | `'all'` | Filter results by type. `'items'` excludes folders. |
| `includeRelativePath` | `boolean` | No | `false` | If true, adds `relativePath` property to results (derived from `Path`). |
| `fileTypes` | `{ include?: string[], exclude?: string[] }` | No | - | Filter by file extension. e.g. `{ include: ['pdf'] }`. |
| `filters` | `Record<string, any>` | No | - | Faceted filters (Array=OR, String=AND). |
| `mapping` | `Record<string, string>` | No | - | Maps internal SharePoint Managed Properties to friendly names. |
| `selectFields` | `string[]` | No | - | Specific fields to request from API to reduce payload size. |
| `searchTitleOnly`| `boolean`| No | `false` | If `true`, searches `Title` property only. |

---

## 📝 Lists & Items (CRUD)

All CRUD methods are available directly from `useSharePoint()`.

### Usage Example
```typescript
import { useSharePoint } from 'vue3-sharepoint-plugin'

const { addItem, updateItem, deleteItem } = useSharePoint()

const newItem = await addItem('Tasks', { Title: 'New Task' })
await updateItem('Tasks', newItem.Id, { Status: 'Active' })
```

### 📘 List API Reference

**`addItem(listTitle, payload)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `listTitle` | `string` | The display name of the list (e.g., 'Tasks'). |
| `payload` | `Record<string, any>` | Object matching list columns (e.g., `{ Title: 'A' }`). |

**`updateItem(listTitle, id, payload)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `listTitle` | `string` | The display name of the list. |
| `id` | `number` | The integer ID of the item. |
| `payload` | `Record<string, any>` | The fields to change (Merge update). |

**`deleteItem(listTitle, id)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `listTitle` | `string` | The display name of the list. |
| `id` | `number` | The integer ID of the item to recycle. |

**`getItem(listTitle, id, select?)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `listTitle` | `string` | The display name of the list. |
| `id` | `number` | The integer ID of the item. |
| `select` | `string[]` | (Optional) Array of specific column names to fetch. |

---

## 📂 Files & Folders

Helper methods for Document Libraries.

### Usage Example
```typescript
const { upload, downloadAndSave, createFolder } = useSharePoint()

await createFolder('/sites/Dev/Docs/NewFolder')
await upload('/sites/Dev/Docs', 'report.pdf', fileBlob)
```

### 📘 Files API Reference

**`upload(serverRelativeUrl, fileName, file)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `serverRelativeUrl` | `string` | The target folder path. |
| `fileName` | `string` | The name of the file (e.g., `report.pdf`). |
| `file` | `Blob` \| `ArrayBuffer` | The file content. |

**`createFolder(serverRelativeUrl)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `serverRelativeUrl` | `string` | The full path where the folder should be created. |

**`updateFileMetadata(serverRelativeUrl, payload)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `serverRelativeUrl` | `string` | The full path to the file. |
| `payload` | `Record<string, any>` | Key-value pairs of metadata to update. |

**`getVersions(serverRelativeUrl)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `serverRelativeUrl` | `string` | The full path to the file. |

**`downloadAndSave(serverRelativeUrl, fileName)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `serverRelativeUrl` | `string` | The full path to the file. |
| `fileName` | `string` | The name the file will be saved as. |

---

## 🔐 Permissions & Users

Utilities for checking permissions and retrieving user info.

### Usage Example
```typescript
const { getPermissions, hasPermission, PermissionKind } = useSharePoint()

const perms = await getPermissions()
if (hasPermission(perms, PermissionKind.ManageWeb)) {
  console.log('User is Admin')
}
```

### 📘 Permissions API Reference

**`getSiteUsers()`**
Returns all users from the User Information List of the site.
*   **Returns:** `Promise<UserInfo[]>`

**`getUserGroups(email?)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `email` | `string` | (Optional) Email of the user. Defaults to current user if omitted. |

**`getPermissions(email?)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `email` | `string` | (Optional) Email of the user. Defaults to current user if omitted. |

**`hasPermission(permissions, kind)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `permissions` | `SPBasePermissions` | The object returned from `getPermissions()`. |
| `kind` | `PermissionKind` | The permission bit to check (e.g. `PermissionKind.EditListItems`). |

---

## 🚀 Batching

Execute multiple **write** operations in a single HTTP request.

### Usage Example
```typescript
const { executeBatch } = useSharePoint()

await executeBatch((batch) => {
  batch.createListItem('Tasks', { Title: 'A' })
  batch.deleteFile('/sites/dev/docs/old.pdf')
})
```

### 📘 Batch API Reference

**`executeBatch(callback)`**
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `callback` | `(batch: IBatch) => void` | A function that receives a `batch` builder. |

**Methods available on `batch` object:**

| Method | Parameters | Description |
| :--- | :--- | :--- |
| `createListItem` | `listTitle`, `payload` | Adds an item to a batch. |
| `updateListItem` | `listTitle`, `id`, `payload` | Adds an update to a batch. |
| `deleteListItem` | `listTitle`, `id` | Adds a delete to a batch. |
| `deleteFile` | `serverRelativeUrl` | Adds a file recycle to a batch. |

---

## 🛠 Utilities

### 📘 Utilities API Reference

**`loadUser()`**
Returns the current logged-in user. Caches the result to avoid repeated API calls.
*   **Returns:** `Promise<UserInfo>` `{ Id, Title, Email, LoginName }`

**`getFields(listTitle)`**

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `listTitle` | `string` | The display name of the list. |

**`getChoices(listTitle, fieldName)`**

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `listTitle` | `string` | The display name of the list. |
| `fieldName` | `string` | The internal name or title of the choice column. |
