<template>
  <div class="dashboard-container">
    <!-- HEADER -->
    <div class="header-section">
      <h1>{{ pageTitle }}</h1>
      <p class="subtitle">(b) Ultricies tortor, lectus eget libero sit sit donec purus. Molestie lectus etiam erat sit</p>
    </div>

    <!-- SEARCH BAR -->
    <div class="search-bar-wrapper">
      <div class="search-input-group">
        <span class="search-icon">
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#5b4bf5" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
        </span>
        <input
            type="text"
            v-model="searchQuery"
            @keydown.enter="performSearch"
            placeholder="Search for Content"
            class="main-search-input"
        />
      </div>
    </div>

    <!-- RESULTS HEADER & STATS -->
    <div class="results-meta">
      <h3>Results</h3>
      <div class="meta-stats">
        <span class="stat-item">
          <span class="icon">☰</span> {{ totalHits }} Results Found
        </span>
        <span class="stat-item warning" v-if="actionCount > 0">
          <span class="icon">⚠</span> {{ actionCount }} Actions Required
        </span>
      </div>
    </div>

    <!-- LOADING STATE -->
    <div v-if="loading" class="loading-state">
      Searching...
    </div>

    <!-- ERROR STATE -->
    <div v-if="error" class="error-state">
      {{ error }}
    </div>

    <!-- TABLE -->
    <div class="table-container" v-if="!loading && formattedResults.length > 0">
      <table>
        <thead>
        <tr>
          <th class="col-file">File <span class="sort-carat">⌄</span></th>
          <th class="col-type">Type <span class="sort-carat">⌄</span></th>
          <th class="col-lang">Language <span class="sort-carat">⌄</span></th>
          <th class="col-owner">Owner <span class="sort-carat">⌄</span></th>
          <th class="col-status">Status <span class="sort-carat">⌄</span></th>
          <th class="col-date">Published <span class="sort-carat">⌄</span></th>
          <th class="col-action">Action <span class="sort-carat">⌄</span></th>
        </tr>
        </thead>
        <tbody>
        <tr v-for="(row, index) in formattedResults" :key="index">
          <!-- File -->
          <td>
            <div class="file-cell">
              <span class="paperclip">🔗</span>
              <a :href="row.path" target="_blank" class="file-link">{{ row.title }}</a>
            </div>
          </td>
          <td>{{ row.type || 'Generic' }}</td>
          <td>{{ row.language || 'English' }}</td>
          <td>{{ row.owner }}</td>
          <td><span class="status-text">{{ row.status }}</span></td>
          <td>{{ row.created }}</td>
          <td>
            <button
                v-if="viewMode === 'myteam' && row.status === 'Pending'"
                class="action-btn"
                @click="handleAction(row)"
            >
              Review / Approve
            </button>
          </td>
        </tr>
        </tbody>
      </table>
    </div>

    <!-- EMPTY STATE -->
    <div v-if="!loading && !error && formattedResults.length === 0" class="empty-state">
      No documents found.
    </div>

    <!-- PAGINATION -->
    <div class="pagination-controls" v-if="totalHits > 20">
      <button :disabled="!hasPrev" @click="prevPage">Previous</button>
      <span>Page {{ pageInfo.current }} of {{ pageInfo.total }}</span>
      <button :disabled="!hasNext" @click="nextPage">Next</button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed, watch } from 'vue'
import { useSharePoint } from '@greener-games/vue3-sharepoint-plugin'
import type { SearchRequestOptions } from '@greener-games/vue3-sharepoint-plugin'

// --- Props ---
const props = defineProps<{
  viewMode: 'mydocuments' | 'myteam'
}>()

// --- State ---
const { createSearch, loadUser, currentUser } = useSharePoint()
const {
  execute,
  results,
  loading,
  error,
  totalHits,
  nextPage,
  prevPage,
  hasPrev,
  hasNext,
  pageInfo
} = createSearch()

const searchQuery = ref('')

// --- Computed ---
const pageTitle = computed(() => {
  return props.viewMode === 'myteam' ? 'My Team' : 'My Documents'
})

const formattedResults = computed(() => {
  if (!results.value || !results.value.items) return []

  return results.value.items.map((item: any) => {
    // Mapping Search Results (Managed Properties) to UI
    return {
      title: item.Title || item.Filename || 'Untitled',
      path: item.Path,
      type: item.DocumentType || item.FileExtension || 'Doc',
      language: item.DocLanguage || 'English',
      owner: item.DocOwner || item.Author || 'Unknown',
      status: item.DocStatus || 'Published',
      created: item.Created ? new Date(item.Created).toLocaleDateString('en-GB').replace(/\//g, '-') : '00-00-0000',
    }
  })
})

const actionCount = computed(() => {
  return formattedResults.value.filter(r => r.status === 'Pending').length
})

// --- Logic ---
const performSearch = async () => {
  if (!currentUser.value) await loadUser()
  const userEmail = currentUser.value?.Email

  if (!userEmail) {
    console.error("User not found")
    return
  }

  // Define our Filters Object
  const filters: Record<string, string> = {}

  // Apply logic based on viewMode
  if (props.viewMode === 'mydocuments') {
    // "Author" is a standard Managed Property in SharePoint
    filters.Author = userEmail
  }
  else if (props.viewMode === 'myteam') {
    // "LineManager" MUST be mapped to a Managed Property in SharePoint Admin Center
    // even if it exists as a column on the list.
    filters.LineManager = userEmail
  }

  // Execute Search
  const options: SearchRequestOptions = {
    // 1. The free text query (e.g., "Invoice")
    query: searchQuery.value,

    // 2. The structured filters (handled by your Client's buildKql)
    filters: filters,

    // 3. Properties to select
    selectFields: [
      'Title',
      'Path',
      'Created',
      'FileExtension',
      'Author',
      'DocumentType',
      'DocLanguage',
      'DocOwner',
      'DocStatus'
    ],

    rowLimit: 20,
    sortList: [{ property: 'Created', direction: 'ascending' }]
  }

  await execute(options)
}

const handleAction = (row: any) => {
  alert(`Opening review dialog for: ${row.title}`)
}

// --- Lifecycle ---
onMounted(() => {
  loadUser().then(() => {
    performSearch()
  })
})

watch(() => props.viewMode, () => {
  searchQuery.value = ''
  performSearch()
})
</script>

<style scoped>
/* GENERAL LAYOUT */
.dashboard-container {
  max-width: 1200px;
  margin: 0 auto;
  font-family: 'Segoe UI', sans-serif;
  color: #333;
  padding: 20px;
  background: white;
}

/* HEADER */
.header-section { margin-bottom: 25px; }
.header-section h1 { font-size: 28px; font-weight: 400; margin: 0 0 8px 0; color: #000; }
.subtitle { font-size: 13px; color: #666; margin: 0; }

/* SEARCH BAR */
.search-bar-wrapper { margin-bottom: 30px; border-bottom: 1px dotted #ccc; padding-bottom: 25px; }
.search-input-group { position: relative; width: 100%; }
.search-icon { position: absolute; left: 15px; top: 50%; transform: translateY(-50%); display: flex; align-items: center; }
.main-search-input { width: 100%; padding: 15px 15px 15px 50px; border: 1px solid #dbeafe; border-radius: 6px; font-size: 16px; color: #3b49df; outline: none; background-color: #fff; box-shadow: 0 2px 4px rgba(0,0,0,0.02); transition: border 0.2s; }
.main-search-input:focus { border-color: #5b4bf5; }
.main-search-input::placeholder { color: #3b49df; opacity: 0.8; }

/* RESULTS META */
.results-meta { display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 15px; }
.results-meta h3 { font-size: 18px; margin: 0; font-weight: 600; color: #333; }
.meta-stats { display: flex; gap: 20px; font-size: 13px; color: #555; }
.stat-item { display: flex; align-items: center; gap: 5px; }
.stat-item.warning { color: #444; }

/* TABLE */
.table-container { border: 1px solid #eef2ff; border-radius: 8px; overflow: hidden; }
table { width: 100%; border-collapse: collapse; font-size: 14px; }
thead { background-color: #eff4ff; }
th { text-align: left; padding: 14px 15px; color: #5b4bf5; font-weight: 600; font-size: 13px; white-space: nowrap; border-bottom: 1px solid #e0e7ff; cursor: pointer; }
th:hover { background-color: #e0eaff; }
.sort-carat { font-size: 10px; margin-left: 4px; }
td { padding: 12px 15px; border-bottom: 1px solid #f0f4f8; vertical-align: middle; }
tbody tr:nth-child(even) { background-color: #fcfdfe; }
tbody tr:hover { background-color: #f8fafc; }

/* CELL STYLES */
.file-cell { display: flex; align-items: center; gap: 8px; }
.paperclip { color: #5b4bf5; font-size: 16px; }
.file-link { color: #5b4bf5; text-decoration: underline; font-weight: 500; }
.status-text { color: #333; }

/* ACTION BUTTON */
.action-btn { background-color: #5b4bf5; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-size: 12px; font-weight: 600; cursor: pointer; white-space: nowrap; box-shadow: 0 2px 4px rgba(91, 75, 245, 0.3); transition: background 0.2s; }
.action-btn:hover { background-color: #4a3ac9; }

/* PAGINATION */
.pagination-controls { margin-top: 20px; display: flex; justify-content: center; align-items: center; gap: 15px; }
.pagination-controls button { padding: 6px 12px; background: #f3f4f6; border: 1px solid #ddd; border-radius: 4px; cursor: pointer; }
.pagination-controls button:disabled { opacity: 0.5; cursor: not-allowed; }

/* LOADING / EMPTY */
.loading-state, .empty-state, .error-state { padding: 40px; text-align: center; color: #666; font-style: italic; }
.error-state { color: red; }
</style>
