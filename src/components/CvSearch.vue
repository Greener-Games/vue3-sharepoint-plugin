<template>
  <div class="page-container">
    <!-- STICKY SEARCH HERO -->
    <div class="search-hero">
      <div class="search-content">
        <!-- Search Input Group -->
        <div class="search-input-group">
          <div class="input-wrapper">
            <svg
              class="search-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              stroke-width="2">
              <circle cx="11" cy="11" r="8"></circle>
              <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
            </svg>

            <!-- v-model binds to typing state -->
            <input
              v-model="searchQuery"
              type="text"
              placeholder="Search for Content"
              class="hero-input"
              @keyup.enter="triggerSearch" />

            <!-- Clear Button -->
            <button
              v-if="searchQuery"
              class="clear-text-btn"
              title="Clear search"
              @click="clearSearchText">
              ✕
            </button>
          </div>

          <!-- Manual Search Button -->
          <button
            class="search-btn"
            :disabled="isSearching"
            @click="triggerSearch">
            {{ isSearching ? 'Searching...' : 'Search' }}
          </button>
        </div>
      </div>
    </div>

    <!-- MAIN GRID -->
    <div class="main-layout">
      <!-- LEFT: RESULTS -->
      <div class="results-column">
        <div class="results-header">
          <h2>Results</h2>
          <div class="results-count">
            <span class="icon">☰</span>
            <!-- Updated to use TotalHits from server -->
            {{ totalHits }} Results
          </div>
        </div>

        <div class="results-wrapper">
          <div v-if="isSearching" class="loading-overlay">
            <div class="spinner"></div>
          </div>

          <div class="results-list" :class="{ 'is-loading': isSearching }">
            <div
              v-if="resultsList.length === 0 && !isSearching"
              class="no-results">
              No documents found.
            </div>

            <div v-for="doc in resultsList" :key="doc.id" class="result-item">
              <div class="doc-icon-wrapper">
                <svg
                  class="doc-icon"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  stroke-width="1.5">
                  <path
                    d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                  <polyline points="14 2 14 8 20 8"></polyline>
                  <line x1="16" y1="13" x2="8" y2="13"></line>
                  <line x1="16" y1="17" x2="8" y2="17"></line>
                  <line x1="10" y1="9" x2="8" y2="9"></line>
                </svg>
              </div>

              <div class="doc-content">
                <a :href="doc.url" target="_blank" class="doc-title">{{
                  doc.title
                }}</a>

                <!-- CONTENT MATCH SNIPPET -->
                <!-- Only visible if a search has been EXECUTED and is not empty -->
                <div v-if="hasSearchTerm && doc.content" class="search-snippet">
                  <div class="snippet-label">...found inside document:</div>
                  <p v-html="formatSummary(doc.content || '')"></p>
                </div>

                <div class="meta-row">
                  <span class="meta-label">Published:</span>
                  <span class="meta-value">{{ doc.date || 'N/A' }}</span>
                </div>

                <div class="meta-grid">
                  <div>
                    <span class="meta-label">Region:</span> {{ doc.region }}
                  </div>
                  <div>
                    <span class="meta-label">Division:</span> {{ doc.division }}
                  </div>
                  <div>
                    <span class="meta-label">Practice:</span> {{ doc.practice }}
                  </div>
                  <div class="full-width">
                    <span class="meta-label">Capabilities:</span>
                    {{ doc.capabilities }}
                  </div>
                  <div class="full-width">
                    <span class="meta-label">Markets:</span> {{ doc.markets }}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- PAGINATION (Server Side) -->
        <div v-if="totalHits > 0" class="pagination-footer">
          <button :disabled="!hasPrevPage" @click="changePage(currentPage - 1)">
            Previous
          </button>
          <span>Page {{ currentPage }} of {{ totalPages }}</span>
          <button :disabled="!hasNextPage" @click="changePage(currentPage + 1)">
            Next
          </button>
        </div>
      </div>

      <!-- RIGHT: STICKY FILTERS -->
      <aside class="refiners-sidebar">
        <div class="sidebar-header">
          <div class="header-title">
            <svg
              width="20"
              height="20"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              stroke-width="2">
              <line x1="4" y1="21" x2="4" y2="14"></line>
              <line x1="4" y1="10" x2="4" y2="3"></line>
              <line x1="12" y1="21" x2="12" y2="12"></line>
              <line x1="12" y1="8" x2="12" y2="3"></line>
              <line x1="20" y1="21" x2="20" y2="16"></line>
              <line x1="20" y1="12" x2="20" y2="3"></line>
            </svg>
            <span>Refiners</span>
          </div>
          <button
            v-if="hasActiveFilters"
            class="clear-btn"
            @click="clearAllFilters">
            Clear all Filters
          </button>
        </div>

        <!-- 1. STANDARD FILTERS LOOP -->
        <div
          v-for="filter in standardFilters"
          :key="filter.key"
          class="filter-card">
          <div class="filter-card-header" @click="toggleSection(filter.key)">
            <span>{{ filter.title }}</span>
            <span class="chevron" :class="{ open: sections[filter.key] }"
              >›</span
            >
          </div>

          <!-- Summary Row (Closed + Active) -->
          <div
            v-if="!sections[filter.key] && activeFilters[filter.key].length > 0"
            class="filter-summary-row">
            <button
              class="summary-bin-btn"
              @click.stop="clearSpecificFilter(filter.key)">
              <svg
                width="14"
                height="14"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                stroke-width="2">
                <polyline points="3 6 5 6 21 6"></polyline>
                <path
                  d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
              </svg>
            </button>
            <div class="summary-text-wrapper">
              <span class="summary-text">{{
                formatSelectionText(filter.key)
              }}</span>
            </div>
            <span v-if="getRemainder(filter.key) > 0" class="summary-badge"
              >+{{ getRemainder(filter.key) }}</span
            >
          </div>

          <!-- Body -->
          <div v-show="sections[filter.key]" class="filter-body">
            <div v-for="opt in filter.options" :key="opt" class="checkbox-row">
              <input
                :id="`${filter.key}-${opt}`"
                v-model="activeFilters[filter.key]"
                type="checkbox"
                :value="opt"
                @change="triggerSearch" />
              <label :for="`${filter.key}-${opt}`">{{ opt }}</label>
            </div>
          </div>
        </div>

        <!-- 2. ADVANCED TOGGLE -->
        <div class="advanced-trigger" @click="showAdvanced = !showAdvanced">
          {{ showAdvanced ? 'Hide' : 'Show' }} Advanced Refiners
          <span class="plus">{{ showAdvanced ? '-' : '+' }}</span>
        </div>

        <!-- 3. ADVANCED FILTERS LOOP -->
        <transition name="expand">
          <div v-if="showAdvanced" class="advanced-section">
            <div
              v-for="filter in advancedFilters"
              :key="filter.key"
              class="filter-card">
              <div
                class="filter-card-header"
                @click="toggleSection(filter.key)">
                <span>{{ filter.title }}</span>
                <span class="chevron" :class="{ open: sections[filter.key] }"
                  >›</span
                >
              </div>

              <!-- Summary Row -->
              <div
                v-if="
                  !sections[filter.key] && activeFilters[filter.key].length > 0
                "
                class="filter-summary-row">
                <button
                  class="summary-bin-btn"
                  @click.stop="clearSpecificFilter(filter.key)">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    stroke-width="2">
                    <polyline points="3 6 5 6 21 6"></polyline>
                    <path
                      d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                  </svg>
                </button>
                <div class="summary-text-wrapper">
                  <span class="summary-text">{{
                    formatSelectionText(filter.key)
                  }}</span>
                </div>
                <span v-if="getRemainder(filter.key) > 0" class="summary-badge"
                  >+{{ getRemainder(filter.key) }}</span
                >
              </div>

              <!-- Body -->
              <div v-show="sections[filter.key]" class="filter-body">
                <div
                  v-for="opt in filter.options"
                  :key="opt"
                  class="checkbox-row">
                  <input
                    :id="`${filter.key}-${opt}`"
                    v-model="activeFilters[filter.key]"
                    type="checkbox"
                    :value="opt"
                    @change="triggerSearch" />
                  <label :for="`${filter.key}-${opt}`">{{ opt }}</label>
                </div>
              </div>
            </div>
          </div>
        </transition>
      </aside>
    </div>
  </div>
</template>

<script setup lang="ts">
  import { ref, computed, onMounted, reactive } from 'vue'
  // NEW PLUGIN IMPORT
  import { useSharePoint } from '@greener-games/vue3-sharepoint-plugin'

  // --- 1. SETUP PLUGIN ---
  const { createSearch } = useSharePoint()
  const searchCtx = createSearch() // Create independent search instance

  // --- STATE ---
  const searchQuery = ref('')
  const executedQuery = ref('')
  const showAdvanced = ref(false)

  // Reactive Filters
  const activeFilters = reactive<Record<string, string[]>>({
    region: [],
    division: [],
    practice: [],
    markets: [],
    capabilities: [],
  })

  // Accordion State
  const sections = reactive<Record<string, boolean>>({
    region: true,
    division: true,
    practice: true,
    markets: true,
    capabilities: true,
  })

  // --- COMPUTED PROPERTIES ---

  // Is Loading?
  const isSearching = computed(() => searchCtx.loading.value)

  interface SearchDoc {
    id: string | number
    url?: string
    title?: string
    content?: string
    date?: string
    region?: string
    division?: string
    practice?: string
    capabilities?: string
    markets?: string
  }

  // Results List (Mapped from plugin state)
  const resultsList = computed(() => (searchCtx.results.value?.items as unknown as SearchDoc[]) || [])

  // Total Hits (From server)
  const totalHits = computed(() => searchCtx.totalHits.value)

  // Pagination Computeds
  const currentPage = computed(() => searchCtx.pageInfo.value.current)
  const totalPages = computed(() => searchCtx.pageInfo.value.total)
  const hasNextPage = computed(() => searchCtx.hasNext.value)
  const hasPrevPage = computed(() => searchCtx.hasPrev.value)

  // Logic: Show snippet only if we have executed a non-empty search
  const hasSearchTerm = computed(() => {
    return executedQuery.value && executedQuery.value.trim().length > 0
  })

  const hasActiveFilters = computed(() => {
    return Object.values(activeFilters).some((arr) => arr.length > 0)
  })

  const DIVISIONS = [
    'Infrastructure',
    'Transportation',
    'Energy',
    'Water',
    'Defence',
  ]
  const PRACTICES = [
    'Project Management',
    'Engineering',
    'Consulting',
    'Digital',
    'Architecture',
  ]
  const MARKETS = ['UK&I', 'North America', 'Middle East', 'Asia Pacific']
  const CAPABILITIES = [
    'Rail Systems',
    'Net Zero',
    'Digital Twin',
    'Cyber Security',
  ]
  const REGIONS = ['UK&I', 'North America', 'Middle East', 'Europe', 'Global']
  // --- CONFIG ---
  const filterConfigs = [
    { title: 'Region', key: 'region', options: REGIONS, isAdvanced: false },
    {
      title: 'Division',
      key: 'division',
      options: DIVISIONS,
      isAdvanced: false,
    },
    {
      title: 'Practice',
      key: 'practice',
      options: PRACTICES,
      isAdvanced: true,
    },
    { title: 'Markets', key: 'markets', options: MARKETS, isAdvanced: true },
    {
      title: 'Capabilities',
      key: 'capabilities',
      options: CAPABILITIES,
      isAdvanced: true,
    },
  ]

  const standardFilters = computed(() =>
    filterConfigs.filter((f) => !f.isAdvanced)
  )
  const advancedFilters = computed(() =>
    filterConfigs.filter((f) => f.isAdvanced)
  )

  // --- HELPERS ---
  const formatSelectionText = (key: string) => {
    const selected = activeFilters[key]
    if (!selected || selected.length === 0) return ''
    return selected.slice(0, 2).join(', ')
  }

  const getRemainder = (key: string) => {
    return Math.max(0, activeFilters[key].length - 2)
  }

  const formatSummary = (content: string) => {
    if (!content) return ''
    // New plugin (Mock or Real) returns <c0> tags for hits.
    return content
      .replace(/<c0>/g, '<mark class="highlight">')
      .replace(/<\/c0>/g, '</mark>')
  }

  // --- ACTIONS ---
  const toggleSection = (key: string) => {
    sections[key] = !sections[key]
  }

  const triggerSearch = async () => {
    // 1. Commit the input to the executed query
    executedQuery.value = searchQuery.value

    // 2. Prepare Filters for API
    const filtersToSend: Record<string, string[]> = {}
    for (const [key, val] of Object.entries(activeFilters)) {
      if (val.length > 0) filtersToSend[key] = [...val]
    }

    await searchCtx.execute({
      query: executedQuery.value,
      scope: ['/sites/TestSite/Shared Documents'],
      rowLimit: 10,
      startRow: 0, // Reset to Page 1
      mapping: {
        Title: 'title',
        Path: 'url',
        HitHighlightedSummary: 'content', // Maps to doc.content
        PublishedDate: 'date',
        // Metadata
        RefinableString01: 'region',
        RefinableString02: 'division',
        Practice: 'practice',
        Capabilities: 'capabilities',
        Markets: 'markets',
      },
    })
  }

  const clearSearchText = () => {
    searchQuery.value = ''
    void triggerSearch()
  }

  const clearAllFilters = () => {
    Object.keys(activeFilters).forEach((k) => (activeFilters[k] = []))
    void triggerSearch()
  }

  const clearSpecificFilter = (key: string) => {
    activeFilters[key] = []
    void triggerSearch()
  }

  // Updated Pagination Action
  const changePage = (p: number) => {
    searchCtx.goToPage(p)
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  // Initial Load
  onMounted(() => { void triggerSearch() })
</script>

<style scoped>
  /* BASE */
  :root {
    --primary-blue: #326cf6;
    --bg-sidebar: #eaf5fa;
  }

  .page-container {
    max-width: 1400px;
    margin: 0 auto;
    font-family: 'Segoe UI', sans-serif;
    color: #333;
    position: relative;
  }

  /* STICKY HEADER */
  .search-hero {
    position: sticky;
    top: 0;
    z-index: 100;
    background: white;
    padding: 20px 0;
    border-bottom: 1px solid #f0f0f0;
  }

  .search-input-group {
    display: flex;
    gap: 10px;
    width: 100%;
  }

  .input-wrapper {
    position: relative;
    flex: 1;
  }

  .hero-input {
    width: 100%;
    padding: 16px 40px 16px 50px;
    border: 1px solid #ccc;
    border-radius: 4px;
    font-size: 18px;
    color: #326cf6;
    box-shadow: 0 2px 4px rgb(0 0 0 / 5%);
    transition: border-color 0.2s;
  }

  .hero-input:focus {
    border-color: #326cf6;
    outline: none;
  }

  .search-icon {
    position: absolute;
    left: 15px;
    top: 50%;
    transform: translateY(-50%);
    width: 24px;
    color: #326cf6;
  }

  /* Clear Text Button */
  .clear-text-btn {
    position: absolute;
    right: 12px;
    top: 50%;
    transform: translateY(-50%);
    background: none;
    border: none;
    color: #999;
    font-size: 14px;
    cursor: pointer;
    padding: 4px;
    border-radius: 50%;
  }

  .clear-text-btn:hover {
    background-color: #eee;
    color: #333;
  }

  /* Search Button */
  .search-btn {
    padding: 0 32px;
    background-color: #326cf6;
    color: white;
    border: none;
    border-radius: 4px;
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    transition: background-color 0.2s;
    white-space: nowrap;
  }

  .search-btn:hover {
    background-color: #2554c7;
  }

  .search-btn:disabled {
    background-color: #a0bbf5;
    cursor: default;
  }

  /* LAYOUT */
  .main-layout {
    display: grid;
    grid-template-columns: 1fr 280px;
    gap: 40px;
    align-items: start;
    margin-top: 20px;
  }

  .results-column {
    min-height: 600px;
  }

  /* RESULTS */
  .results-header {
    display: flex;
    justify-content: space-between;
    margin-bottom: 20px;
    border-bottom: 1px solid #eee;
    padding-bottom: 10px;
  }

  .result-item {
    display: flex;
    gap: 20px;
    padding: 24px 0;
    border-bottom: 1px solid #e0e0e0;
  }

  .doc-icon-wrapper {
    color: #99abb4;
    width: 32px;
    padding-top: 4px;
  }

  .doc-title {
    font-size: 18px;
    font-weight: 600;
    color: #3f51b5;
    text-decoration: none;
    margin-bottom: 8px;
    display: block;
  }

  /* SEARCH SNIPPET STYLES */
  .search-snippet {
    background-color: #fdfdfd;
    border-left: 3px solid #326cf6;
    padding: 8px 12px;
    margin-bottom: 12px;
    border-radius: 0 4px 4px 0;
  }

  .snippet-label {
    font-size: 11px;
    text-transform: uppercase;
    color: #888;
    margin-bottom: 4px;
    font-weight: 600;
  }

  .search-snippet p {
    margin: 0;
    font-size: 14px;
    color: #444;
    font-style: italic;
    line-height: 1.5;
  }

  :deep(.highlight) {
    background-color: #fff4ce;
    color: #000;
    font-weight: bold;
    padding: 0 2px;
    border-radius: 2px;
  }

  /* METADATA */
  .meta-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 10px 30px;
    margin-top: 8px;
    font-size: 14px;
  }

  .meta-label {
    font-weight: 700;
    color: #000;
    margin-right: 4px;
  }

  .full-width {
    width: 100%;
  }

  /* SIDEBAR */
  .refiners-sidebar {
    background-color: #eaf5fa;
    border-radius: 4px;
    padding: 15px;
    position: sticky;
    top: 110px;
    height: fit-content;
    max-height: 80vh;
    overflow-y: auto;
  }

  .sidebar-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 10px;
  }

  .header-title {
    display: flex;
    align-items: center;
    gap: 8px;
    font-weight: 600;
  }

  .clear-btn {
    background: none;
    border: none;
    color: #3f51b5;
    text-decoration: underline;
    cursor: pointer;
    font-size: 12px;
  }

  /* FILTER CARDS */
  .filter-card {
    background: white;
    border: 1px solid #dbeaf0;
    border-radius: 8px;
    margin-bottom: 12px;
    overflow: hidden;
  }

  .filter-card-header {
    padding: 12px;
    display: flex;
    justify-content: space-between;
    font-weight: 600;
    font-size: 14px;
    cursor: pointer;
    user-select: none;
    transition: background 0.2s;
  }

  .filter-card-header:hover {
    background: #f9f9f9;
  }

  .chevron {
    transition: transform 0.2s;
    font-size: 18px;
    color: #326cf6;
  }

  .chevron.open {
    transform: rotate(90deg);
  }

  .filter-body {
    padding: 0 12px 12px;
    border-top: 1px solid #f0f0f0;
  }

  /* SUMMARY ROW */
  .filter-summary-row {
    display: flex;
    align-items: center;
    gap: 8px;
    background-color: #f4f7f9;
    padding: 8px 12px;
    margin: 0 12px 12px;
    border-radius: 4px;
  }

  .summary-bin-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    background: none;
    border: 1px solid transparent;
    color: #326cf6;
    cursor: pointer;
    padding: 2px;
    border-radius: 4px;
  }

  .summary-bin-btn:hover {
    background: #e1e8ed;
  }

  .summary-text-wrapper {
    flex: 1;
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
  }

  .summary-text {
    font-size: 13px;
    color: #333;
  }

  .summary-badge {
    font-size: 12px;
    font-weight: 700;
    color: #333;
  }

  /* CHECKBOXES */
  .checkbox-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 0;
  }

  .checkbox-row label {
    font-size: 14px;
    color: #333;
    cursor: pointer;
    flex: 1;
  }

  .checkbox-row input {
    cursor: pointer;
    accent-color: #326cf6;
  }

  /* ADVANCED BUTTON */
  .advanced-trigger {
    background: #d0e5f0;
    padding: 12px;
    border-radius: 4px;
    display: flex;
    align-items: center;
    gap: 10px;
    font-weight: 600;
    font-size: 14px;
    cursor: pointer;
    margin: 20px 0 12px;
    user-select: none;
  }

  .advanced-trigger .plus {
    margin-left: auto;
    font-size: 18px;
    font-weight: 400;
  }

  /* ANIMATION */
  .expand-enter-active,
  .expand-leave-active {
    transition: all 0.3s ease;
    max-height: 1000px;
    opacity: 1;
    overflow: hidden;
  }

  .expand-enter-from,
  .expand-leave-to {
    max-height: 0;
    opacity: 0;
    transform: translateY(-10px);
  }

  .pagination-footer {
    margin-top: 30px;
    display: flex;
    gap: 10px;
    justify-content: center;
  }

  .pagination-footer button {
    padding: 8px 16px;
    background: white;
    border: 1px solid #ccc;
    cursor: pointer;
  }

  .pagination-footer button:disabled {
    opacity: 0.5;
  }

  @media (width <= 900px) {
    .main-layout {
      grid-template-columns: 1fr;
    }

    .refiners-sidebar {
      position: static;
      margin-bottom: 40px;
    }
  }
</style>