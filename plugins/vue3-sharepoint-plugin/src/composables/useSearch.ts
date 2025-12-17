import { ref, computed, Ref } from 'vue'
import { ISharePointClient, SearchRequestOptions, SearchResult } from '../types'

export interface SearchInstance {
  loading: Ref<boolean>
  error: Ref<string | null>
  results: Ref<SearchResult<any> | null>
  execute: <T = any>(opts: SearchRequestOptions) => Promise<void>
  nextPage: () => void
  prevPage: () => void
  goToPage: (page: number) => void
  pageInfo: Ref<{ current: number; total: number; pageSize: number }>
  totalHits: Ref<number>
  hasNext: Ref<boolean>
  hasPrev: Ref<boolean>
}

export function useSearch(client: ISharePointClient): SearchInstance {
  const loading = ref(false)
  const error = ref<string | null>(null)
  const results = ref<SearchResult<any> | null>(null)

  // Store last options to handle pagination correctly
  const _lastOptions = ref<SearchRequestOptions | null>(null)

  const execute = async <T = any>(opts: SearchRequestOptions) => {
    loading.value = true
    error.value = null
    // Clone options to freeze state for pagination
    const safeOptions = JSON.parse(JSON.stringify(opts))

    try {
      const data = await client.search<T>(safeOptions)
      results.value = data
      _lastOptions.value = safeOptions
    } catch (e: any) {
      error.value = e.message
      results.value = null
    } finally {
      loading.value = false
    }
  }

  const totalHits = computed(() => results.value?.totalHits || 0)

  const pageInfo = computed(() => {
    if (!results.value || !_lastOptions.value)
      return { current: 1, total: 0, pageSize: 10 }
    const limit = _lastOptions.value.rowLimit || 10
    const current = Math.floor(results.value.startRow / limit) + 1
    const total = Math.ceil(results.value.totalHits / limit)
    return { current, total, pageSize: limit }
  })

  const hasNext = computed(() => {
    if (!results.value || !_lastOptions.value) return false
    return (
      results.value.startRow + (_lastOptions.value.rowLimit || 10) <
      results.value.totalHits
    )
  })

  const hasPrev = computed(() => !!results.value && results.value.startRow > 0)

  const nextPage = () => {
    if (!hasNext.value || !_lastOptions.value) return
    const nextRow =
      (results.value?.startRow || 0) + (_lastOptions.value.rowLimit || 10)
    execute({ ..._lastOptions.value, startRow: nextRow })
  }

  const prevPage = () => {
    if (!hasPrev.value || !_lastOptions.value) return
    const pageSize = _lastOptions.value.rowLimit || 10
    let prevRow = (results.value?.startRow || 0) - pageSize
    if (prevRow < 0) prevRow = 0
    execute({ ..._lastOptions.value, startRow: prevRow })
  }

  const goToPage = (pageNumber: number) => {
    if (!_lastOptions.value) return
    const pageSize = _lastOptions.value.rowLimit || 10
    const newStart = (pageNumber - 1) * pageSize
    execute({ ..._lastOptions.value, startRow: newStart })
  }

  return {
    loading,
    error,
    results,
    execute,
    nextPage,
    prevPage,
    goToPage,
    pageInfo,
    totalHits,
    hasNext,
    hasPrev,
  }
}
