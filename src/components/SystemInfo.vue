<template>
  <div class="system-info-container">
    <div class="header-section">
      <h1>System Information</h1>
      <p class="subtitle">Overview of the current SharePoint context, including current user, web, and lists.</p>
    </div>

    <div v-if="loading" class="loading-state">
      <div class="spinner"></div>
      Loading system info...
    </div>

    <div v-else class="grid-layout">
      <!-- Current User Card -->
      <div class="info-card user-card">
        <div class="card-header">
          <h3>Current User</h3>
        </div>
        <div class="card-body">
          <div class="user-avatar">
            {{ userInitials }}
          </div>
          <div class="user-details">
            <div class="detail-row">
              <span class="label">Name:</span>
              <span class="value">{{ currentUser?.Title }}</span>
            </div>
            <div class="detail-row">
              <span class="label">Email:</span>
              <span class="value">{{ currentUser?.Email }}</span>
            </div>
            <div class="detail-row">
              <span class="label">Login:</span>
              <span class="value">{{ currentUser?.LoginName }}</span>
            </div>
          </div>
        </div>
      </div>

      <!-- Web Info Card -->
      <div class="info-card web-card">
        <div class="card-header">
          <h3>Web Information</h3>
        </div>
        <div class="card-body">
          <div class="detail-row">
            <span class="label">Title:</span>
            <span class="value">{{ webInfo?.Title }}</span>
          </div>
          <div class="detail-row">
            <span class="label">URL:</span>
            <span class="value link" @click="openLink(webInfo?.Url)">{{ webInfo?.Url }}</span>
          </div>
          <div class="detail-row">
            <span class="label">ID:</span>
            <span class="value mono">{{ webInfo?.Id }}</span>
          </div>
          <div class="detail-row">
            <span class="label">Description:</span>
            <span class="value">{{ webInfo?.Description || 'N/A' }}</span>
          </div>
        </div>
      </div>

      <!-- Lists Card -->
      <div class="info-card lists-card full-width">
        <div class="card-header">
          <h3>Lists & Libraries ({{ lists.length }})</h3>
          <button @click="refreshLists" class="refresh-btn" title="Refresh Lists">↻</button>
        </div>
        <div class="card-body no-padding">
          <div class="table-responsive">
            <table>
              <thead>
                <tr>
                  <th>Title</th>
                  <th>Item Count</th>
                  <th>Hidden</th>
                  <th>ID</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="list in lists" :key="list.Id">
                  <td class="font-medium">{{ list.Title }}</td>
                  <td>{{ list.ItemCount }}</td>
                  <td>
                    <span :class="['badge', list.Hidden ? 'badge-warn' : 'badge-success']">
                      {{ list.Hidden ? 'Hidden' : 'Visible' }}
                    </span>
                  </td>
                  <td class="mono small">{{ list.Id }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed } from 'vue'
import { useSharePoint, type UserInfo, type WebInfo, type ListInfo } from 'vue3-sharepoint-plugin'

const { loadUser, getWebInfo, getLists } = useSharePoint()

const currentUser = ref<UserInfo | null>(null)
const webInfo = ref<WebInfo | null>(null)
const lists = ref<ListInfo[]>([])
const loading = ref(true)

const userInitials = computed(() => {
  if (!currentUser.value?.Title) return '?'
  return currentUser.value.Title
    .split(' ')
    .map(n => n[0])
    .join('')
    .substring(0, 2)
    .toUpperCase()
})

const loadData = async () => {
  loading.value = true
  try {
    const [user, web, allLists] = await Promise.all([
      loadUser(),
      getWebInfo(),
      getLists()
    ])
    currentUser.value = user
    webInfo.value = web
    lists.value = allLists
  } catch (e) {
    console.error('Failed to load system info', e)
  } finally {
    loading.value = false
  }
}

const refreshLists = async () => {
  const allLists = await getLists()
  lists.value = allLists
}

const openLink = (url?: string) => {
  if (url) window.open(url, '_blank')
}

onMounted(() => {
  loadData()
})
</script>

<style scoped>
.system-info-container {
  max-width: 1200px;
  margin: 0 auto;
  font-family: 'Segoe UI', sans-serif;
  color: #333;
  padding: 20px;
}

.header-section {
  margin-bottom: 30px;
  border-bottom: 1px solid #e5e7eb;
  padding-bottom: 15px;
}

.header-section h1 {
  font-size: 28px;
  margin: 0 0 8px 0;
  font-weight: 400;
  color: #111827;
}

.subtitle {
  font-size: 14px;
  color: #6b7280;
  margin: 0;
}

.loading-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 60px;
  color: #6b7280;
  gap: 15px;
}

.spinner {
  width: 30px;
  height: 30px;
  border: 3px solid #e5e7eb;
  border-top-color: #2563eb;
  border-radius: 50%;
  animation: spin 1s linear infinite;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

.grid-layout {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 24px;
}

.full-width {
  grid-column: span 2;
}

.info-card {
  background: white;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
}

.card-header {
  padding: 15px 20px;
  background-color: #f9fafb;
  border-bottom: 1px solid #e5e7eb;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.card-header h3 {
  margin: 0;
  font-size: 16px;
  font-weight: 600;
  color: #374151;
}

.card-body {
  padding: 20px;
}

.card-body.no-padding {
  padding: 0;
}

/* User Card */
.user-card .card-body {
  display: flex;
  align-items: center;
  gap: 20px;
}

.user-avatar {
  width: 60px;
  height: 60px;
  background-color: #bfdbfe;
  color: #1e40af;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 24px;
  font-weight: 700;
}

.user-details {
  flex: 1;
}

/* Details */
.detail-row {
  display: flex;
  margin-bottom: 8px;
  font-size: 14px;
}

.detail-row:last-child {
  margin-bottom: 0;
}

.label {
  width: 100px;
  font-weight: 600;
  color: #4b5563;
  flex-shrink: 0;
}

.value {
  color: #111827;
  word-break: break-all;
}

.value.link {
  color: #2563eb;
  cursor: pointer;
  text-decoration: underline;
}

.value.mono {
  font-family: monospace;
  background: #f3f4f6;
  padding: 2px 6px;
  border-radius: 4px;
  font-size: 12px;
}

/* Table */
.table-responsive {
  overflow-x: auto;
}

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}

th {
  text-align: left;
  padding: 12px 20px;
  background-color: #f9fafb;
  border-bottom: 1px solid #e5e7eb;
  font-weight: 600;
  color: #4b5563;
  font-size: 13px;
  text-transform: uppercase;
  letter-spacing: 0.05em;
}

td {
  padding: 12px 20px;
  border-bottom: 1px solid #f3f4f6;
  color: #374151;
}

tr:last-child td {
  border-bottom: none;
}

.font-medium {
  font-weight: 500;
  color: #111827;
}

.small {
  font-size: 12px;
}

.refresh-btn {
  background: none;
  border: none;
  font-size: 18px;
  cursor: pointer;
  color: #6b7280;
  transition: color 0.2s;
}

.refresh-btn:hover {
  color: #2563eb;
}

.badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 9999px;
  font-size: 12px;
  font-weight: 500;
}

.badge-success {
  background-color: #d1fae5;
  color: #065f46;
}

.badge-warn {
  background-color: #fef3c7;
  color: #92400e;
}

@media (max-width: 768px) {
  .grid-layout {
    grid-template-columns: 1fr;
  }
  .full-width {
    grid-column: span 1;
  }
}
</style>
