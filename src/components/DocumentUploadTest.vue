<template>
  <div class="drop-off-library">
    <!-- HEADER -->
    <div class="header-section">
      <h1>Drop Off Library</h1>
      <p class="subtitle">(b) Ultricies tortor, lectus eget libero sit sit donec purus. Molestie lectus etiam erat sit</p>
    </div>

    <!-- DRAG AND DROP ZONE -->
    <div
        class="upload-zone"
        :class="{ 'is-dragging': isDragging }"
        @dragover.prevent="isDragging = true"
        @dragleave.prevent="isDragging = false"
        @drop.prevent="handleDrop"
    >
      <div class="upload-content">
        <div class="upload-icon-wrapper">
          <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="17 8 12 3 7 8"></polyline>
            <line x1="12" y1="3" x2="12" y2="15"></line>
          </svg>
        </div>

        <div class="upload-text">
          <button class="link-btn" @click="triggerFileInput">Click to Upload</button>
          <span class="text-muted"> or Drag & Drop files</span>
        </div>
        <p class="file-limit">(Max File Size: 25mb)</p>

        <input
            ref="fileInputRef"
            type="file"
            multiple
            class="hidden-input"
            @change="handleFileSelect"
        />
      </div>
    </div>

    <!-- INSTRUCTION TEXT -->
    <div v-if="files.length > 0" class="instruction-text">
      Enter missing field data to submit files.
    </div>

    <!-- DATA TABLE -->
    <div v-if="files.length > 0" class="table-container">
      <table>
        <thead>
        <tr>
          <th style="width: 20%">File</th>
          <th style="width: 15%">Type</th>
          <th style="width: 15%">Language</th>
          <th style="width: 15%">Owner</th>
          <th style="width: 10%">Created</th>
          <th style="width: 15%">Status</th>
          <th style="width: 5%"></th> <!-- Size -->
          <th style="width: 5%"></th> <!-- Delete -->
        </tr>
        </thead>
        <tbody>
        <tr v-for="(item, index) in files" :key="index">
          <!-- File Name -->
          <td>
            <div class="file-name-cell">
              <span class="paperclip-icon">🔗</span>
              <span class="filename-text" :title="item.file.name">{{ truncateName(item.file.name) }}</span>
            </div>
          </td>

          <!-- Type Dropdown -->
          <td>
            <select v-model="item.meta.type" class="table-input">
              <option value="" disabled selected>Select Type</option>
              <option value="CV">Custom CV</option>
              <option value="Headshot">Headshot</option>
              <option value="Certificate">Certificate</option>
            </select>
          </td>

          <!-- Language Dropdown -->
          <td>
            <select v-model="item.meta.language" class="table-input">
              <option value="" disabled selected>Select Language</option>
              <option value="English">English</option>
              <option value="French">French</option>
              <option value="Spanish">Spanish</option>
              <option value="N/A">N/A</option>
            </select>
          </td>

          <!-- Owner Input -->
          <td>
            <input
                v-model="item.meta.owner"
                type="text"
                placeholder="Enter name..."
                class="table-input"
            >
          </td>

          <!-- Created Date -->
          <td class="text-muted">{{ item.created }}</td>

          <!-- Status -->
          <td>
              <span :class="['status-text', getStatus(item).class]">
                {{ getStatus(item).text }}
                <span v-if="getStatus(item).valid" class="check-icon">✓</span>
              </span>
          </td>

          <!-- Size -->
          <td class="text-muted text-right">{{ formatSize(item.file.size) }}</td>

          <!-- Delete Action -->
          <td class="text-center">
            <button class="trash-btn" @click="removeFile(index)">
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
            </button>
          </td>
        </tr>
        </tbody>
      </table>
    </div>

    <!-- FOOTER ACTIONS -->
    <div class="footer-actions">
      <button class="btn-outline" @click="triggerFileInput">Attach More File(s)</button>
      <button
          class="btn-primary"
          :disabled="!canSubmit || isUploading"
          @click="submitFiles"
      >
        {{ isUploading ? 'Uploading...' : 'Submit' }}
      </button>
    </div>

    <!-- GLOBAL MESSAGES -->
    <div v-if="globalMessage" class="message-area" :class="{ error: isError }">
      {{ globalMessage }}
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, reactive, computed } from 'vue'
import { useSharePoint } from '@greener-games/vue3-sharepoint-plugin'

// --- Types ---
interface FileMeta {
  type: string
  language: string
  owner: string
}

interface FileRow {
  file: File
  id: string
  created: string
  meta: FileMeta
  uploadStatus: 'pending' | 'success' | 'error'
}

// --- Config ---
// Hardcoded path based on your previous snippet, since the UI no longer has an input for it
const TARGET_FOLDER = '/sites/TestSite/Shared Documents'

// --- State ---
const { upload, updateFileMetadata } = useSharePoint()
const fileInputRef = ref<HTMLInputElement | null>(null)
const files = reactive<FileRow[]>([])
const isDragging = ref(false)
const isUploading = ref(false)
const globalMessage = ref('')
const isError = ref(false)

// --- Computeds ---
const canSubmit = computed(() => {
  if (files.length === 0) return false
  // Check if every file has required fields filled
  return files.every(f => f.meta.type && f.meta.language && f.meta.owner)
})

// --- Helpers ---
const triggerFileInput = () => {
  fileInputRef.value?.click()
}

const formatSize = (bytes: number) => {
  if (bytes === 0) return '0mb'
  const mb = bytes / (1024 * 1024)
  if (mb < 0.1) return '<0.1mb'
  return mb.toFixed(1) + 'mb'
}

const truncateName = (name: string) => {
  return name.length > 25 ? name.substring(0, 22) + '...' : name
}

const getStatus = (item: FileRow) => {
  if (item.uploadStatus === 'success') return { text: 'Uploaded', class: 'text-success', valid: true }
  if (item.uploadStatus === 'error') return { text: 'Failed', class: 'text-error', valid: false }

  const isValid = item.meta.type && item.meta.language && item.meta.owner
  return isValid
      ? { text: 'Ready to go', class: 'text-success', valid: true }
      : { text: 'Requires info', class: 'text-danger', valid: false }
}

// --- File Handling ---
const processFiles = (fileList: FileList | null) => {
  if (!fileList) return

  const newFiles = Array.from(fileList).map(f => ({
    file: f,
    id: Math.random().toString(36).substring(7),
    created: new Date().toLocaleDateString('en-GB').replace(/\//g, '-'), // 00-00-0000 format
    meta: {
      type: '',
      language: '',
      owner: '' // Defaults to empty, user must fill
    },
    uploadStatus: 'pending' as const
  }))

  files.push(...newFiles)
}

const handleFileSelect = (event: Event) => {
  const input = event.target as HTMLInputElement
  processFiles(input.files)
  input.value = '' // Reset
}

const handleDrop = (event: DragEvent) => {
  isDragging.value = false
  if (event.dataTransfer?.files) {
    processFiles(event.dataTransfer.files)
  }
}

const removeFile = (index: number) => {
  files.splice(index, 1)
}

// --- Upload Logic ---
const submitFiles = async () => {
  if (!canSubmit.value) return

  isUploading.value = true
  globalMessage.value = ''
  isError.value = false

  let failureCount = 0

  for (const item of files) {
    if (item.uploadStatus === 'success') continue

    try {
      // 1. Upload File
      await upload(TARGET_FOLDER, item.file, item.file.name)

      // 2. Set Metadata
      // Construct URL (SharePoint usually needs server relative URL for updating metadata)
      const cleanFolder = TARGET_FOLDER.replace(/\/+$/, '')
      const fileUrl = `${cleanFolder}/${item.file.name}`

      // Map UI fields to SharePoint Internal Column Names
      // Ensure these keys (DocumentType, DocLanguage, DocOwner) match your SP List settings!
      const metadata = {
        DocumentType: item.meta.type,
        DocLanguage: item.meta.language,
        DocOwner: item.meta.owner
      }

      await updateFileMetadata(fileUrl, metadata)

      item.uploadStatus = 'success'
    } catch (e) {
      console.error('Upload failed for', item.file.name, e)
      item.uploadStatus = 'error'
      failureCount++
    }
  }

  isUploading.value = false
  if (failureCount > 0) {
    isError.value = true
    globalMessage.value = `${failureCount} file(s) failed to upload.`
  } else {
    globalMessage.value = 'All files submitted successfully.'
    // Optional: Clear list on success
    // files.splice(0, files.length)
  }
}
</script>

<style scoped>
/* RESET & BASICS */
.drop-off-library {
  max-width: 1100px;
  margin: 0 auto;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  color: #333;
  padding: 20px;
  background-color: #fff;
}

/* HEADER */
.header-section {
  margin-bottom: 30px;
  border-bottom: 1px dotted #ccc; /* Mimics top line in image */
  padding-bottom: 10px;
}

.header-section h1 {
  font-size: 28px;
  margin: 0 0 5px;
  font-weight: 400;
  color: #000;
}

.subtitle {
  font-size: 13px;
  color: #666;
  margin: 0;
}

/* UPLOAD ZONE */
.upload-zone {
  background-color: #f2f6fc; /* Very light blue/grey */
  border: 2px dashed #dbeafe; /* Light blue dashed border */
  border-radius: 12px;
  padding: 60px 20px;
  text-align: center;
  transition: background-color 0.2s;
  margin-bottom: 20px;
}

.upload-zone.is-dragging {
  background-color: #e0eaff;
  border-color: #3b82f6;
}

.upload-icon-wrapper {
  background-color: #5b4bf5; /* Purple icon bg */
  width: 48px;
  height: 48px;
  border-radius: 8px;
  display: flex;
  align-items: center;
  justify-content: center;
  margin: 0 auto 15px;
  box-shadow: 0 4px 6px rgb(91 75 245 / 20%);
}

.upload-text {
  font-size: 16px;
  margin-bottom: 5px;
}

.link-btn {
  background: none;
  border: none;
  color: #3b49df; /* Blue link color */
  font-weight: 600;
  cursor: pointer;
  padding: 0;
  font-size: inherit;
}

.link-btn:hover {
  text-decoration: underline;
}

.text-muted {
  color: #666;
}

.file-limit {
  font-size: 12px;
  color: #888;
  margin: 0;
}

.hidden-input {
  display: none;
}

/* INSTRUCTION */
.instruction-text {
  font-style: italic;
  font-size: 13px;
  color: #555;
  margin-bottom: 15px;
  text-align: center;
}

/* TABLE */
.table-container {
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  overflow: hidden; /* For rounded corners */
  margin-bottom: 30px;
}

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}

thead {
  background-color: #f8f9fa;
}

th {
  text-align: left;
  padding: 12px 15px;
  color: #4f46e5; /* Purple/Blue header text */
  font-weight: 600;
  border-bottom: 1px solid #e5e7eb;
}

td {
  padding: 12px 15px;
  border-bottom: 1px solid #e5e7eb;
  vertical-align: middle;
}

tr:last-child td {
  border-bottom: none;
}

/* TABLE INPUTS */
.table-input {
  width: 100%;
  padding: 6px 8px;
  border: 1px solid #e5e7eb; /* Subtle border */
  border-radius: 4px;
  font-size: 13px;
  color: #333;
  outline: none;
  background: #fff;
}

.table-input:focus {
  border-color: #5b4bf5;
}

.table-input::placeholder {
  color: #999;
}

/* CELL STYLES */
.file-name-cell {
  display: flex;
  align-items: center;
  gap: 8px;
}

.paperclip-icon {
  color: #5b4bf5;
  font-size: 16px;
}

.filename-text {
  color: #5b4bf5;
  text-decoration: underline;
  cursor: pointer;
  white-space: nowrap;
}

.status-text {
  font-weight: 600;
  font-size: 12px;
  display: inline-flex;
  align-items: center;
  gap: 4px;
}

.text-danger { color: #ef4444; }
.text-success { color: #10b981; }
.text-error { color: #dc2626; }

.check-icon {
  border: 1px solid #10b981;
  border-radius: 50%;
  width: 14px;
  height: 14px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 10px;
  color: #10b981;
}

.trash-btn {
  background: none;
  border: none;
  cursor: pointer;
  color: #666;
  opacity: 0.7;
  transition: opacity 0.2s;
}

.trash-btn:hover {
  opacity: 1;
  color: #ef4444;
}

.text-right { text-align: right; }
.text-center { text-align: center; }

/* FOOTER */
.footer-actions {
  display: flex;
  justify-content: center;
  gap: 15px;
  margin-top: 20px;
}

.btn-outline {
  background: white;
  border: 1px solid #5b4bf5;
  color: #5b4bf5;
  padding: 10px 24px;
  border-radius: 6px;
  font-weight: 600;
  cursor: pointer;
  font-size: 14px;
  transition: background 0.2s;
}

.btn-outline:hover {
  background: #f0f0ff;
}

.btn-primary {
  background: #c7d2fe; /* Disabled look initially (light purple) */
  color: white;
  border: none;
  padding: 10px 40px;
  border-radius: 6px;
  font-weight: 600;
  cursor: not-allowed;
  font-size: 14px;
}

.btn-primary:not(:disabled) {
  background: #5b4bf5; /* Active purple */
  cursor: pointer;
  box-shadow: 0 2px 4px rgb(91 75 245 / 30%);
}

.message-area {
  margin-top: 20px;
  text-align: center;
  padding: 10px;
  border-radius: 4px;
  background-color: #ecfdf5;
  color: #065f46;
}

.message-area.error {
  background-color: #fef2f2;
  color: #991b1b;
}
</style>