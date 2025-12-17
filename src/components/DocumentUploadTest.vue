<template>
  <div class="page-container">
    <!-- HERO / HEADER -->
    <div class="search-hero">
      <div class="search-content">
        <h1>Document Upload Test</h1>
        <p>Upload one or multiple documents to a specific SharePoint folder with metadata.</p>
      </div>
    </div>

    <!-- MAIN CONTENT -->
    <div class="main-layout">
      <div class="upload-container">

        <!-- CONFIGURATION FORM -->
        <div class="form-card">
          <div class="form-header">
            <h2>Upload Configuration</h2>
          </div>
          <div class="form-body">
            <!-- Target Path -->
            <div class="input-group">
              <label>Target Folder Path</label>
              <input
                v-model="folderPath"
                type="text"
                placeholder="/sites/MySite/Shared Documents"
                class="text-input"
              />
              <span class="hint">Relative path to the document library/folder</span>
            </div>

            <!-- Metadata Fields -->
            <div class="meta-row">
              <div class="input-group">
                <label>Field A</label>
                <input
                  v-model="fieldA"
                  type="text"
                  placeholder="Enter value for Field A"
                  class="text-input"
                />
              </div>
              <div class="input-group">
                <label>Field B</label>
                <input
                  v-model="fieldB"
                  type="text"
                  placeholder="Enter value for Field B"
                  class="text-input"
                />
              </div>
            </div>
          </div>
        </div>

        <!-- FILE SELECTION -->
        <div class="form-card">
          <div class="form-header">
            <h2>Documents</h2>
          </div>
          <div class="form-body">
            <div class="file-drop-area">
              <input
                type="file"
                multiple
                id="fileInput"
                @change="handleFileSelect"
                class="file-input-hidden"
              />
              <label for="fileInput" class="file-drop-label">
                <span class="icon">📂</span>
                <span>Click to select files</span>
              </label>
            </div>

            <!-- Selected Files List -->
            <div v-if="files.length > 0" class="file-list">
              <div v-for="(fileWrapper, index) in files" :key="index" class="file-item">
                <div class="file-info">
                  <span class="file-name">{{ fileWrapper.file.name }}</span>
                  <span class="file-size">{{ formatSize(fileWrapper.file.size) }}</span>
                </div>
                <div class="file-status">
                  <span :class="['status-badge', fileWrapper.status]">
                    {{ fileWrapper.status }}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- ACTIONS -->
        <div class="actions-row">
          <button
            @click="triggerUpload"
            :disabled="isUploading || files.length === 0"
            class="primary-btn"
          >
            {{ isUploading ? 'Uploading...' : 'Start Upload' }}
          </button>

          <div v-if="globalMessage" class="global-message" :class="{ error: isError }">
            {{ globalMessage }}
          </div>
        </div>

      </div>
    </div>
  </div>
</template>


<script setup lang="ts">
  import { ref, reactive } from 'vue'
  import { useSharePoint } from 'vue3-sharepoint-plugin'

  // --- Types ---
  type UploadStatus = 'pending' | 'uploading' | 'success' | 'error'

  interface FileWrapper {
    file: File
    status: UploadStatus
  }

  // --- SharePoint Integration ---
  const { upload, updateFileMetadata } = useSharePoint()

  // --- State ---
  const folderPath = ref('/sites/ProposalExpertiseHubUAT/V2/Shared Documents') // Default example
  const fieldA = ref('')
  const fieldB = ref('')
  const files = reactive<FileWrapper[]>([])
  const isUploading = ref(false)
  const globalMessage = ref('')
  const isError = ref(false)

  // --- Actions ---

  const handleFileSelect = (event: Event) => {
    const input = event.target as HTMLInputElement
    if (input.files) {
      const newFiles = Array.from(input.files).map(f => ({
        file: f,
        status: 'pending' as UploadStatus
      }))
      files.push(...newFiles)
    }
    // Reset input so same files can be selected again if needed
    input.value = ''
  }

  const formatSize = (bytes: number) => {
    if (bytes === 0) return '0 B'
    const k = 1024
    const sizes = ['B', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
  }

  // --- Upload Logic ---
  const triggerUpload = async () => {
    if (files.length === 0) return

    isUploading.value = true
    globalMessage.value = ''
    isError.value = false

    try {
      // Process files sequentially or parallel
      for (const fileWrapper of files) {
        if (fileWrapper.status === 'success') continue // Skip already uploaded

        fileWrapper.status = 'uploading'
        try {
          await uploadSingleFile(fileWrapper.file)
          fileWrapper.status = 'success'
        } catch (e) {
          console.error(e)
          fileWrapper.status = 'error'
          isError.value = true
        }
      }
      globalMessage.value = isError.value ? 'Some uploads failed.' : 'All files uploaded successfully!'
    } catch (e) {
      globalMessage.value = 'An unexpected error occurred.'
      isError.value = true
    } finally {
      isUploading.value = false
    }
  }

  // Real upload function using useSharePoint
  const uploadSingleFile = async (file: File) => {
    // 1. Upload
    // Assuming signature: upload(folderPath, fileContent, fileName)
    // We use the file name as provided by the user selection
    await upload(folderPath.value, file, file.name)

    // 2. Metadata
    // Only update if fields are provided
    if (fieldA.value || fieldB.value) {
      const cleanFolder = folderPath.value.replace(/\/+$/, '')
      const fileUrl = `${cleanFolder}/${file.name}`

      // Construct metadata object
      // Note: Keys "FieldA" and "FieldB" should match internal SharePoint column names
      const metadata: Record<string, string> = {}
      if (fieldA.value) metadata.FieldA = fieldA.value
      if (fieldB.value) metadata.FieldB = fieldB.value

      await updateFileMetadata(fileUrl, metadata)
    }
  }
</script>

<style scoped>
  /* Reuse styles consistent with CvSearch */
  .page-container {
    max-width: 1000px;
    margin: 0 auto;
    font-family: 'Segoe UI', sans-serif;
    color: #333;
  }

  .search-hero {
    background: white;
    padding: 30px 0;
    border-bottom: 1px solid #f0f0f0;
    margin-bottom: 30px;
  }

  .search-content h1 {
    font-size: 28px;
    color: #326cf6;
    margin-bottom: 10px;
  }

  .upload-container {
    display: flex;
    flex-direction: column;
    gap: 20px;
  }

  .form-card {
    background: white;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    overflow: hidden;
  }

  .form-header {
    background: #f9f9f9;
    padding: 15px 20px;
    border-bottom: 1px solid #e0e0e0;
  }

  .form-header h2 {
    margin: 0;
    font-size: 18px;
    color: #333;
  }

  .form-body {
    padding: 20px;
  }

  .input-group {
    display: flex;
    flex-direction: column;
    margin-bottom: 15px;
  }

  .input-group label {
    font-weight: 600;
    margin-bottom: 5px;
    font-size: 14px;
  }

  .text-input {
    padding: 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
    font-size: 16px;
  }

  .meta-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
  }

  .hint {
    font-size: 12px;
    color: #888;
    margin-top: 4px;
  }

  /* File Drop */
  .file-drop-area {
    position: relative;
    border: 2px dashed #326cf6;
    border-radius: 8px;
    padding: 30px;
    text-align: center;
    background: #f4f8ff;
    transition: background 0.2s;
  }

  .file-drop-area:hover {
    background: #e8f0fe;
  }

  .file-input-hidden {
    position: absolute;
    left: 0;
    top: 0;
    width: 100%;
    height: 100%;
    opacity: 0;
    cursor: pointer;
  }

  .file-drop-label {
    pointer-events: none;
    font-weight: 600;
    color: #326cf6;
  }

  .file-list {
    margin-top: 20px;
    border-top: 1px solid #eee;
  }

  .file-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 12px 0;
    border-bottom: 1px solid #eee;
  }

  .file-name {
    font-weight: 600;
    display: block;
  }

  .file-size {
    font-size: 12px;
    color: #888;
  }

  .status-badge {
    font-size: 12px;
    padding: 4px 8px;
    border-radius: 12px;
    font-weight: bold;
    text-transform: uppercase;
  }

  .status-badge.pending { background: #eee; color: #555; }
  .status-badge.uploading { background: #fff4ce; color: #b58900; }
  .status-badge.success { background: #dff0d8; color: #3c763d; }
  .status-badge.error { background: #f2dede; color: #a94442; }

  .actions-row {
    margin-top: 10px;
    display: flex;
    align-items: center;
    gap: 20px;
  }

  .primary-btn {
    background-color: #326cf6;
    color: white;
    border: none;
    padding: 12px 32px;
    font-size: 16px;
    font-weight: 600;
    border-radius: 4px;
    cursor: pointer;
  }

  .primary-btn:disabled {
    background-color: #a0bbf5;
    cursor: not-allowed;
  }

  .global-message {
    font-weight: 600;
    color: #3c763d;
  }
  .global-message.error {
    color: #a94442;
  }
</style>