<template>
  <div class="app-container">
    <examples-nav v-model="currentExample" />

    <div class="content-area">
      <keep-alive>
        <component :is="currentComponent" v-bind="componentProps" />
      </keep-alive>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import ExamplesNav from "@/components/ExamplesNav.vue";
import CvSearch from "@/components/CvSearch.vue";
import DocumentUploadTest from "@/components/DocumentUploadTest.vue";
import MyDocumentsDashboard from "@/components/MyDocumentsDashboard.vue";
import SystemInfo from "@/components/SystemInfo.vue";

const currentExample = ref('search')

const currentComponent = computed(() => {
  switch (currentExample.value) {
    case 'search': return CvSearch
    case 'upload': return DocumentUploadTest
    case 'mydocs': return MyDocumentsDashboard
    case 'myteam': return MyDocumentsDashboard
    case 'system': return SystemInfo
    default: return CvSearch
  }
})

const componentProps = computed(() => {
  if (currentExample.value === 'mydocs') {
    return { viewMode: 'mydocuments' }
  }
  if (currentExample.value === 'myteam') {
    return { viewMode: 'myteam' }
  }
  return {}
})
</script>

<style scoped>
.app-container {
  min-height: 100vh;
  background-color: #f8fafc;
}

.content-area {
  padding-bottom: 40px;
}
</style>
