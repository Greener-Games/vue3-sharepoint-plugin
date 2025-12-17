<template>
<!--  <div>
    <a href="https://vite.dev" target="_blank">
      <img src="/vite.svg" class="logo" alt="Vite logo" />
    </a>
    <a href="https://vuejs.org/" target="_blank">
      <img src="./assets/vue.svg" class="logo vue" alt="Vue logo" />
    </a>
  </div>
  <HelloWorld msg="Vite + Vue" />-->
  <cv-search/>
  <document-upload-test/>
  <my-documents-dashboard view-mode="myteam"/>
  <my-documents-dashboard view-mode="mydocuments"/>
</template>

<script setup lang="ts">
import CvSearch from "@/components/CvSearch.vue";
import DocumentUploadTest from "@/components/DocumentUploadTest.vue";
import MyDocumentsDashboard from "@/components/MyDocumentsDashboard.vue";
</script>

<style scoped>
.logo {
  height: 6em;
  padding: 1.5em;
  will-change: filter;
  transition: filter 300ms;
}
.logo:hover {
  filter: drop-shadow(0 0 2em #646cffaa);
}
.logo.vue:hover {
  filter: drop-shadow(0 0 2em #42b883aa);
}
</style>
