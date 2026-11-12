<script setup>
import { ref } from 'vue';
import DocumentViewer from './DocumentViewer.vue';

const file = ref(null);
const status = ref('Choose a document.');
const page = ref(0);

function loaded(info) {
  status.value = `${info.pages} pages`;
  page.value = 1;
}
function failed(problem) {
  status.value = problem.message;
  page.value = 0;
}
</script>

<template>
  <main style="max-width: 1100px; margin: 0 auto; padding: 16px; font: 15px system-ui, sans-serif">
    <h1 style="font-size: 20px">ComOffice viewer in Vue</h1>
    <p>
      <input type="file" accept=".docx,.xlsx,.xlsm,.pptx,.csv,.tsv" @change="e => (file = e.target.files[0] ?? null)" />
      {{ ' ' }}<span role="status">{{ status }}{{ page > 0 ? `, page ${page}` : '' }}</span>
    </p>
    <DocumentViewer :file="file" style="height: 70vh" @load="loaded" @error="failed" @pagechange="p => (page = p)" />
  </main>
</template>
