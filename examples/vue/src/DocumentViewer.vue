<!--
  <comofficekit-viewer> as a Vue component: props become the element's attributes, its events are emitted, and `file`
  (a File or Blob the user chose) is sent with open(). Other attributes, such as style, zoom, credentials or toolbar,
  fall through to the element.
-->
<script setup>
import { ref, watch } from 'vue';

const props = defineProps({
  src: String,
  documentId: String,
  file: Blob,
});
const emit = defineEmits(['load', 'error', 'pagechange']);
const element = ref(null);

watch(() => props.file, file => {
  // The element says what went wrong in its toolbar and in its error event.
  if (file) element.value.open(file).catch(() => {});
});
</script>

<template>
  <comofficekit-viewer
    ref="element"
    :src="src"
    :document="documentId"
    @load="event => emit('load', event.detail)"
    @error="event => emit('error', event.detail)"
    @pagechange="event => emit('pagechange', event.detail.page)"
  />
</template>
