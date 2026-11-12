import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

// The viewer service: server/ in this repository, or any service that answers api/openapi.yaml. The page calls it on
// its own origin, under /documents, and loads the component from it, so there is no CORS to set up. In production,
// route the same two paths to the service.
const service = process.env.VIEWER_SERVICE ?? 'http://localhost:5000';
const proxy = { '/documents': service, '/comofficekit-viewer.js': service };

export default defineConfig({
  // <comofficekit-viewer> is a custom element, not a Vue component.
  plugins: [vue({ template: { compilerOptions: { isCustomElement: tag => tag.startsWith('comofficekit-') } } })],
  server: { proxy },
  preview: { proxy },
});
