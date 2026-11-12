import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// The viewer service: server/ in this repository, or any service that answers api/openapi.yaml. The page calls it on
// its own origin, under /documents, and loads the component from it, so there is no CORS to set up. In production,
// route the same two paths to the service.
const service = process.env.VIEWER_SERVICE ?? 'http://localhost:5000';
const proxy = { '/documents': service, '/comofficekit-viewer.js': service };

export default defineConfig({
  plugins: [react()],
  server: { proxy },
  preview: { proxy },
});
