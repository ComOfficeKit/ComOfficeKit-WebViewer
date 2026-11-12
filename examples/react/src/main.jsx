import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';

// The component comes from the viewer service, so it is always the version that service answers to. It is loaded at
// run time, not bundled: the address is a variable, which Vite leaves alone.
const component = '/comofficekit-viewer.js';
await import(/* @vite-ignore */ component);

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
