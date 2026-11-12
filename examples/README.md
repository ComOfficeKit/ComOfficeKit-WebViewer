# Starters

Small pages that put `<comofficekit-viewer>` in an application. Each one calls a viewer service: run `server/` first,
as [the README](../README.md#try-it) says.

| Folder | What it shows |
|---|---|
| [html](html) | A plain page on its own site: the component loaded from the service, a file the user picks, a document at an address, the events |
| [react](react) | A React app with a `DocumentViewer` component that wraps the element: props in, callbacks out |
| [vue](vue) | A Vue app with a `DocumentViewer` component that wraps the element: props in, events out |

## html

The page is on one site and the service on another, so the service must allow the page's origin:

```bash
# the service, in one terminal
dotnet run --project server -- --urls http://localhost:5000 --AllowedOrigins http://localhost:8000

# the page, in another
cd examples/html && python3 -m http.server 8000 --bind 127.0.0.1
```

Open `http://localhost:8000/`. The page uses `http://localhost:5000/` as the service, or the address in `?service=`.

A document shown by its address is fetched by the browser, so its own server must allow the page's origin too.

## react and vue

Each is a [Vite](https://vite.dev) app (Node 20.19 or later), its packages pinned by `package-lock.json`. In development and preview, Vite passes `/documents`
and `/comofficekit-viewer.js` on to the viewer service. The page and the service then share an origin, so there is
no CORS to set up.

```bash
# the service, in one terminal
dotnet run --project server -- --urls http://localhost:5000

# the app, in another
cd examples/react        # or examples/vue
npm ci                   # the versions package-lock.json pins
npm run dev
```

`VIEWER_SERVICE=http://host:port` points the app at a service elsewhere. Once built (`npm run build`), the app is
static files in `dist/`. Have the server that serves them pass the same two paths on to the viewer service.

`DocumentViewer` takes the element's attributes as props:
- `src`;
- `documentId` (the `document` attribute);
- anything else, such as `zoom` or a style.

It also takes `file`, which is sent with `open()`. Its callbacks are:
- React: `onLoad`, `onError` and `onPageChange`;
- Vue: the events `load`, `error` and `pagechange`.

Vue is told that `comofficekit-*` tags are custom elements in `vite.config.js`.
