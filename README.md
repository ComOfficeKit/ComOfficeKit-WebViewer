# ComOffice web viewer

`<comofficekit-viewer>` shows Word, Excel and PowerPoint documents in a web page. Your server draws them with the
ComOffice .NET SDK, and the page shows one image per page. The browser needs no plug-in and no Office, and it runs
none of the document's code.

The component is view only: it scrolls, zooms, moves between pages and opens the document as PDF. Editing is not
part of it.

This repository holds:

| Folder | What it is |
|---|---|
| `viewer/` | The web component: one JavaScript module, no dependencies, no build step |
| `server/` | The viewer service it calls: a small ASP.NET service on the ComOffice SDK, with a page to try it |
| `api/openapi.yaml` | The contract between the two, so any back end can serve the component |
| `examples/` | Starters: the component in an application ([examples/README.md](examples/README.md)) |

## Try it

You need .NET 8 or later on Windows x64 or Linux x64, and a ComOffice licence key. The evaluation key in
[the ComOffice for Windows README](https://github.com/ComOfficeKit/ComOfficeKit-Windows#readme) works without signing
up, until the date given there.

```bash
scripts/fetch-package.sh                    # the ComOffice .NET package, into packages/ (fetch-package.ps1 on Windows)
export COMOFFICEKIT_LICENSE_KEY='<your licence key>'
dotnet run --project server -- --urls http://localhost:5000
```

Open `http://localhost:5000/` and choose a document.

## Use the component

Serve `comofficekit-viewer.js` with your page. The service serves it too, at `/comofficekit-viewer.js`.

```html
<script type="module" src="/viewer/comofficekit-viewer.js"></script>

<comofficekit-viewer service="/viewer/" src="/files/contract.docx"></comofficekit-viewer>
```

The component fetches `src`, sends it to the service, and shows the pages as they scroll into view. To show a file
the user picks:

```js
const viewer = document.querySelector('comofficekit-viewer');
input.addEventListener('change', () => viewer.open(input.files[0]));
```

When the document is already on your server, your server can open it at the service itself and give the page only
its id. The browser then never handles the file:

```html
<comofficekit-viewer service="/viewer/" document="3f2a…"></comofficekit-viewer>
```

### Attributes

| Attribute | What it does |
|---|---|
| `service` | The viewer service's address. The page's own folder when left out |
| `src` | A document to fetch and show |
| `document` | The id of a document the service already has open. It is shown, and never uploaded or closed by the component |
| `name` | The file name to send with `src`'s document, when its address does not end in one |
| `zoom` | `1` is 100 %. Left out, a page wider than the view is fitted to its width |
| `credentials` | `include` sends cookies to the service and to `src`'s host |
| `toolbar` | `none` hides the toolbar |

### Properties, methods and events

- `open(file, name?)` sends a `File` or `Blob` and shows it. It returns a promise of `{ id, pages, kind, fonts }`.
- `close()` closes the document the component sent. Removing the element does the same.
- `goToPage(n)` scrolls to page `n`, counting from 1. The properties `zoom`, `documentId`, `pageCount` and `currentPage`
  give the state.
- The `load` event fires when a document is shown. Its detail is `{ id, pages, kind, fonts }`.
- The `error` event's detail is `{ message, status }`.
- The `pagechange` event's detail is `{ page }`.

### Styling

Style it from outside with `::part(toolbar)`, `::part(button)`, `::part(status)`, `::part(pages)` and `::part(page)`,
and with these custom properties:
- `--comofficekit-viewer-background`
- `--comofficekit-viewer-gap`
- `--comofficekit-viewer-shadow`

The element is 70 % of the window's height unless you give it a height.

## The service

`server/` answers the requests in `api/openapi.yaml`. It reads its settings from the command line, environment
variables or `appsettings.json`:

| Setting | What it does |
|---|---|
| `AllowedOrigins` | Other origins whose pages may call the service, comma-separated. None by default |
| `AllowCredentials` | `true` lets those pages send cookies |
| `WorkingDirectory` | Where documents are kept while open. The temporary folder by default |
| `MaxOpenDocuments` | 16 by default. Opening one more closes the one open longest |
| `TimeLimitSeconds` | The longest one request may keep the engine busy. 60 by default |
| `Kestrel:Limits:MaxRequestBodySize` | The largest document accepted. 30,000,000 bytes by default |

It has no sign-in. Put it behind your own before it faces the internet, and keep `AllowedOrigins` to your own sites.

The engine works on one document at a time within a process, so the service answers one request at a time. To serve
many users, run several processes and send each document's requests to the process that opened it.

### In a Linux container

```bash
scripts/fetch-package.sh
dotnet publish server -c Release -r linux-x64 --self-contained false -o publish
docker build -t comofficekit-viewer .
docker run --rm -p 8080:8080 -e COMOFFICEKIT_LICENSE_KEY comofficekit-viewer
```

The image has no fonts of its own. Documents are drawn in the two the package carries until their own fonts are
installed; the `Dockerfile` shows how.

## Licence

This repository is under the MIT licence ([LICENSE](LICENSE)). The ComOffice SDK the service uses is not open source:
it is used under its own licence, with a licence key. A licence bound to one application must name the application
`comofficekit-viewer-service`.
