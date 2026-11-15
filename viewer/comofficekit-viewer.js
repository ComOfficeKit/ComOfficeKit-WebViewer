// <comofficekit-viewer>: Word, Excel and PowerPoint documents in a web page, drawn by a ComOffice viewer service as
// one image per page. View only. One file, no dependencies, no build step:
//
//   <script type="module" src="comofficekit-viewer.js"></script>
//   <comofficekit-viewer service="/viewer/" src="/files/contract.docx"></comofficekit-viewer>
//
// Attributes
//   service      the viewer service's address (api/openapi.yaml); the page's own folder when left out
//   src          a document to fetch and show: the component sends it to the service
//   document     the id of a document the service already has open (your server opened it): shown, never uploaded
//   name         the file name to send with src's document, when its address does not end in one
//   zoom         1 is 100 %; the toolbar changes it. Left out, a page wider than the view is fitted to its width
//   credentials  "include" to send cookies to the service and to src's host (fetch's credentials)
//   toolbar      "none" hides the toolbar
//   text-layer   "none" leaves out the text laid over each page, and the search box with it
//
// Each page is an image with its text laid over it, transparent: the text can be selected, copied and searched. The
// search box in the toolbar finds text in the whole document (Enter for the next match, Shift+Enter for the one before).
//
// Properties and methods: open(file, name), close(), goToPage(n), find(text), findNext(), findPrevious(), zoom,
// documentId, pageCount, currentPage.
// Events: load (detail: id, pages, kind, fonts), error (detail: message, status), pagechange (detail: page),
// find (detail: text, matches, current).
// Parts for ::part(): toolbar, button, search, status, pages, page, text-layer, match.
// Custom properties: --comofficekit-viewer-background, --comofficekit-viewer-gap, --comofficekit-viewer-shadow.

const MIN_ZOOM = 0.25;
const MAX_ZOOM = 4;
const MIN_DPI = 24;
const MAX_DPI = 300;

const STYLE = `
:host { display: block; position: relative; height: 70vh; min-height: 240px; contain: content;
        font: 14px system-ui, sans-serif; color: #1d2329;
        border: 1px solid #ced4da; border-radius: 4px; overflow: hidden; }
:host([hidden]) { display: none; }
.frame { display: flex; flex-direction: column; height: 100%; }
[part~=toolbar] { display: flex; flex-wrap: wrap; align-items: center; gap: 4px 6px; padding: 6px 8px;
                  background: #fff; border-bottom: 1px solid #ced4da; }
:host([toolbar=none]) [part~=toolbar] { display: none; }
[part~=button] { font: inherit; min-width: 32px; height: 28px; padding: 0 8px; border: 1px solid #ced4da;
                 border-radius: 4px; background: #f8f9fa; color: inherit; cursor: pointer; }
[part~=button]:hover:not(:disabled) { background: #e9ecef; }
[part~=button]:disabled { opacity: .45; cursor: default; }
[part~=button]:focus-visible, [part~=pages]:focus-visible { outline: 2px solid #1c7ed6; outline-offset: -2px; }
a[part~=button] { display: inline-flex; align-items: center; text-decoration: none; }
a[part~=button][hidden] { display: none; }
.zoom { min-width: 52px; text-align: center; font-variant-numeric: tabular-nums; }
.page-number { font-variant-numeric: tabular-nums; padding: 0 4px; white-space: nowrap; }
[part~=status] { flex: 1 1 12em; min-width: 0; color: #5c6670; overflow: hidden; text-overflow: ellipsis;
                 white-space: nowrap; }
[part~=status].problem { color: #c92a2a; white-space: normal; }
/* Centred by margins, not align-items: a page wider than the view must still scroll to its left edge. */
[part~=pages] { flex: 1; overflow: auto; display: flex; flex-direction: column; align-items: flex-start;
                gap: var(--comofficekit-viewer-gap, 16px); padding: 16px;
                background: var(--comofficekit-viewer-background, #e9ecef); }
.sheet { position: relative; flex: none; margin-inline: auto; }
[part~=page] { display: block; width: 100%; height: 100%; background: #fff;
               box-shadow: var(--comofficekit-viewer-shadow, 0 1px 4px rgba(0, 0, 0, .2)); }
/* The page's text, over its image: transparent, in place, so it can be selected and found. Built in points × 4/3,
   the page's size at 100 %, and scaled with the zoom. */
[part~=text-layer] { position: absolute; left: 0; top: 0; transform-origin: 0 0; overflow: hidden;
                     line-height: 1; font-family: sans-serif; font-size: 0; color: transparent; cursor: text;
                     user-select: text; -webkit-user-select: text; }
[part~=text-layer] .line { position: absolute; transform-origin: 0 0; white-space: pre; }
[part~=text-layer] .line > span { position: absolute; white-space: pre; transform-origin: 0 0; }
[part~=text-layer] ::selection { background: rgba(28, 126, 214, .3); color: transparent; }
[part~=match] { position: absolute; background: rgba(250, 176, 5, .35); pointer-events: none; }
[part~=match].current { background: rgba(232, 89, 12, .45); }
[part~=search] { font: inherit; width: 11em; height: 28px; padding: 0 8px; border: 1px solid #ced4da;
                 border-radius: 4px; background: #fff; color: inherit; }
[part~=search]:focus-visible { outline: 2px solid #1c7ed6; outline-offset: -1px; }
.find { display: inline-flex; align-items: center; gap: 4px; }
:host([text-layer=none]) .find { display: none; }
.matches { font-variant-numeric: tabular-nums; color: #5c6670; white-space: nowrap; }
[part~=page].waiting { background: #fff linear-gradient(90deg, #fff, #f1f3f5, #fff) 0 0 / 200% 100%;
                       animation: waiting 1.4s linear infinite; }
@keyframes waiting { to { background-position: -200% 0; } }
@media (prefers-reduced-motion: reduce) { [part~=page].waiting { animation: none; } }
`;

const MARKUP = `
<div class="frame">
  <div part="toolbar" role="toolbar" aria-label="Document">
    <button part="button" class="zoom-out" type="button" aria-label="Zoom out" title="Zoom out">−</button>
    <span class="zoom" aria-live="polite">100%</span>
    <button part="button" class="zoom-in" type="button" aria-label="Zoom in" title="Zoom in">+</button>
    <button part="button" class="fit" type="button" title="Fit the page's width">Fit width</button>
    <button part="button" class="previous" type="button" aria-label="Previous page" title="Previous page">‹</button>
    <span class="page-number">– / –</span>
    <button part="button" class="next" type="button" aria-label="Next page" title="Next page">›</button>
    <a part="button" class="pdf" target="_blank" rel="noopener" hidden>PDF</a>
    <span class="find" role="search">
      <input part="search" type="search" placeholder="Search" aria-label="Search the document" disabled>
      <button part="button" class="find-previous" type="button" aria-label="Previous match" title="Previous match" disabled>▲</button>
      <button part="button" class="find-next" type="button" aria-label="Next match" title="Next match" disabled>▼</button>
      <span class="matches" aria-live="polite"></span>
    </span>
    <span part="status" role="status"></span>
  </div>
  <div part="pages" tabindex="0" role="document" aria-label="Pages"></div>
</div>
`;

export class ComOfficeKitViewer extends HTMLElement {
  static observedAttributes = ['src', 'document', 'service', 'zoom', 'credentials', 'text-layer'];

  #root;
  #pages;
  #status;
  #choice = 0;          // each load is numbered: an answer that arrives after a newer load is set aside
  #id = null;           // the open document's id at the service
  #owned = false;       // true when this element sent the document, and so closes it at the service
  #info = null;
  #zoom = 1;
  #pageWidth = 816;     // CSS pixels at 100 %: from the first page drawn; until then, US Letter
  #pageShape = 8.5 / 11;
  #sized = false;       // true once a page has been drawn and has set #pageWidth and #pageShape
  #current = 0;
  #lazy = null;
  #scrollFrame = 0;
  #texts = new Map();   // page number → its text from the service, or the promise of it
  #layers = new Map();  // page number → the promise of its text laid over it
  #allText = null;      // the promise of every page's text, asked for by the first search
  #query = '';
  #matches = [];        // { page, line, start, length }, in reading order
  #match = -1;          // the current one, or -1
  #ascent = null;       // a sans-serif's ascent per em, as the browser draws it

  constructor() {
    super();
    this.#root = this.attachShadow({ mode: 'open' });
    this.#root.innerHTML = `<style>${STYLE}</style>${MARKUP}`;
    this.#pages = this.#root.querySelector('[part~=pages]');
    this.#status = this.#root.querySelector('[part~=status]');
    const on = (selector, action) => this.#root.querySelector(selector).addEventListener('click', action);
    on('.zoom-in', () => { this.zoom = this.#zoom * 1.25; });
    on('.zoom-out', () => { this.zoom = this.#zoom / 1.25; });
    on('.fit', () => this.#fitWidth());
    on('.previous', () => this.goToPage(this.#current - 1));
    on('.next', () => this.goToPage(this.#current + 1));
    on('.find-next', () => this.findNext());
    on('.find-previous', () => this.findPrevious());
    const search = this.#root.querySelector('[part~=search]');
    search.addEventListener('keydown', event => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      if (search.value !== this.#query) this.find(search.value).catch(() => {});
      else if (event.shiftKey) this.findPrevious();
      else this.findNext();
    });
    search.addEventListener('input', () => { if (!search.value) this.find('').catch(() => {}); });
    this.#pages.addEventListener('scroll', () => {
      if (!this.#scrollFrame) this.#scrollFrame = requestAnimationFrame(() => { this.#scrollFrame = 0; this.#track(); });
    });
    this.#toolbar();
  }

  connectedCallback() {
    if (this.hasAttribute('zoom')) this.#zoom = this.#clampZoom(Number(this.getAttribute('zoom')));
    this.#reload();
  }

  disconnectedCallback() {
    this.close();
  }

  attributeChangedCallback(name, old, value) {
    if (old === value || !this.isConnected) return;
    if (name === 'zoom') {
      this.zoom = Number(value);
    } else if (name === 'text-layer') {
      this.#refreshTextLayers();
    } else if (name !== 'credentials') {
      this.#reload();
    }
  }

  /** The viewer service's address, as a URL ending in a slash. */
  get service() {
    const base = new URL(this.getAttribute('service') || '.', document.baseURI);
    if (!base.pathname.endsWith('/')) base.pathname += '/';
    return base.href;
  }
  set service(value) { this.#setOrRemove('service', value); }

  get src() { return this.getAttribute('src') || ''; }
  set src(value) { this.#setOrRemove('src', value); }

  /** The open document's id at the service, or null. */
  get documentId() { return this.#id; }

  /** How many pages the open document has; 0 with none. */
  get pageCount() { return this.#info?.pages ?? 0; }

  /** The page in the middle of the view, from 1; 0 with no document. */
  get currentPage() { return this.#current; }

  /** 1 is 100 %. */
  get zoom() { return this.#zoom; }
  set zoom(value) {
    const zoom = this.#clampZoom(Number(value));
    if (zoom === this.#zoom) return;
    // Keep the same place in view: the share of the way down the pages stays the same.
    const along = this.#pages.scrollHeight > 0 ? this.#pages.scrollTop / this.#pages.scrollHeight : 0;
    this.#zoom = zoom;
    this.#layout();
    this.#pages.scrollTop = along * this.#pages.scrollHeight;
    this.#sharpen();
    this.#toolbar();
  }

  /**
   * Sends a file to the service and shows it. Resolves with what the service said about it: its id, pages, kind
   * and the fonts it was drawn in instead of its own. Rejects as the error event says.
   */
  async open(file, name = file?.name) {
    if (!(file instanceof Blob)) throw new TypeError('open() takes a File or a Blob');
    if (!name) throw new TypeError('open() needs the file name, for its format: open(blob, "report.docx")');
    const choice = this.#begin(`Opening ${name}…`);
    let response = null;
    let answer = null;
    try {
      response = await fetch(`${this.service}documents?name=${encodeURIComponent(name)}`,
                             { method: 'POST', body: file, credentials: this.#credentials() });
      answer = await response.json().catch(() => null);
    } catch {
      // No answer at all: the service is down or unreachable, or it closed the connection, as it may for a file
      // past its size limit.
    }
    if (choice !== this.#choice) {
      if (response?.ok && answer?.id) this.#delete(answer.id);
      return Promise.reject(new DOMException('A newer document replaced this one', 'AbortError'));
    }
    if (!response?.ok || !answer?.id) {
      const reason = !response ? 'the viewer service did not answer.'
        : answer?.detail ?? (response.status === 413 ? 'it is larger than the viewer service accepts.'
                                                     : `the viewer service answered ${response.status}.`);
      return this.#fail(`${name} did not open: ${reason}`, response?.status ?? 0);
    }
    this.#owned = true;
    return this.#show(answer, name);
  }

  /** Shows no document, and closes the one this element sent at the service. */
  close() {
    this.#choice++;
    if (this.#id && this.#owned) this.#delete(this.#id);
    this.#id = null;
    this.#owned = false;
    this.#info = null;
    this.#current = 0;
    this.#sized = false;
    this.#lazy?.disconnect();
    this.#lazy = null;
    this.#texts = new Map();
    this.#layers = new Map();
    this.#allText = null;
    this.#query = '';
    this.#matches = [];
    this.#match = -1;
    this.#root.querySelector('[part~=search]').value = '';
    this.#root.querySelector('.matches').textContent = '';
    this.#pages.replaceChildren();
    this.#say('');
    this.#toolbar();
  }

  /** Scrolls page n, counting from 1, into view. */
  goToPage(n) {
    const page = this.#pages.children[Math.round(n) - 1];
    if (!page) return;
    page.scrollIntoView({ block: 'start' });
    this.#setCurrent(Math.round(n));
  }

  /**
   * Finds text in the whole document, without regard to case, marks every match on the pages and goes to the first.
   * Resolves with how many there are. An empty string clears the search.
   */
  async find(text) {
    const query = String(text ?? '');
    const choice = this.#choice;
    this.#query = query;
    this.#root.querySelector('[part~=search]').value = query;
    this.#matches = [];
    this.#match = -1;
    if (!query || !this.#id || this.#textLayerOff()) {
      this.#markAll();
      this.#findStatus();
      return 0;
    }
    this.#root.querySelector('.matches').textContent = 'Searching…';
    const pages = await this.#everyPagesText();
    if (choice !== this.#choice || query !== this.#query) return this.#matches.length;
    const needle = query.toLocaleLowerCase();
    pages.forEach((page, i) => (page?.lines ?? []).forEach((line, l) => {
      const haystack = line.text.toLocaleLowerCase();
      for (let at = haystack.indexOf(needle); at >= 0; at = haystack.indexOf(needle, at + needle.length)) {
        this.#matches.push({ page: i + 1, line: l, start: at, length: needle.length });
      }
    }));
    // A page draws its text in the order it paints it (a Word table's cells before the paragraphs around them):
    // the matches go from the top of each page down, and along a row from the left.
    const where = match => pages[match.page - 1].lines[match.line];
    this.#matches.sort((a, b) => {
      if (a.page !== b.page) return a.page - b.page;
      const la = where(a), lb = where(b);
      if (Math.abs(la.y - lb.y) > Math.min(la.height, lb.height) / 2) return la.y - lb.y;
      return la.x - lb.x || a.line - b.line || a.start - b.start;
    });
    this.#markAll();
    if (this.#matches.length) this.#goToMatch(0);
    else this.#findStatus();
    return this.#matches.length;
  }

  /** Goes to the next match, after the last back to the first. */
  findNext() {
    if (this.#matches.length) this.#goToMatch((this.#match + 1) % this.#matches.length);
  }

  /** Goes to the match before, before the first round to the last. */
  findPrevious() {
    if (this.#matches.length) this.#goToMatch((this.#match - 1 + this.#matches.length) % this.#matches.length);
  }

  // ---- loading

  #reload() {
    const id = this.getAttribute('document');
    const src = this.getAttribute('src');
    if (id) {
      this.#attach(id).catch(() => {});
    } else if (src) {
      this.#fetchAndOpen(src).catch(() => {});
    } else {
      this.close();
    }
  }

  async #fetchAndOpen(src) {
    const url = new URL(src, document.baseURI);
    const name = this.getAttribute('name') || decodeURIComponent(url.pathname.split('/').pop() || '');
    const choice = this.#begin(`Fetching ${name || url.href}…`);
    let file;
    try {
      const response = await fetch(url, { credentials: this.#credentials() });
      if (!response.ok) throw new Error(`its server answered ${response.status}`);
      file = await response.blob();
    } catch (e) {
      if (choice !== this.#choice) return;
      return this.#fail(`${name || url.href} could not be fetched: ${e.message}.`, 0);
    }
    if (choice !== this.#choice) return;
    return this.open(file, name);
  }

  async #attach(id) {
    const choice = this.#begin('Opening…');
    let response = null;
    let answer = null;
    try {
      response = await fetch(`${this.service}documents/${encodeURIComponent(id)}`, { credentials: this.#credentials() });
      answer = await response.json().catch(() => null);
    } catch { }
    if (choice !== this.#choice) return;
    if (!response?.ok || !answer?.id) {
      const reason = !response ? 'the viewer service did not answer.'
        : answer?.detail ?? (response.status === 404 ? 'the viewer service does not have it open.'
                                                     : `the viewer service answered ${response.status}.`);
      return this.#fail(`The document did not open: ${reason}`, response?.status ?? 0);
    }
    this.#owned = false;
    return this.#show(answer, null);
  }

  // Starts a load: drops what was shown, and numbers the load.
  #begin(message) {
    this.close();
    this.#say(message);
    return this.#choice;
  }

  #fail(message, status) {
    this.#say(message, true);
    this.dispatchEvent(new CustomEvent('error', { detail: { message, status } }));
    return Promise.reject(new Error(message));
  }

  #show(info, name) {
    this.#id = info.id;
    this.#info = info;
    const fonts = (info.fonts ?? []).map(f => `${f.requested} (drawn in ${f.used})`).join(', ');
    this.#say(`${name ? `${name}: ` : ''}${info.pages} page${info.pages === 1 ? '' : 's'}`
              + (fonts ? `. Fonts the service lacks: ${fonts}` : ''));
    const pdf = this.#root.querySelector('.pdf');
    pdf.href = `${this.service}documents/${encodeURIComponent(info.id)}/pdf`;

    // Each page waits as a placeholder, and asks for its image when it comes near the view.
    this.#lazy = new IntersectionObserver(entries => {
      for (const entry of entries) {
        if (entry.isIntersecting) {
          this.#lazy.unobserve(entry.target);
          this.#request(entry.target);
        }
      }
    }, { root: this.#pages, rootMargin: '100% 0px' });
    const pages = [];
    for (let n = 1; n <= info.pages; n++) {
      // A sheet holds the page's image and, over it, the page's text.
      const page = document.createElement('div');
      page.className = 'sheet';
      page.dataset.page = String(n);
      const image = document.createElement('img');
      image.setAttribute('part', 'page');
      image.className = 'waiting';
      image.alt = `Page ${n} of ${info.pages}`;
      image.decoding = 'async';
      if (this.#credentials() === 'include') image.crossOrigin = 'use-credentials';
      image.addEventListener('load', () => this.#drawn(page));
      image.addEventListener('error', () => { image.alt = `Page ${n} of ${info.pages} could not be drawn`; });
      page.append(image);
      pages.push(page);
    }
    this.#pages.replaceChildren(...pages);
    this.#layout();
    this.#pages.scrollTop = 0;
    for (const page of pages) this.#want(page);
    this.#setCurrent(info.pages > 0 ? 1 : 0);
    this.#toolbar();
    const detail = { id: info.id, pages: info.pages, kind: info.kind, fonts: info.fonts ?? [] };
    this.dispatchEvent(new CustomEvent('load', { detail }));
    return detail;
  }

  // ---- pages

  // Sharp on any screen: the dots per inch the page is shown at, times the screen's density.
  #dpi() {
    const density = Math.min(Math.max(window.devicePixelRatio || 1, 1), 3);
    return Math.round(Math.min(Math.max(96 * this.#zoom * density, MIN_DPI), MAX_DPI));
  }

  // A page in or near the view is asked for now; one further away when it comes near. The observer alone would do,
  // but it waits for the browser to draw, which a page in a background tab never does.
  #want(page) {
    const view = this.#pages.getBoundingClientRect();
    const box = page.getBoundingClientRect();
    if (box.bottom >= view.top - view.height && box.top <= view.bottom + view.height) {
      this.#lazy.unobserve(page);
      this.#request(page);
    } else {
      this.#lazy.observe(page);
    }
  }

  #request(page) {
    const dpi = this.#dpi();
    page.dataset.dpi = String(dpi);
    page.firstElementChild.src = `${this.service}documents/${encodeURIComponent(this.#id)}/pages/${page.dataset.page}?dpi=${dpi}`;
    this.#addTextLayer(page);
  }

  #drawn(page) {
    const image = page.firstElementChild;
    const dpi = Number(page.dataset.dpi);
    image.classList.remove('waiting');
    page.dataset.width = String(image.naturalWidth * 96 / dpi);
    page.dataset.height = String(image.naturalHeight * 96 / dpi);
    // The first page drawn sets the size the pages still waiting take, and, with no zoom asked for, a page wider
    // than the view is fitted to it.
    if (!this.#sized) {
      this.#sized = true;
      this.#pageWidth = Number(page.dataset.width);
      this.#pageShape = image.naturalWidth / image.naturalHeight;
      if (!this.hasAttribute('zoom') && this.#room() > 0 && this.#pageWidth * this.#zoom > this.#room()) {
        this.#fitWidth();
        return;
      }
    }
    this.#layout();
  }

  // Sizes every page: its own size once drawn, otherwise the first drawn page's.
  #layout() {
    for (const page of this.#pages.children) {
      const width = Number(page.dataset.width) || this.#pageWidth;
      const height = Number(page.dataset.height) || width / this.#pageShape;
      page.style.width = `${width * this.#zoom}px`;
      page.style.height = `${height * this.#zoom}px`;
      const layer = page.querySelector('[part~=text-layer]');
      if (layer) layer.style.transform = `scale(${this.#zoom * width / Number(layer.dataset.width)})`;
    }
  }

  // ---- text

  #textLayerOff() { return this.getAttribute('text-layer') === 'none'; }

  // A page's text, asked for once: from every page's text when a search has asked for it, else on its own.
  #pageText(n) {
    if (!this.#texts.has(n)) {
      const url = `${this.service}documents/${encodeURIComponent(this.#id)}/pages/${n}/text`;
      this.#texts.set(n, fetch(url, { credentials: this.#credentials() })
        .then(response => (response.ok ? response.json() : null)).catch(() => null)
        .then(text => { this.#texts.set(n, text); return text; }));
    }
    return Promise.resolve(this.#texts.get(n));
  }

  #everyPagesText() {
    if (!this.#allText) {
      const url = `${this.service}documents/${encodeURIComponent(this.#id)}/text`;
      this.#allText = fetch(url, { credentials: this.#credentials() })
        .then(response => (response.ok ? response.json() : null)).catch(() => null)
        .then(answer => {
          const pages = answer?.pages ?? [];
          pages.forEach((text, i) => {
            if (!this.#texts.has(i + 1) || this.#texts.get(i + 1) instanceof Promise) this.#texts.set(i + 1, text);
          });
          return pages;
        });
    }
    return this.#allText;
  }

  // Lays the page's text over it, once: a second call waits for the first.
  #addTextLayer(page) {
    const n = Number(page.dataset.page);
    if (this.#textLayerOff()) return Promise.resolve();
    if (!this.#layers.has(n)) {
      const choice = this.#choice;
      this.#layers.set(n, this.#pageText(n).then(text => {
        if (choice !== this.#choice || !text || this.#textLayerOff() || !page.isConnected
            || page.querySelector('[part~=text-layer]')) return;
        page.append(this.#buildTextLayer(text));
        this.#layout();
        this.#markAll();
      }));
    }
    return this.#layers.get(n);
  }

  #refreshTextLayers() {
    this.#layers = new Map();
    for (const page of this.#pages.children) {
      page.querySelector('[part~=text-layer]')?.remove();
      if (!this.#textLayerOff() && page.dataset.dpi) this.#addTextLayer(page);
    }
    if (this.#textLayerOff()) this.find('').catch(() => {});
    this.#toolbar();
  }

  // The page's text as the service gave it, in points, laid out in CSS pixels at 100 % (a point is 4/3 of one):
  // each line at its start and turn, each span stretched to the width it is drawn at.
  #buildTextLayer(text) {
    const k = 4 / 3;
    const layer = document.createElement('div');
    layer.setAttribute('part', 'text-layer');
    layer.dataset.width = String(text.width * k);
    layer.style.width = `${text.width * k}px`;
    layer.style.height = `${text.height * k}px`;
    const measure = document.createElement('canvas').getContext('2d');
    const ascent = this.#ascentRatio(measure);
    text.lines.forEach((line, l) => {
      const row = document.createElement('div');
      row.className = 'line';
      row.dataset.line = String(l);
      row.style.left = `${line.x * k}px`;
      row.style.top = `${line.baseline * k}px`;
      if (line.rotation) row.style.transform = `rotate(${line.rotation}deg)`;
      let end = null;
      const piece = (content, x, width, baseline, size) => {
        const span = document.createElement('span');
        span.textContent = content;
        const px = Math.max(size * k, 1);
        span.style.fontSize = `${px}px`;
        span.style.left = `${(x - line.x) * k}px`;
        span.style.top = `${(baseline - line.baseline) * k - ascent * px}px`;
        measure.font = `${px}px sans-serif`;
        const natural = measure.measureText(content).width;
        if (natural > 0 && width > 0) span.style.transform = `scaleX(${(width * k) / natural})`;
        row.append(span);
      };
      let at = 0;
      line.spans.forEach((span, i) => {
        // A space the engine put between two spans drawn apart: across the gap, so copied text keeps it.
        if (span.start > at && end !== null) {
          piece(line.text.slice(at, span.start), end.x, Math.max(span.x - end.x, 0), end.baseline, end.size);
        }
        // A span drawn a little into the next ends where the next starts, so a selection never shades the same
        // place twice.
        const next = line.spans[i + 1];
        const width = next && next.x > span.x ? Math.min(span.width, next.x - span.x) : span.width;
        piece(line.text.slice(span.start, span.start + span.length), span.x, width, span.baseline, span.fontSize);
        end = { x: span.x + width, baseline: span.baseline, size: span.fontSize };
        at = span.start + span.length;
      });
      layer.append(row, document.createElement('br'));
    });
    return layer;
  }

  #ascentRatio(measure) {
    if (this.#ascent === null) {
      measure.font = '100px sans-serif';
      const metrics = measure.measureText('Hg');
      this.#ascent = metrics.fontBoundingBoxAscent ? metrics.fontBoundingBoxAscent / 100 : 0.8;
    }
    return this.#ascent;
  }

  // Marks every match on the pages whose text is laid out, the current one apart.
  #markAll() {
    for (const mark of this.#root.querySelectorAll('[part~=match]')) mark.remove();
    if (!this.#matches.length) return;
    const k = 4 / 3;
    this.#matches.forEach((match, i) => {
      const page = this.#pages.children[match.page - 1];
      const row = page?.querySelector(`[part~=text-layer] .line[data-line="${match.line}"]`);
      const line = this.#texts.get(match.page)?.lines?.[match.line];
      if (!row || !line) return;
      for (const span of line.spans) {
        const from = Math.max(match.start, span.start);
        const to = Math.min(match.start + match.length, span.start + span.length);
        if (from >= to) continue;
        const before = span.advances.slice(0, from - span.start).reduce((a, b) => a + b, 0);
        const width = span.advances.slice(from - span.start, to - span.start).reduce((a, b) => a + b, 0);
        const mark = document.createElement('span');
        mark.setAttribute('part', 'match');
        if (i === this.#match) mark.classList.add('current');
        mark.style.left = `${(span.x + before - line.x) * k}px`;
        mark.style.width = `${width * k}px`;
        mark.style.top = `${(line.y - line.baseline) * k}px`;
        mark.style.height = `${line.height * k}px`;
        row.append(mark);
      }
    });
  }

  async #goToMatch(i) {
    this.#match = i;
    const match = this.#matches[i];
    const page = this.#pages.children[match.page - 1];
    if (!page) return;
    if (!page.dataset.dpi) { this.#lazy?.unobserve(page); this.#request(page); }
    await this.#addTextLayer(page);
    this.#markAll();
    this.#findStatus();
    const mark = page.querySelector('[part~=match].current');
    if (mark) mark.scrollIntoView({ block: 'center', inline: 'nearest' });
    else page.scrollIntoView({ block: 'start' });
    this.#setCurrent(match.page);
  }

  #findStatus() {
    const count = this.#matches.length;
    this.#root.querySelector('.matches').textContent =
      !this.#query ? '' : count ? `${this.#match + 1} / ${count}` : 'No matches';
    this.#toolbar();
    if (this.#query) {
      this.dispatchEvent(new CustomEvent('find', { detail: { text: this.#query, matches: count, current: this.#match + 1 } }));
    }
  }

  // After zooming in, the pages already drawn are asked for again at the sharper size, as they come into view.
  #sharpen() {
    if (!this.#lazy) return;
    const dpi = this.#dpi();
    for (const page of this.#pages.children) {
      if (page.dataset.dpi && Number(page.dataset.dpi) < dpi * 0.9) {
        delete page.dataset.dpi;
        this.#want(page);
      }
    }
  }

  // The width a page has in the view, inside the padding.
  #room() {
    const styles = getComputedStyle(this.#pages);
    return this.#pages.clientWidth - parseFloat(styles.paddingLeft) - parseFloat(styles.paddingRight);
  }

  #fitWidth() {
    const room = this.#room();
    if (room > 0) this.zoom = room / this.#pageWidth;
  }

  // The current page is the one across the middle of the view.
  #track() {
    const view = this.#pages.getBoundingClientRect();
    const middle = view.top + view.height / 2;
    let best = 0;
    let distance = Infinity;
    for (const page of this.#pages.children) {
      const box = page.getBoundingClientRect();
      const d = box.top <= middle && box.bottom >= middle ? 0 : Math.min(Math.abs(box.top - middle), Math.abs(box.bottom - middle));
      if (d < distance) { distance = d; best = Number(page.dataset.page); }
      if (d === 0) break;
    }
    this.#setCurrent(best);
  }

  #setCurrent(n) {
    if (n === this.#current) return;
    this.#current = n;
    this.#toolbar();
    if (n > 0) this.dispatchEvent(new CustomEvent('pagechange', { detail: { page: n } }));
  }

  // ---- toolbar and status

  #toolbar() {
    const $ = selector => this.#root.querySelector(selector);
    const open = !!this.#info;
    $('.zoom').textContent = `${Math.round(this.#zoom * 100)}%`;
    $('.zoom-out').disabled = !open || this.#zoom <= MIN_ZOOM;
    $('.zoom-in').disabled = !open || this.#zoom >= MAX_ZOOM;
    $('.fit').disabled = !open;
    $('.previous').disabled = !open || this.#current <= 1;
    $('.next').disabled = !open || this.#current >= this.pageCount;
    $('.page-number').textContent = open ? `${this.#current} / ${this.pageCount}` : '– / –';
    $('.pdf').hidden = !open;
    $('[part~=search]').disabled = !open || this.#textLayerOff();
    $('.find-next').disabled = $('.find-previous').disabled = !this.#matches.length;
  }

  #say(message, problem = false) {
    this.#status.textContent = message;
    this.#status.classList.toggle('problem', problem);
  }

  // A framework may set a property to null or undefined: that removes the attribute, rather than writing the word.
  #setOrRemove(name, value) {
    if (value == null || value === '') this.removeAttribute(name);
    else this.setAttribute(name, String(value));
  }

  #credentials() {
    return this.getAttribute('credentials') === 'include' ? 'include' : 'same-origin';
  }

  #clampZoom(value) {
    return Number.isFinite(value) && value > 0 ? Math.min(Math.max(value, MIN_ZOOM), MAX_ZOOM) : 1;
  }

  #delete(id) {
    fetch(`${this.service}documents/${encodeURIComponent(id)}`, { method: 'DELETE', credentials: this.#credentials(),
                                                                   keepalive: true }).catch(() => {});
  }
}

if (!customElements.get('comofficekit-viewer')) customElements.define('comofficekit-viewer', ComOfficeKitViewer);
