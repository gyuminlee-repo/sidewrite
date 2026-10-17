# taskpane/

## OVERVIEW
Browser code loaded inside Word/Excel (and the standalone figures page), served as static files by the daemon HTTP server. Earned its file: distinct runtime (Office.js WebView, no Node), own CSP and asset-path rules; heavy logic lives in `shared/` (own AGENTS.md).

## STRUCTURE
```
taskpane/
├── word/index.html     # Word host page (719 lines): tabs review, response, tools, chat, library, setup
├── excel/index.html    # Excel host page (226 lines): tabs chat, library, setup
├── figures/            # standalone figures page: index.html, figures.js (888), figures.css
├── shared/             # host-aware JS + styles.css (see shared/AGENTS.md)
└── icon-32.png, icon-64.png
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Add a button/tab to Word pane | `word/index.html` + `shared/*-pane.js` | DOM ids are asserted by `tests/*-dom.test.mjs` |
| Pane asset serving, cache busting | `daemon/index.mjs` ~l.102-250 | `/shared/taskpane.js` and `styles.css` get `?v=<bridge token>` rewritten into HTML |
| Figures page | `figures/` + `daemon/figure-http.mjs` | route `/figures/`, API behind bridge token |
| Figures page tests | `tests/figure-link.test.mjs`, `tests/helpers/figures-page.mjs` | fake DOM |

## CONVENTIONS
- Asset paths are route-root absolute: `/shared/...`, `/figures/...`. Never relative `../`.
- Native ES modules with explicit `.js` suffix; no bundler, no build step, no npm imports in browser code.
- Daemon serves `taskpane/` with a strict under-root check (`fsPath === dir || startsWith(dir + sep)`); new top-level folders here are publicly served on 127.0.0.1.
- Ports: HTTP 47834, WebSocket 47833 (`WS_PORT = HTTP_PORT - 1`); the WS port is a literal in `shared/taskpane.js` and in the CSP `connect-src`.
- Host pages carry a meta CSP: `script-src 'self' https://appsforoffice.microsoft.com`, WS only to `127.0.0.1:47833`, no object/base URI.
- Bridge token: pane fetches it from the daemon token endpoint; figures API rejects missing/wrong token, foreign Origin/Host.

## ANTI-PATTERNS
- Never drop the CSP to fix a blank pane; widen only the needed directive and check in both Word and Excel.
- Never inject daemon or document text with `innerHTML`; use `textContent`. Existing `innerHTML` uses in `shared/taskpane.js` are fixed templates or clears only; fill dynamic slots afterwards with `textContent`.
- Never import a daemon module from pane code; the pane cannot reach `daemon/`. Duplicated vocabulary (for example in `shared/response-view.js`) must be changed on both sides.
- Never rename a DOM id without grepping `tests/`; DOM tests look ids up by literal string.
- Do not edit `manifests/*.xml` to point at a new page path without updating sideload (`app/sideload.mjs`) expectations; Office caches manifests.

## VERIFY
```bash
node --check taskpane/shared/*.js taskpane/figures/figures.js
node --test tests/*-dom.test.mjs tests/figure-link.test.mjs
```
Real Office behavior is only covered by `scripts/e2e/` (Windows Word required).
