<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# Michael — Outlook Add-in (Z.AI GLM)

## Purpose
"Michael" is an Outlook add-in that adds Z.AI GLM-powered email actions —
summarize, translate, translate+summarize, draft reply, and calendar-event
extraction — to the Outlook reading pane. Deployed at `alansynn.com/michael/`.
The add-in is **English internally**; the chat default language is **Korean**
(configurable). The same origin also serves a `commands.html` function-file for
the Quick Translate ribbon command.

This is an Office Add-in (XML manifest, not unified JSON — required for classic
Outlook Mac). The taskpane runs inside the Office webview; `office.js` is loaded
from Microsoft CDNs; all provider traffic goes to `api.z.ai`.

## Key Files
| File | Description |
|------|-------------|
| `manifest.xml` | Dev XML manifest (origin `https://localhost:3000`). webpack rewrites it to prod (`alansynn.com/michael/`) + strips the localhost AppDomain. |
| `webpack.config.mjs` | Two entries (`taskpane`, `commands`); `splitChunks:false`; `DefinePlugin` dev-gates `__ZAI_API_KEY__` + `__ZAI_CODING_BASE_URL__`; prod manifest transform; asset/font copiers. |
| `package.json` | Scripts: `build` / `build:dev` / `lint` / `test` (node:test) / `test:e2e` (Playwright) / `validate` / `sideload:mac` / `sideload:prod`. |
| `babel.config.json` | Bare `@babel/preset-env`, syntax-only (no polyfills — modern webview). |
| `playwright.config.mjs` | E2E against the built `dist/` served by `scripts/serve-dist.mjs` on port 8214. |
| `README.md` | Setup, CSP, sideload, deploy. `CHANGELOG.md` + `CONTRIBUTING.md` are project records. |

## Subdirectories
| Directory | Purpose |
|-----------|---------|
| `src/` | All application source (taskpane + commands + shared Z.AI client). See `src/AGENTS.md`. |
| `assets/` | PNG/SVG icons + the meet-michael banners referenced by manifest + HTML. See `assets/AGENTS.md`. |
| `e2e/` | Playwright smoke specs (static + mocked-Office). See `e2e/AGENTS.md`. |
| `scripts/` | Mac sideload, dist static server, Z.AI model benchmark. See `scripts/AGENTS.md`. |
| `test/` | node:test global setup (`__ZAI_API_KEY__` stub). See `test/AGENTS.md`. |
| `.github/` | CI/deploy pipeline. See `.github/AGENTS.md`. |

## For AI Agents

### Working In This Directory
- **Security invariants (do not break):** the Z.AI API key is entered at runtime
  (Settings → General) and stored in `Office.context.roamingSettings` (Office-managed).
  Never log it, never put it in a URL/query, never persist it to localStorage. It is
  sent only as `Authorization: Bearer <key>`. The build-time `ZAI_API_KEY` env var is
  **dev-only**; production builds embed `""`.
- **ESM throughout.** `package.json` is `"type":"module"`. Tests are `.test.mjs`.
- **Two-entry webpack graph with `splitChunks:false`** — do not re-enable splitting;
  it produces an orphan chunk + a broken `<script src>` in `commands.html` (the
  function-file 404s). Shared modules inline into both entries.
- **manifest.xml is the DEV manifest.** Never hand-edit prod URLs; the prod transform
  in `webpack.config.mjs` rewrites `https://localhost:3000/` → `https://alansynn.com/michael/`
  and drops the localhost AppDomain. Validate the *emitted* prod manifest:
  `npx office-addin-manifest validate dist/manifest.prod.xml`.
- **DOMPurify + marked are bundled (never CDN).** All LLM/email output reaches the DOM
  through `renderMarkdown()` (`src/taskpane/ui/dom.js`) — the single sanitization boundary.
- Direct writes are OK for `~/.claude/**`, `.omc/**`, `.claude/**`, `CLAUDE.md`, `AGENTS.md`.

### Testing Requirements
- `npm run lint` — office-addin-lint (eslint + prettier). When `--fix` is unsupported,
  `npx prettier --write <file>` resolves prettier/prettier errors.
- `npm test` — `node --test --import=./test/setup.mjs`. Pure-leaf modules have co-located
  `.test.mjs`; DOM-layer tests live in `src/taskpane/ui/*.test.mjs` (jsdom).
- `npm run test:e2e` — Playwright; requires `npm run build` first (serves `dist/`).
- `npm run validate` — manifest schema (dev). Also validate the prod manifest post-build.
- Gate before commit: lint + unit + build (prod AND dev) + e2e + manifest validate.

### Common Patterns
- **Pure-leaf modules** for testability (`theme.js`, `prompts.js`, `date-context.js`,
  `settings.js`, `model-catalog.js`, `language.js`). Keep DOM/Office/generation out of them.
- **Single Office boundary**: all `Office.context.roamingSettings` access goes through
  `src/taskpane/storage.js`; all `mailbox.item` access through `src/taskpane/mailbox.js`.
- **Single-flight generation**: `src/taskpane/ui/flow-control.js` owns the one
  `AbortController` shared by every email flow + the calendar parse — start a new flow
  to abort the prior; guard every DOM mutation (including catch paths) with `isActiveFlow`.
- **Attr-driven result type**: `#result-section` carries `data-result-type` (+ optional
  `data-reply-raw`); copy logic reads these instead of inferring from rendered text.
- **CSP is strict**: `script-src 'self' + office.js CDNs` (no `'unsafe-inline'` for script);
  `'unsafe-inline'` only for `style-src` (Fluent UI). `connect-src` scoped to `api.z.ai`
  + ARIA telemetry.

## Dependencies

### External
- `office.js` (Microsoft CDN) — Office Add-in host API.
- `webpack` 5.108 — bundler. `html-loader`, `asset/resource`, `CopyWebpackPlugin`, `HtmlWebpackPlugin`.
- `marked` + `dompurify` (bundled) — markdown → sanitized HTML.
- `node:test` — unit runner (no framework). `jsdom` for DOM-layer tests.
- `@playwright/test` — e2e.
- `office-addin-lint` / `office-addin-manifest` / `office-addin-dev-certs` — tooling.
- `@fontsource/inter` — self-hosted Latin font (Korean falls back to system font).

<!-- MANUAL: Project-specific notes below this line are preserved on regeneration. -->
