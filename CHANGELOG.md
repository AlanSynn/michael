# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] — 2026-07-05

Comprehensive overhaul: operability fixes, modular refactor, settings redesign,
hardened security posture, modernized CI, and a real test suite.

### Added
- Mac sideload automation (`scripts/sideload-mac.sh`, `npm run sideload:mac` /
  `sideload:prod`) targeting Outlook's 2025 `wef`-folder install path.
- Manifest fix: the previously-undefined `Commands.Url` resource is now declared,
  so the Quick Translate ribbon command resolves.
- Content-Security-Policy meta tag on the taskpane (defense-in-depth).
- Test suite: `node:test` unit tests for the pure leaves + `shared/zai.js`,
  jsdom DOM-interaction tests against the real `taskpane.html`, and a Playwright
  e2e smoke suite (CSP/console-error gate + an office.js-mocked integration test).
- CI pipeline (`ci.yml`): validate → lint → unit → build → e2e, then deploy to
  GitHub Pages on main. Playwright browser cached; report artifact uploaded.
- `dependabot.yml` for npm + GitHub Actions.
- `LICENSE`, `CHANGELOG.md`, `CONTRIBUTING.md`.

### Changed
- Decomposed the 2,200-line `taskpane.js` monolith into a layered module tree
  (`index.js` → `ui/*` → `generation.js` → `storage.js` / `mailbox.js` → pure
  leaves → `shared/zai.js`), acyclic, single-boundary facades.
- Settings panel redesigned into General / Templates / Developer tabs with the
  existing theme tokens; booleans use segmented toggles, templates use
  collapsible accordions.
- API key now lives only in `Office.context.roamingSettings` (Office-managed);
  it is never logged or sent anywhere except the Z.AI API.
- `marked` + `DOMPurify` are bundled (no CDN/SRI), eliminating the CDN-failure
  failure class; `renderMarkdown` is the single sanitizing trust boundary.
- Webpack config converted to ESM (`webpack.config.mjs`) with `"type": "module"`.
- CI bumped to Node 22 with `actions/*@v4` and `peaceiris/actions-gh-pages@v4`.

### Removed
- Duplicate Z.AI client (`zaiClient.js`, `zaiConfig.js`) — single ESM provider.
- Dead code: `fonts.js`, `ZAI_PROVIDER_CONFIG`, the `TYPES.CALENDAR` enum member,
  redundant `DefinePlugin`/`devServer` duplicates, stale `initializeApp` paths.
- `pnpm-lock.yaml` — standardized on npm (`package-lock.json`).

### Security
- Sanitization is fail-closed and library-bundled; CSP restricts `connect-src`
  to `https://api.z.ai`. See `CONTRIBUTING.md` → Security.

## [1.0.0]

Initial public release of the Michael Outlook add-in.
