<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# scripts

## Purpose
Local tooling: Mac sideload automation, a static server for the built `dist/`
(used by Playwright), and a Z.AI model benchmark. None ship to production.

## Key Files
| File | Description |
|------|-------------|
| `sideload-mac.sh` | Installs a manifest into `~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/` (creates the `wef` folder if absent). Driven by `npm run sideload:mac` (dev manifest) / `sideload:prod` (downloads prod manifest). Restart Outlook after. |
| `serve-dist.mjs` | Static file server for `dist/` on port 8214 — Playwright `webServer` target. |
| `benchmark-zai-models.mjs` | Runs a prompt across Z.AI models, reports latency/quality. Ad-hoc dev utility. |

## For AI Agents

### Working In This Directory
- **Mac sideload is filesystem-based** (Outlook Mac removed the Upload UI in 2025). The script
  copies the manifest into `wef/`. If the user reports "add-in opens in a browser / sideload screen",
  the cause is usually a missing install — re-run `sideload:mac` (or `:prod`) + restart Outlook.
- **`serve-dist.mjs` port 8214 is wired into `playwright.config.mjs`.** Change both together.
- **Benchmark script sends the email/prompt to Z.AI** — it consumes the dev `ZAI_API_KEY`.
  Keep it out of CI; it's a local eyeballing tool.

### Testing Requirements
- No tests. `sideload-mac.sh` is verified by the manual sideload smoke test.
- `serve-dist.mjs` is exercised by `npm run test:e2e`.

### Common Patterns
- `.mjs` (ESM) for Node scripts; `.sh` for shell automation.
- Sideload scripts are idempotent (overwrite the manifest copy each run).

## Dependencies

### Internal
- `sideload-mac.sh` consumes `manifest.xml` (dev) or `https://alansynn.com/michael/manifest.prod.xml`.
- `serve-dist.mjs` serves `dist/` (build output).
- `benchmark-zai-models.mjs` reuses the Z.AI client conventions.

### External
- Node stdlib (`http`, `fs`) for the server; `curl` for the prod sideload download.

<!-- MANUAL: -->
