<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# e2e

## Purpose
Playwright end-to-end smoke tests. Run against the built `dist/` served statically
(by `scripts/serve-dist.mjs` on port 8214). These are NOT run by `npm test`
(`node --test`); they run via `npm run test:e2e` (`playwright test`). Office is
mocked/stubbed — Outlook cannot be driven headlessly.

## Key Files
| File | Description |
|------|-------------|
| `taskpane.smoke.spec.js` | Loads the built taskpane, asserts shell renders, key controls present, no console errors. |
| `commands.smoke.spec.js` | Loads the built commands function-file, asserts it loads without 404 (guards the `splitChunks:false` orphan-chunk regression). |

## For AI Agents

### Working In This Directory
- **Build first:** `npm run build` (prod) before `npm run test:e2e`. Specs serve `dist/`.
- **Do not run these under `node --test`.** They are Playwright specs and will fail/misreport
  under the node:test runner. Use `npm run test:e2e` only.
- **Office is mocked.** Anything exercising a real `mailbox.item` flow is out of scope here;
  that path is the user's manual sideload smoke test, not自动化. Keep specs to static-load +
  render + no-error invariants.
- The commands spec is the regression guard for the function-file 404 — if you ever re-enable
  `splitChunks`, this spec breaks first. Good.

### Testing Requirements
- `npm run test:e2e` (after `npm run build`).
- `playwright.config.mjs` sets the webServer to `scripts/serve-dist.mjs` on port 8214.

### Common Patterns
- One smoke spec per entry; assert presence + absence of console errors.
- Keep Office-dependent behavior out (covered by unit tests with stubs + manual sideload).

## Dependencies

### Internal
- Builds on `dist/` (output of `npm run build`).
- `scripts/serve-dist.mjs` (Playwright `webServer`).

### External
- `@playwright/test`.

<!-- MANUAL: -->
