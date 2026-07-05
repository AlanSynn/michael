<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# test

## Purpose
Global test setup for the `node --test` unit runner. Loaded via
`--import=./test/setup.mjs`; exists to make the webpack-injected dev globals
present in plain Node so generation/prompt modules import cleanly.

## Key Files
| File | Description |
|------|-------------|
| `setup.mjs` | Stubs `globalThis.__ZAI_API_KEY__` (dev-only global; `""` in prod builds) so modules that reference it load without a ReferenceError. |

## For AI Agents

### Working In This Directory
- **This is the `node:test` setup, not Playwright.** `npm test` runs
  `node --test --import=./test/setup.mjs`. Playwright (`npm run test:e2e`) does not use it.
- **Only globals that genuinely break import-time evaluation belong here.** Prefer per-test
  `globalThis.Office = ...` stubs (see `mailbox.test.mjs`) over piling state into `setup.mjs`.
- Keep it minimal + side-effect-free (no network, no DOM).

### Testing Requirements
- It IS the test requirement. Add a global only when a module under test fails to import without it.

### Common Patterns
- One setup module; per-test fixture stubbing at the top of each `.test.mjs`.

## Dependencies

### Internal
- Loaded by every `npm test` invocation.

### External
- Node test runner (`node:test`, `node:assert`).

<!-- MANUAL: -->
