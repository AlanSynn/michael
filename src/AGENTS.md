<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# src

## Purpose
All application source for the two webpack entries. `taskpane/` is the reading-pane
UI + its Office/generation/storage wiring; `commands/` is the function-file for the
Quick Translate ribbon command; `shared/` is the single Z.AI provider client shared
by both entries (inlined per `splitChunks:false`).

## Subdirectories
| Directory | Purpose |
|-----------|---------|
| `taskpane/` | Taskpane module tree (entry `index.js` + pure leaves + `ui/` controllers). See `taskpane/AGENTS.md`. |
| `commands/` | Function-file (`commands.html` + `commands.js`) for the Quick Translate ribbon button. See `commands/AGENTS.md`. |
| `shared/` | Single Z.AI GLM client (`zai.js`) — chat-completion + /models + helpers. See `shared/AGENTS.md`. |

## For AI Agents

### Working In This Directory
- Dependency direction is **strict and cycle-free**: `index.js` → `ui/*` →
  `generation.js` → (`storage.js`, `mailbox.js`, pure leaves) → `shared/zai.js`.
  Never let a lower layer import a higher one (e.g. `generation.js` must not import DOM).
- Pure leaves (`prompts.js`, `theme.js`, `settings.js`, `model-catalog.js`, `language.js`,
  `date-context.js`, `prompt-templates.js`) must stay DOM-/Office-/generation-free so they
  remain unit-testable in plain Node.
- The Z.AI client in `shared/zai.js` is the ONLY place that calls `fetch` against
  `api.z.ai`. Both entries import it; do not duplicate the client.

### Testing Requirements
- Pure-leaf + provider: co-located `.test.mjs`, run by `npm test`.
- DOM-layer: `taskpane/ui/*.test.mjs` (jsdom).
- After changing the webpack graph, run BOTH `npm run build` and `npm run build:dev`
  (each cleans `dist/`).

### Common Patterns
- Co-located `.test.mjs` next to the module under test.
- `/* global ... */` declares host/brower globals per file (kept current — unused entries
  trip `no-undef` / `no-unused-vars`).

## Dependencies

### Internal
- `taskpane/` ↔ `commands/` share `shared/zai.js`, `taskpane/prompts.js`,
  `taskpane/language.js`, `taskpane/prompt-templates.js`.

### External
- `office.js` (host), `marked` + `dompurify` (taskpane rendering).

<!-- MANUAL: -->
