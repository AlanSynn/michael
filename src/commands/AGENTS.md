<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# commands

## Purpose
Function-file for the Quick Translate ribbon command. `commands.html` is the
invisible page Office loads for the ribbon button; `commands.js` registers the
handler, translates the selected email body via Z.AI, replaces the selection, and
calls `event.completed()` (in a `finally`) so the ribbon never hangs.

## Key Files
| File | Description |
|------|-------------|
| `commands.html` | Function-file page (no UI). `<script src>` is webpack-emitted; keep `splitChunks:false` or the orphan-chunk 404 returns. |
| `commands.js` | `Office.actions.associate` (or `Office.initialize`) → handler. Reads body, calls Z.AI via shared client, `replaceSelectionWithText`, `event.completed()`. |

## For AI Agents

### Working In This Directory
- **`event.completed()` must run on every path** — wrap it in `finally`. A missed
  `completed()` leaves the ribbon button in its pressed/spinning state (R7 was this bug).
- **Reuse the shared modules, do not fork a second Z.AI client.** Import `generateText` (or
  equivalent) from `../shared/zai.js`, `getLanguageText` from `../taskpane/language.js`,
  `fillTemplate` from `../taskpane/prompts.js`, and `commandTranslate` from
  `../taskpane/prompt-templates.js`. The old CJS `zaiClient.js`/`zaiConfig.js` were deleted
  for exactly this reason — do not resurrect them.
- **`commands.js` is ESM now.** Webpack bundles it into the function-file chunk; both
  entries inline the shared modules (no `splitChunks`).
- The handler may run with no selection — fail soft (notify + `completed()`), don't throw uncaught.

### Testing Requirements
- No co-located unit test (handler needs the Office host). Covered by `e2e/commands.smoke.spec.js`
  against the built bundle.
- After any change: `npm run build` (and `build:dev`) — both entries must emit cleanly.

### Common Patterns
- `Office.actions.associate("CommandName", handler)` for the modern host; keep an
  `Office.initialize` fallback for classic Mac.
- `finally { event.completed(); }`.

## Dependencies

### Internal
- `../shared/zai.js`, `../taskpane/language.js`, `../taskpane/prompts.js`,
  `../taskpane/prompt-templates.js`, `../taskpane/mailbox.js` (`replaceSelectionWithText`).

### External
- `office.js` (FunctionFile host API).

<!-- MANUAL: -->
