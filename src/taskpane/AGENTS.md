<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# taskpane

## Purpose
The reading-pane module tree + its webpack entry. `index.js` boots on
`Office.onReady`, wires DOM listeners + Office events, then delegates each user
action to a controller under `ui/`. Pure leaves hold all language, theme,
prompt-template, settings-normalization, and model-catalog logic so it stays
unit-testable without Office or a DOM.

## Key Files
| File | Description |
|------|-------------|
| `index.js` | Entry. `Office.onReady` (guarded), `wireListeners()` (idempotent flag), api-key input debounce → model refresh, flat ItemChanged (→ autorun + clearEmailContentCache + calendar button). |
| `generation.js` | Z.AI orchestration: `generateContent`, `generateTldrContent` (routes through `fillTemplate`), `refreshModelCatalog`, `requireModel`/`requireTemplate`. |
| `storage.js` | ONE `Office.context.roamingSettings` boundary: settings get/save, migration, template defaults, model-catalog persistence. |
| `mailbox.js` | Async wrappers over `mailbox.item` (body/subject/replace) + flat `onItemChanged`/`onSettingsChanged`. Per-itemId in-flight body cache. |
| `prompts.js` | `fillTemplate(tpl, vars)` — the single `$`-safe template filler (raw `.replace` corrupts `$&`/`$$`/`` $` ``/`$'`). |
| `settings.js` | `mergeWithDefaults`, `normalizeSettings`, `isAutorunEnabled`, `getAutorunAction` (kills the 3× autorun switch). |
| `model-catalog.js` | Mutable model list + cache helpers; `boundModels` dedupes + caps before bounding. |
| `theme.js` | `isDarkTheme`, `resolveTheme` (single fallback — fixes the prior dark/system mismatch). |
| `language.js` | `getLanguageText` (single source — was duplicated in taskpane + commands). |
| `fonts.js` | Font-size scale + apply. |
| `date-context.js` | `getCurrentDateContext()` — local current-date injected into calendar prompts (and AOE fallback lives in here / calendar path). |
| `prompt-templates.js` | `DEFAULT_SETTINGS`, `DEFAULT_PROMPT_TEMPLATES`, factories. `calendarParse` body uses plain `Text` (displayNewAppointmentForm). |
| `taskpane.html` / `taskpane.css` | UI shell + Settings panel; CSS tokens in `:root` / `[data-theme="dark"]` reused unchanged across redesigns. |

## Subdirectories
| Directory | Purpose |
|-----------|---------|
| `ui/` | DOM controllers: flows, calendar, settings-view, dom, flow-control. See `ui/AGENTS.md`. |

## For AI Agents

### Working In This Directory
- **Keep the dependency layering intact:** `index.js` → `ui/*` → `generation.js` →
  (`storage.js`, `mailbox.js`) → pure leaves → `shared/zai.js`. Leaves never import DOM/Office/generation.
- **Single Office boundary:** all `roamingSettings` through `storage.js`; all
  `mailbox.item` through `mailbox.js`. Do not reach into `Office.context` elsewhere.
- **Template filling must go through `fillTemplate`.** Never use a raw
  `template.replace("{{x}}", value)` chain — `$&`, `$$`, `` $` ``, `$'` inside email
  bodies corrupt the output. `generateTldrContent` was the regression site; keep it routed.
- **`getEmailContent()` caches the in-flight Promise per itemId.** Call `clearEmailContentCache()`
  on ItemChanged so a stale body never leaks forward. Rejected promises self-evict for retry.
- **Settings reflection:** theme, font-size, showReply, tldrMode, autorun must visibly
  apply on save AND reload. If you add a setting, wire both the apply path and persistence.
- **Autorun switch lives once** in `settings.js` (`getAutorunAction`); do not re-fork it per flow.

### Testing Requirements
- Pure leaves + provider-facing modules: co-located `.test.mjs`, run by `npm test`.
- DOM-layer tests live in `ui/*.test.mjs` (jsdom).
- `test/setup.mjs` stubs `globalThis.__ZAI_API_KEY__` so generation/prompt modules load.

### Common Patterns
- Pure-leaf modules return plain data; side effects pushed up to callers.
- `/* global ... */` kept current per file (unused globals trip lint).
- Async wrappers over Office callback APIs return Promises; failures throw, callers guard.

## Dependencies

### Internal
- `ui/*` consumes `generation.js`, `storage.js`, `mailbox.js`, all leaves, `shared/zai.js`.
- `commands/` reuses `prompts.js`, `language.js`, `prompt-templates.js`, `shared/zai.js`.

### External
- `office.js` (host), `marked` + `dompurify` (via `ui/dom.js`).

<!-- MANUAL: -->
