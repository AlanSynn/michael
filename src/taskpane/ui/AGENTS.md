<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# ui

## Purpose
DOM controllers — each owns one interaction surface. They mutate the DOM, drive
`generation.js`/`mailbox.js`, and (in the case of flows + calendar) join the shared
single-flight flow controller so only one generation owns the result section at a time.

## Key Files
| File | Description |
|------|-------------|
| `flows.js` | summarize / translate / T&S / generate-reply. Each flow: `beginFlow()` → `isActiveFlow` guards on every DOM write incl. all catch paths. Reply stamps `data-result-type="reply"` + `data-reply-raw`. |
| `calendar.js` | Detect event (`checkIfCalendarEvent`, opt-in auto-detect), parse JSON (`parseEventDetailsWithZai`), `displayNewAppointmentForm`. Joins shared flow control. `isValidDateTime` accepts Z/offset/fractional seconds. Own probe controller for button-state auto-detect. |
| `settings-view.js` | Settings panel controller: open/close, save (snapshot + rollback on error), focus return, Esc-to-close. `getMissingApiKeyMessage`. |
| `dom.js` | `$`, notifications (with timer cleanup), loading, results render, clipboard (attr-driven), `renderMarkdown` (DOMPurify + marked — single sanitization boundary), `stampResultType`, `escapeHtml`. |
| `flow-control.js` | Shared single-flight: `beginFlow()`, `isActiveFlow(signal)`, `isCancelError(error)`. Owns the ONE AbortController shared by flows + calendar. |

## For AI Agents

### Working In This Directory
- **Every async flow must start with `beginFlow()` and guard every DOM mutation
  (success, error, AND finally) with `isActiveFlow(signal)`.** Catches should early-return
  on `isCancelError(error) || !isActiveFlow(signal)`. Without this, a superseded flow
  clobbers the newer flow's result/notifications.
- **Never own a separate AbortController for a flow.** Use the shared one in `flow-control.js`.
  (`calendar.js` keeps a private controller ONLY for the button-state probe, which is not a
  result-section flow — leave that as-is.)
- **Copy logic is attr-driven, not text-inferred.** `stampResultType(type, replyRaw?)`
  writes `data-result-type` (+ optional `data-reply-raw`) on `#result-section`; `copyResult` /
  `copyReply` read them. Do not reintroduce subject-substring matching — it false-positived
  on emails whose subject contained "Re:".
- **`renderMarkdown()` is the only path email/LLM text reaches the DOM.** It runs marked →
  DOMPurify with `ALLOWED_URI_REGEXP` + an `afterSanitizeAttributes` hook that forces
  `target=_blank rel=noopener noreferrer` on links. Both libs are bundled (no CDN fail-open).
- **Calendar appointment body is plain text** (`displayNewAppointmentForm`), so `prompt-templates.js`
  `calendarParse` uses `contentType: "Text"` — do not switch it back to `"HTML"`.
- **`/* global */`** in `calendar.js` keeps `AbortController` (probe uses it) — don't drop it.

### Testing Requirements
- `dom-interaction.test.mjs`, `flow-control.test.mjs` cover DOM + flow semantics (jsdom + clipboard mock).
- Playwright e2e (`e2e/`) smoke-tests the built taskpane; run `npm run build` first.
- When touching copy/clipboard, test `stampResultType` + `copyResult`/`copyReply` together.

### Common Patterns
- `beginFlow()` at flow entry; `isActiveFlow(signal)` gates post-await writes.
- `stampResultType` before rendering results; hide `#copy-reply` for non-reply types.
- Notifications clear their pending timer on replace (no overlap).

## Dependencies

### Internal
- `flows.js`, `calendar.js` → `../generation.js`, `../mailbox.js`, `../storage.js`, `../prompts.js`,
  `../language.js`, `../date-context.js`, `../prompt-templates.js`, `./dom.js`, `./flow-control.js`.
- `settings-view.js` → `../storage.js`, `./dom.js`.
- `dom.js` → `marked`, `dompurify` (bundled).

### External
- `office.js` (`calendar.js` uses `Office.context.mailbox.displayNewAppointmentForm`).

<!-- MANUAL: -->
