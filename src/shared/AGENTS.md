<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# shared

## Purpose
The single Z.AI GLM provider client, imported by both webpack entries (taskpane +
commands). All `fetch` calls against `api.z.ai` live here — never duplicate the client.

## Key Files
| File | Description |
|------|-------------|
| `zai.js` | Chat-completion + `/models` client. Reads `__ZAI_API_KEY__` (dev-gated by webpack DefinePlugin → `""` in prod) + `__ZAI_CODING_BASE_URL__` (dev-only; prod uses canonical `https://api.z.ai/api/coding/paas/v4`). Sends the key only as `Authorization: Bearer <key>`. Exports pure helpers: `extractModelNames`, `normalizeTextContent`, `dedupeModels`, `getDefaultZaiModels`. |
| `zai.test.mjs` | Realistic `/models` payloads + `globalThis.fetch` mock covering the helpers + client paths. |

## For AI Agents

### Working In This Directory
- **Key handling is the security boundary.** The runtime key (user-entered, in
  `roamingSettings`) is passed as a function argument; the build-time `__ZAI_API_KEY__`
  is dev-only. Never log the key, never place it in a URL/query, never persist to localStorage.
- **Prod endpoint is hardcoded canonical** (`https://api.z.ai/api/coding/paas/v4`);
  `__ZAI_CODING_BASE_URL__` only overrides it in dev builds. Do not wire a prod base-URL escape hatch.
- **Promote pure helpers, keep `fetch` here.** `extractModelNames`, `normalizeTextContent`,
  `dedupeModels`, `getDefaultZaiModels` are pure + exported + unit-tested. New helpers follow suit.
- **`AbortSignal` is threaded through** every request so the shared single-flight controller
  (`taskpane/ui/flow-control.js`) can cancel in-flight provider calls. Keep the `signal` parameter.

### Testing Requirements
- `zai.test.mjs` mocks `globalThis.fetch`; add cases for any new endpoint or helper.
- Pure helpers are safe to test in plain Node (no Office/DOM).

### Common Patterns
- Each public function takes `apiKey` as an explicit arg (not read from a global at call time),
  keeping the key flow auditable.
- `dedupeModels` + cap applied at the catalog boundary (`taskpane/model-catalog.js`), not here.

## Dependencies

### Internal
- Consumed by `taskpane/generation.js` and `commands/commands.js`.

### External
- Browser `fetch` (the only `api.z.ai` call site). CSP `connect-src` scopes it.

<!-- MANUAL: -->
