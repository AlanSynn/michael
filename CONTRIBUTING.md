# Contributing to Michael

Thanks for your interest in improving Michael. This is a small, focused project;
the conventions below keep it maintainable.

## Setup

```bash
git clone https://github.com/AlanSynn/michael.git
cd michael
npm install
```

Requirements: Node.js 20+ (22 in CI), npm.

## Day-to-day commands

| Command | Purpose |
| --- | --- |
| `npm run validate` | Validate the XML manifest against the Office schema. |
| `npm run lint` / `npm run lint:fix` | ESLint (office-addin config) check / auto-fix. |
| `npm test` | Unit tests (`node:test`) — pure leaves, `shared/zai.js`, jsdom DOM tests. |
| `npm run test:e2e` | Playwright e2e (builds nothing itself — run `npm run build` first). |
| `npm run build` / `build:dev` | Production / development webpack bundle into `./dist`. |
| `npm run dev-server` | Local HTTPS dev server on `https://localhost:3000`. |
| `npm run sideload:mac` / `sideload:prod` | Install the dev / prod manifest into Outlook for Mac. |

Before opening a PR, run:

```bash
npm run validate && npm run lint && npm test && npm run build && npm run test:e2e
```

(This mirrors the CI `verify` job.)

## Architecture

Strict one-way dependency direction, no cycles:

```
index.js  →  ui/*  →  generation.js  →  storage.js / mailbox.js  →  pure leaves  →  shared/zai.js
```

- `storage.js` is the **only** module that touches `Office.context.roamingSettings`.
- `mailbox.js` is the **only** module that touches `Office.context.mailbox.item`
  (the function-file `commands.js` imports these facades rather than reaching in).
- `generation.js` is DOM-free. DOM reads/writes (including `getApiKey`) live in
  `ui/dom.js`; loading-spinner lifecycle is owned by the flow that triggers it.
- `renderMarkdown` in `ui/dom.js` is the single sanitizing trust boundary for
  rendering untrusted (LLM/email) markdown.

When adding a feature, place it at the lowest layer that has its dependencies.
Don't introduce a second implementation of something a facade already owns.

## Commit + pull-request style

- Conventional Commits prefixes (`feat:`, `fix:`, `test:`, `refactor:`, `docs:`,
  `chore:`, `ci:`).
- Keep commits atomic and focused; one logical change per commit.
- PRs should reference the issue they close (if any) and summarize behavior
  changes. Screenshots help for UI changes.

## Sideload smoke test

Outlook cannot be driven headlessly, so the real proof of an operability change
is a manual sideload: `npm run sideload:mac`, fully quit Outlook (Cmd+Q),
reopen, open an email, click the Michael button → the taskpane must open
**embedded** (not a browser tab). Then exercise Summarize / Translate / T&S /
Reply / Settings save+reload / model Refresh / the Quick Translate ribbon command.

## Security review checklist

Before merging anything that touches input handling, settings, or the network:

- [ ] Untrusted text (email body, LLM output) reaches the DOM **only** through
      `renderMarkdown` (DOMPurify-sanitized).
- [ ] No `innerHTML` assignment of unsanitized content; use `escapeHtml` for
      plain-text contexts.
- [ ] The API key is read only via `getApiKey()` and sent only to
      `https://api.z.ai`; never logged, never persisted outside
      `Office.context.roamingSettings`.
- [ ] No new outbound hosts without a matching CSP `connect-src` entry.
- [ ] No secrets, tokens, or keys committed. The build injects `ZAI_API_KEY`
      only via `DefinePlugin` for local dev; production reads it from saved
      Outlook add-in settings.

## Reporting a security issue

Please do **not** open a public issue for security vulnerabilities. Email the
maintainer privately so a fix can be coordinated before disclosure.
