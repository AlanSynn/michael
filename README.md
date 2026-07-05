<div align="center">
  <img src="assets/meet-michael-black.png" alt="Meet Michael Logo" width="300"/>
</div>

Meet Michael, your AI sidekick for Outlook. Michael helps you summarize, translate, and act on emails directly inside Outlook using **Z.AI GLM Coding Plan**.

## Features

- **Summarize:** Generate a concise email summary with action items and open questions.
- **Translate:** Translate email content into your chosen language.
- **Translate & Summarize:** Get both a translation and a summary in one step.
- **Reply Drafting:** Generate a reply draft from the current email.
- **Calendar Event Creation:** Extract event details from email content and create calendar entries.
- **Customizable Settings** (redesigned, theme-harmonized panel):
  - **Account** — enter the Z.AI API key; it is saved in Outlook add-in settings (`Office.context.roamingSettings`).
  - **Models** — choose primary + reply models; refresh the live Z.AI catalog with cached/fallback behavior.
  - **Language & Output** — default language, event-title language, TL;DR toggle, show-reply toggle, result font size.
  - **Appearance & Behavior** — theme, auto-run toggle + auto-run action.
  - **Prompt Templates** — grouped into collapsible Email / Command & System / Calendar accordions; save, clear, copy, or export them.
  - **Developer** — dev mode + dev server URL.
- **Theme Support:** Light, dark, or system theme.

## Install into Outlook for Mac

> **2025 note:** recent Outlook for Mac builds removed the *Get Add-ins → Upload* UI path. Manifests must be placed in Outlook's `wef` documents folder, then Outlook restarted. The scripts below automate this.
> Docs: [Sideload Office Add-ins on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).

### Production (recommended)

The deployed add-in lives at **`https://alansynn.com/michael/`**. Install its manifest:

```bash
npm run sideload:prod
```

This downloads `https://alansynn.com/michael/manifest.prod.xml` and copies it into Outlook's `wef` folder.

### Development (localhost)

```bash
npm install
npm start          # dev server on https://localhost:3000
npm run sideload:mac   # installs the local manifest.xml (localhost)
```

After either command: **fully quit Outlook (Cmd+Q), reopen it, open an email, and click the Michael button.** The taskpane must open *embedded* — if a browser tab opens with a sideloading warning, the manifest is not installed; re-run the sideload script.

## Setup for Development

1. **Clone the repository**
   ```bash
   git clone https://github.com/AlanSynn/michael.git
   cd michael
   ```
2. **Install dependencies**
   ```bash
   npm install
   ```
3. **Validate / test / build**
   ```bash
   npm run validate   # manifest schema check
   npm test           # unit tests (node:test, no extra deps)
   npm run build      # production bundle into ./dist
   npm run build:dev  # development bundle (localhost URLs)
   ```
4. **Run the dev server + sideload**
   ```bash
   npm start
   npm run sideload:mac
   ```

## Z.AI / GLM Coding Plan Configuration

### Provider values

- **Environment variable name:** `ZAI_API_KEY`
- **Coding endpoint:** `https://api.z.ai/api/coding/paas/v4`
- **General endpoint:** `https://api.z.ai/api/paas/v4`
- **Recommended primary model:** `glm-4.5-air`
- **Included fallback models:** `glm-4.5-air`, `glm-4.5-flash`, `glm-4.7`, `glm-5.1`, `glm-5-turbo`, `glm-5`, `glm-4.7-flash`, `glm-4.7-flashx`, `glm-4.6`, `glm-4.5`, `glm-4.5-x`, `glm-4.5-airx`, `glm-4-32b-0414-128k`

### Obtain credentials

1. Create a Z.AI API key from the Z.AI Open Platform / API Keys page.
2. Open Michael inside Outlook.
3. Enter the API key in **Settings → General** (Account section).
4. Save settings. The key is stored in Outlook add-in settings (`Office.context.roamingSettings`).

### Endpoint and model guidance

- Use the **coding endpoint** for GLM Coding Plan-compatible coding integrations.
- Use `glm-4.5-air` as the default for the Outlook taskpane based on current latency benchmarks.
- Use `glm-4.5-flash` when you want a short-response-first option.
- Use `glm-5.1` when you want a heavier flagship option.
- Michael reads the API key and prompt/model settings from Outlook add-in saved settings.

### Dropdown refresh behavior

- Opening Settings refreshes the dropdowns and template fields from saved Outlook settings.
- The model dropdown auto-refreshes from `https://api.z.ai/api/coding/paas/v4/models` when an API key has been saved.
- The API key and prompt fields stay blank until the user saves values.
- If live discovery fails, Michael falls back to a cached or baked safe model list that includes `glm-5-turbo`.
- Clicking **Refresh models** retries live discovery manually.
- The saved Outlook settings are also used by the ribbon quick-translate command.
- **Load Saved Defaults** restores the prompt defaults saved in Outlook add-in settings.
- **Save as Defaults** stores the current prompt set as the saved default prompt set.
- **Load Built-in Defaults** fills the form with the built-in prompt templates without forcing them globally.
- **Clear Templates** empties all prompt fields from saved Outlook settings.
- **Reset All Settings** clears saved Outlook settings and restores a blank form state.

## Usage

1. Select an email in Outlook.
2. Open the Michael taskpane from the ribbon.
3. Use **Summarize**, **Translate**, **Translate & Summarize**, or **Reply**.
4. If the email looks like an event invitation, use **Create Calendar Event**.
5. Open Settings to update the API key, model selection, or individual prompt templates.

## Technology Stack

- Office Add-ins Platform (Outlook Mail, XML manifest, VersionOverrides V1_0 + V1_1)
- JavaScript (ES modules) + `office.js`
- HTML5 / CSS3 (theme-token CSS variables, no UI framework)
- Webpack 5 + Babel
- Node.js / npm; unit tests via the built-in `node:test` runner, e2e via Playwright
- Z.AI coding-plan chat-completions integration target

## Project Structure

```
src/
  taskpane/
    index.js              Office.onReady + listener wiring (thin entry)
    ui/                   flows, calendar, settings-view, dom (presentation)
    ui/dom.js             DOM helpers + getApiKey + the renderMarkdown sanitizer
    generation.js         Z.AI orchestration over the shared client (DOM-free)
    storage.js            single Office.roamingSettings boundary
    mailbox.js            thin async mailbox wrappers + flat event handlers
    settings.js, prompts.js, theme.js, language.js, model-catalog.js   pure leaves (unit-tested)
    prompt-templates.js   default settings + prompt templates
  shared/zai.js           single Z.AI provider (chat + model discovery)
  commands/commands.js    Quick Translate ribbon command (function-file)
```

Dependency direction is one-way (entry → ui → generation → storage/mailbox → pure leaves → shared/zai); no cycles.

## Testing

```bash
npm test           # unit tests (node:test) + jsdom DOM-interaction tests
npm run build      # Playwright e2e serves the built ./dist
npm run test:e2e   # Playwright smoke + office.js-mocked integration
```

Unit tests cover the pure leaves (`theme`, `language`, `prompts`, `settings`,
`model-catalog`, `prompt-templates`), `shared/zai.js` (`fetch`-mocked `/models`
discovery and `/chat/completions` generation), and `ui/dom.js` (jsdom, against
the real `taskpane.html`). The Playwright suite gates on console errors / CSP
violations and verifies bootstrap wiring with a stubbed Office host.

CI runs `validate → lint → test → build → test:e2e` before deploying. See
[`CONTRIBUTING.md`](CONTRIBUTING.md) for the full `npm` script reference.

## Reference Docs

- Z.AI API introduction: <https://docs.z.ai/api-reference/introduction>
- Z.AI Quick Start: <https://docs.z.ai/guides/overview/quick-start>
- GLM-5-Turbo overview: <https://docs.z.ai/guides/llm/glm-5-turbo>
- Chat Completion model list: <https://docs.z.ai/api-reference/llm/chat-completion>
- GLM-4.5 / GLM-4.5-Air overview: <https://docs.z.ai/guides/llm/glm-4.5>

## Security

- **API key storage** — the Z.AI key is stored only in Outlook add-in settings
  (`Office.context.roamingSettings`), never in browser `localStorage` or logs.
  It is sent exclusively to `https://api.z.ai`.
- **Content sanitization** — email/LLM markdown reaches the DOM only through
  `renderMarkdown` (`marked` + `DOMPurify`), which is bundled (no CDN/SRI) and
  fail-closed. `escapeHtml` is used for plain-text concatenation.
- **Content-Security-Policy** — the taskpane ships a restrictive CSP
  (`connect-src 'self' https://api.z.ai`; no `unsafe-inline` scripts).

See [`CONTRIBUTING.md`](CONTRIBUTING.md) → Security review checklist. Report
vulnerabilities privately to the maintainer rather than opening a public issue.

## License

MIT License — see [LICENSE](LICENSE).
