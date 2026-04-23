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
- **Customizable Settings:**
  - Enter the API key directly in the taskpane UI and save it in Outlook add-in settings.
  - Choose the primary and reply models used for AI flows.
  - Refresh the provider model catalog from Z.AI, with cached/fallback behavior if live discovery fails.
  - Set default translation language and event-title language.
  - Choose theme, font size, TL;DR mode, reply visibility, and autorun behavior.
  - Configure each prompt individually: summarize, translate, translate+summarize, reply, quick translate command, TL;DR, calendar parse, and calendar check.
  - Save prompt defaults in Outlook add-in settings, clear them, or load built-in defaults.
- **Theme Support:** Light, dark, or system theme.

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
3. **Start the development server**
   ```bash
   npm start
   ```
   This starts the local add-in web server, typically at `https://localhost:3000`.
4. **Sideload the add-in in Outlook**
   Use `manifest.xml` from the repo root.
   - Microsoft Docs: [Sideload Outlook add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)

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
3. Enter the API key in **Settings → General**.
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

- Office Add-ins Platform
- JavaScript (ES6+)
- HTML5 / CSS3
- Webpack
- Node.js / npm
- Z.AI coding-plan chat-completions integration target

## Reference Docs

- Z.AI API introduction: <https://docs.z.ai/api-reference/introduction>
- Z.AI Quick Start: <https://docs.z.ai/guides/overview/quick-start>
- GLM-5-Turbo overview: <https://docs.z.ai/guides/llm/glm-5-turbo>
- Chat Completion model list: <https://docs.z.ai/api-reference/llm/chat-completion>
- GLM-4.5 / GLM-4.5-Air overview: <https://docs.z.ai/guides/llm/glm-4.5>

## License

MIT License
