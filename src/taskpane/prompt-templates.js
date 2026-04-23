/**
 * Z.AI GLM Coding Plan defaults for Michael taskpane integration.
 * Import these constants from taskpane.js when the provider migration is wired.
 */

/**
 * @typedef {Object} ZaiProviderConfig
 * @property {string} providerName
 * @property {string} planName
 * @property {string} apiKeyEnvVar
 * @property {string} codingBaseUrl
 * @property {string} generalBaseUrl
 * @property {string} defaultModel
 * @property {string} replyModel
 * @property {readonly string[]} modelSuggestions
 */

/**
 * @typedef {Object} PromptTemplates
 * @property {string} summarize
 * @property {string} translate
 * @property {string} translateSummarize
 * @property {string} reply
 * @property {string} commandTranslate
 * @property {string} tldrPrompt
 * @property {string} calendarParse
 * @property {string} calendarCheck
 */

/**
 * @typedef {Object} SettingsDefaults
 * @property {string} model
 * @property {string} replyModel
 * @property {string} defaultLanguage
 * @property {string} eventTitleLanguage
 * @property {string} theme
 * @property {string} fontSize
 * @property {string} tldrMode
 * @property {string} showReply
 * @property {string} autorun
 * @property {string} autorunOption
 * @property {string} devMode
 * @property {string} devServer
 */

/** @type {readonly string[]} */
export const TEMPLATE_KEYS = Object.freeze([
  "summarize",
  "translate",
  "translateSummarize",
  "reply",
  "commandTranslate",
  "tldrPrompt",
  "calendarParse",
  "calendarCheck",
]);

/** @type {ZaiProviderConfig} */
export const ZAI_PROVIDER_CONFIG = Object.freeze({
  providerName: "Z.AI",
  planName: "GLM Coding Plan",
  apiKeyEnvVar: "ZAI_API_KEY",
  codingBaseUrl: "https://api.z.ai/api/coding/paas/v4",
  generalBaseUrl: "https://api.z.ai/api/paas/v4",
  defaultModel: "glm-4.5-air",
  replyModel: "glm-4.5-air",
  modelSuggestions: Object.freeze(["glm-4.5-air", "glm-4.5-flash", "glm-4.7"]),
});

export const PROVIDER_MESSAGES = Object.freeze({
  apiKeyLabel: "Z.AI API Key",
  modelLabel: "GLM Model",
  missingApiKey: "Please enter your Z.AI API key in Settings > General and save it.",
  exportTitle: "Michael Prompt Templates (Z.AI GLM Coding Plan)",
  templatesReset: "Templates reset to Z.AI defaults",
  templatesCleared: "Prompt templates cleared from Outlook add-in settings.",
  sessionDefaultsSaved: "Current prompt templates saved as Outlook defaults.",
});

/** @type {PromptTemplates} */
export const DEFAULT_PROMPT_TEMPLATES = Object.freeze({
  summarize: `You are Michael, an Outlook email assistant powered by Z.AI GLM models.

Review the email below and respond in Markdown with these sections only:
1. TL;DR — 1-3 bullet points with the most important takeaways.
2. Summary — a concise summary of the message.
3. Action Items — concrete asks, decisions, deadlines, or follow-ups.
4. Risks / Open Questions — blockers, ambiguities, or missing information.

Keep the answer factual, compact, and grounded in the email. If the email does not contain enough information for a section, say "None".

Subject: {subject}
Content:
{content}`,
  translate: `You are Michael, an Outlook email assistant powered by Z.AI GLM models.

Translate the email below into {language}. Preserve meaning, tone, names, dates, numbers, and formatting.

Return Markdown with these sections only:
1. TL;DR — a very short summary in {language}.
2. Translation — the full translated email.

Do not add commentary outside the requested sections.

Subject: {subject}
Content:
{content}`,
  translateSummarize: `You are Michael, an Outlook email assistant powered by Z.AI GLM models.

Translate the email below into {language} and summarize it.

Return Markdown with these sections only:
1. TL;DR — a very short summary in {language}.
2. Summary — the key points in {language}.
3. Translation — the full translated email in {language}.

Preserve meaning, tone, names, dates, numbers, and formatting. Do not add commentary outside the requested sections.

Subject: {subject}
Content:
{content}`,
  reply: `You are Michael, an Outlook email assistant powered by Z.AI GLM models.

Draft a clear, professional reply in {language} to the email below.

Output format:
Subject: <reply subject>

<body>

Requirements:
- Keep the reply actionable and courteous.
- Preserve any important dates, commitments, or questions from the source email.
- If the original email is missing context needed for a confident reply, say so briefly in the body.
- Do not include code fences or commentary outside the reply.

Subject: {subject}
Content:
{content}`,
  commandTranslate: `You are Michael, an Outlook email assistant powered by Z.AI GLM models.

Translate the email below into {language}.

Requirements:
- Return only the translated email body.
- Preserve meaning, tone, names, dates, numbers, and paragraph structure.
- Do not add a summary, title, bullets, or commentary.

Subject: {subject}
Content:
{content}`,
  tldrPrompt: `Provide a very concise TL;DR in {language} for the email below.

Rules:
- Focus on the main point, key actions, and deadlines.
- Keep it short and direct.
- Do not add introductions or commentary.

Subject: {subject}
Content:
{content}`,
  calendarParse: `Analyze the following email content and extract information needed to create a calendar event for Microsoft Graph API.
The response must be valid JSON and follow the required schema exactly.
Return only JSON with no extra commentary.

Event title language instructions:
{languageInstructions}

Required JSON format:
{
  "subject": "Meeting title",
  "body": {
    "contentType": "HTML",
    "content": "Meeting description"
  },
  "start": {
    "dateTime": "YYYY-MM-DDTHH:mm:ss",
    "timeZone": "Asia/Seoul"
  },
  "end": {
    "dateTime": "YYYY-MM-DDTHH:mm:ss",
    "timeZone": "Asia/Seoul"
  },
  "location": {
    "displayName": "Location name"
  },
  "attendees": [
    {
      "emailAddress": {
        "address": "attendee@email.com",
        "name": "Attendee Name"
      },
      "type": "required"
    }
  ],
  "isOnlineMeeting": true,
  "onlineMeetingProvider": "teamsForBusiness"
}

Important notes:
1. Convert dates and times to ISO 8601 format (YYYY-MM-DDTHH:mm:ss)
2. Email addresses must be valid
3. Use "Asia/Seoul" as the default timezone
4. Set isOnlineMeeting to true if Teams or video conference details are present
5. Mark unknown values as null

Email content:
{content}`,
  calendarCheck: `Check whether the following email content is a calendar event, meeting invite, appointment, or schedule-related notice.

Return "true" only if the content clearly includes one or more of these:
- Date and time information
- Meeting or appointment language
- Attendee information
- Location information
- Calendar or RSVP actions

Return "false" otherwise.

Email content:
{content}`,
});

/** @type {SettingsDefaults} */
export const DEFAULT_SETTINGS = Object.freeze({
  model: "",
  replyModel: "",
  defaultLanguage: "ko",
  eventTitleLanguage: "en",
  theme: "system",
  fontSize: "medium",
  tldrMode: "true",
  showReply: "true",
  autorun: "false",
  autorunOption: "summarize",
  devMode: "false",
  devServer: "",
});

export const BLANK_PROMPT_TEMPLATES = Object.freeze({
  summarize: "",
  translate: "",
  translateSummarize: "",
  reply: "",
  commandTranslate: "",
  tldrPrompt: "",
  calendarParse: "",
  calendarCheck: "",
});

/**
 * @returns {PromptTemplates}
 */
export function createDefaultPromptTemplates() {
  return {
    summarize: DEFAULT_PROMPT_TEMPLATES.summarize,
    translate: DEFAULT_PROMPT_TEMPLATES.translate,
    translateSummarize: DEFAULT_PROMPT_TEMPLATES.translateSummarize,
    reply: DEFAULT_PROMPT_TEMPLATES.reply,
    commandTranslate: DEFAULT_PROMPT_TEMPLATES.commandTranslate,
    tldrPrompt: DEFAULT_PROMPT_TEMPLATES.tldrPrompt,
    calendarParse: DEFAULT_PROMPT_TEMPLATES.calendarParse,
    calendarCheck: DEFAULT_PROMPT_TEMPLATES.calendarCheck,
  };
}

export function createBlankPromptTemplates() {
  return {
    summarize: BLANK_PROMPT_TEMPLATES.summarize,
    translate: BLANK_PROMPT_TEMPLATES.translate,
    translateSummarize: BLANK_PROMPT_TEMPLATES.translateSummarize,
    reply: BLANK_PROMPT_TEMPLATES.reply,
    commandTranslate: BLANK_PROMPT_TEMPLATES.commandTranslate,
    tldrPrompt: BLANK_PROMPT_TEMPLATES.tldrPrompt,
    calendarParse: BLANK_PROMPT_TEMPLATES.calendarParse,
    calendarCheck: BLANK_PROMPT_TEMPLATES.calendarCheck,
  };
}

/**
 * @returns {SettingsDefaults & { templates: PromptTemplates }}
 */
export function createDefaultSettings() {
  return {
    ...DEFAULT_SETTINGS,
    templates: createDefaultPromptTemplates(),
  };
}

export function createBlankSettings() {
  return {
    ...DEFAULT_SETTINGS,
    templates: createBlankPromptTemplates(),
  };
}
