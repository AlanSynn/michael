/* global console */

// Z.AI orchestration over the shared zai.js client + storage + mailbox.
// Pure generation logic — no DOM. UI status updates (loading spinners, API-key
// field reads) stay in the UI layer (dom.js / flows.js / settings-view.js).

import { fetchAvailableModels, generateText } from "../shared/zai.js";
import { DEFAULT_SETTINGS } from "./prompt-templates.js";
import { getSettings } from "./storage.js";

export { fetchAvailableModels };

// Generation tuning. TL;DR is capped tight so the quick summary renders first;
// full generations get the larger budget. Temperature is low for deterministic
// translation/summary output.
const TLDR_MAX_TOKENS = 800;
const FULL_MAX_TOKENS = 8192;
const DEFAULT_TEMPERATURE = 0.4;

/**
 * Generate text via Z.AI. Callers own the loading-spinner lifecycle (flows.js
 * and calendar.js already hide it in their own finally blocks). Pass an
 * AbortSignal to make the request cancellable (see flows.js beginFlow).
 */
export async function generateContent(
  prompt,
  apiKey,
  modelOverride = null,
  isTldr = false,
  signal = null
) {
  let model = "";

  if (modelOverride) {
    model = modelOverride;
  } else {
    try {
      model = requireModel("model", "Primary model");
    } catch (error) {
      console.error("Error getting model from settings:", error);
      throw error;
    }
  }

  try {
    return await generateText(prompt, {
      apiKey,
      model,
      maxTokens: isTldr ? TLDR_MAX_TOKENS : FULL_MAX_TOKENS,
      temperature: DEFAULT_TEMPERATURE,
      signal: signal || undefined,
    });
  } catch (error) {
    console.error("Error generating content:", error);
    throw error;
  }
}

/**
 * Generate the TL;DR for an email. The caller already fetched the body + subject
 * for the main flow, so accept them here rather than re-calling body.getAsync
 * (which doubled the async traffic on every TL;DR flow and raced against a
 * mid-flight email switch).
 *
 * @param {object} opts
 * @param {string} opts.apiKey
 * @param {string} opts.emailContent
 * @param {string} opts.subject
 * @param {string} [opts.language]
 * @param {string|null} [opts.modelOverride]
 * @param {AbortSignal|null} [opts.signal]
 */
export async function generateTldrContent({
  apiKey,
  emailContent,
  subject,
  language = "Korean",
  modelOverride = null,
  signal = null,
}) {
  const tldrPrompt = requireTemplate("tldrPrompt", "TL;DR")
    .replace("{subject}", subject)
    .replace("{content}", emailContent)
    .replace("{language}", language);

  return generateContent(tldrPrompt, apiKey, modelOverride, true, signal);
}

export function getLanguage() {
  const settings = getSettings();
  return settings.defaultLanguage || DEFAULT_SETTINGS.defaultLanguage;
}

export function getTemplateValue(templateKey) {
  const settings = getSettings();
  return settings.templates && typeof settings.templates[templateKey] === "string"
    ? settings.templates[templateKey].trim()
    : "";
}

export function requireTemplate(templateKey, label) {
  const template = getTemplateValue(templateKey);
  if (!template) {
    throw new Error(`${label} prompt is empty. Configure it in Settings > Templates.`);
  }
  return template;
}

export function requireModel(settingKey, label) {
  const settings = getSettings();
  const model = typeof settings[settingKey] === "string" ? settings[settingKey].trim() : "";
  if (!model) {
    throw new Error(`${label} is empty. Select a model in Settings > General.`);
  }
  return model;
}
