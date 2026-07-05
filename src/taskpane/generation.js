/* global console, document */

// Z.AI orchestration over the shared zai.js client + storage + mailbox.
// Pure generation logic; UI status updates stay in the settings-view/UI layer.

import { fetchAvailableModels, generateText } from "../shared/zai.js";
import { DEFAULT_SETTINGS } from "./prompt-templates.js";
import { getSettings } from "./storage.js";
import { getEmailContent, getSubject } from "./mailbox.js";

export { fetchAvailableModels };

/**
 * Generate text via Z.AI. Hides the loading spinner on completion for non-TL;DR
 * requests (preserves prior behavior; UI coupling to be lifted in a later pass).
 */
export async function generateContent(prompt, apiKey, modelOverride = null, isTldr = false) {
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
      maxTokens: isTldr ? 800 : 8192,
      temperature: 0.4,
    });
  } catch (error) {
    console.error("Error generating content:", error);
    throw error;
  } finally {
    if (!isTldr) {
      const loading = document.getElementById("loading");
      if (loading) {
        loading.style.display = "none";
      }
    }
  }
}

export async function generateTldrContent(
  prompt,
  apiKey,
  language = "Korean",
  modelOverride = null
) {
  const subject = getSubject();
  const emailContent = await getEmailContent();

  const tldrPrompt = requireTemplate("tldrPrompt", "TL;DR")
    .replace("{subject}", subject)
    .replace("{content}", emailContent)
    .replace("{language}", language);

  return generateContent(tldrPrompt, apiKey, modelOverride, true);
}

/** Read the configured Z.AI API key (input field first, then saved settings). */
export function getApiKey() {
  const input = document.getElementById("dropdown-api-key");
  if (input && typeof input.value === "string" && input.value.trim()) {
    return input.value.trim();
  }

  const settings = getSettings();
  return typeof settings.apiKey === "string" ? settings.apiKey.trim() : "";
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
