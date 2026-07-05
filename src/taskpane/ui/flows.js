/* global document, console */

// Email action flows: summarize, translate, translate+summarize, reply, plus
// runAutorun() which dispatches the configured auto-action. Each flow builds
// its prompt with prompts.fillTemplate and renders via ui/dom.

import {
  generateContent,
  generateTldrContent,
  getLanguage,
  requireTemplate,
  requireModel,
} from "../generation.js";
import { getEmailContent, getSubject } from "../mailbox.js";
import { getLanguageText } from "../language.js";
import { fillTemplate } from "../prompts.js";
import { getAutorunAction } from "../settings.js";
import { getSettings } from "../storage.js";
import {
  TYPES,
  showNotification,
  showLoading,
  hideLoading,
  showResults,
  updateResults,
  updateExpandButton,
  formatReplyOutput,
  copyReply,
  getTldrModeSetting,
  getApiKey,
} from "./dom.js";
import { toggleSettingsView, getMissingApiKeyMessage } from "./settings-view.js";

/** Run the configured auto-action for the current item, if any. */
export function runAutorun() {
  const action = getAutorunAction(getSettingsSafe());
  switch (action) {
    case "summarize":
      summarizeEmail();
      return;
    case "translate":
      translateEmail();
      return;
    case "translateAndSummarize":
      translateAndSummarizeEmail();
      return;
    case "reply":
      generateReply();
      return;
    default:
      return;
  }
}

// --- Summarize --------------------------------------------------------------

export async function summarizeEmail() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView();
    return;
  }

  showLoading("Summarizing email...");

  try {
    const emailContent = await getEmailContent();
    const subject = getSubject();
    const prompt = fillTemplate(requireTemplate("summarize", "Summarize"), {
      subject,
      content: emailContent,
    });

    if (getTldrModeSetting()) {
      const tldrContent = await generateTldrContent(prompt, apiKey, getLanguageText(getLanguage()));
      hideLoading();
      showResults(tldrContent, TYPES.SUMMARIZE);

      const fullContent = await generateContent(prompt, apiKey);
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      const summary = await generateContent(prompt, apiKey);
      showResults(summary, TYPES.SUMMARIZE);
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    hideLoading();
  }
}

// --- Translate --------------------------------------------------------------

export async function translateEmail() {
  const apiKey = getApiKey();
  const language = getLanguage();
  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView();
    return;
  }

  showLoading(`Translating to ${getLanguageText(language)}...`);

  try {
    const emailContent = await getEmailContent();
    const subject = getSubject();
    const prompt = fillTemplate(requireTemplate("translate", "Translate"), {
      subject,
      content: emailContent,
      language: getLanguageText(language),
    });

    if (getTldrModeSetting()) {
      const tldrContent = await generateTldrContent(prompt, apiKey, getLanguageText(language));
      hideLoading();
      showResults(tldrContent, TYPES.TRANSLATE);

      const fullContent = await generateContent(prompt, apiKey);
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      const translation = await generateContent(prompt, apiKey);
      showResults(translation, TYPES.TRANSLATE);
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    hideLoading();
  }
}

// --- Translate + Summarize --------------------------------------------------

export async function translateAndSummarizeEmail() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView();
    return;
  }

  showLoading("Translating and summarizing...");

  try {
    const emailContent = await getEmailContent();
    const subject = getSubject();
    const language = getLanguageText(getLanguage());

    const prompt = fillTemplate(requireTemplate("translateSummarize", "Translate & Summarize"), {
      subject,
      content: emailContent,
      language,
    });

    if (getTldrModeSetting()) {
      const tldrContent = await generateTldrContent(prompt, apiKey);
      hideLoading();
      showResults(tldrContent, TYPES.TRANSLATE_SUMMARIZE);

      const fullContent = await generateContent(prompt, apiKey);
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      const result = await generateContent(prompt, apiKey);
      showResults(result, TYPES.TRANSLATE_SUMMARIZE);
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    hideLoading();
  }
}

// --- Reply ------------------------------------------------------------------

export async function generateReply() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView();
    return;
  }

  showLoading("Generating reply...");

  try {
    const emailContent = await getEmailContent();
    const subject = getSubject();

    const prompt = fillTemplate(requireTemplate("reply", "Reply"), {
      subject,
      content: emailContent,
      language: getLanguageText(getLanguage()),
    });

    let replyModelOverride = null;
    try {
      replyModelOverride = requireModel("replyModel", "Reply model");
    } catch (error) {
      console.error("Error getting reply model:", error);
    }

    const result = await generateContent(prompt, apiKey, replyModelOverride);
    const formattedReply = formatReplyOutput(result);

    const tldrContent = document.getElementById("tldr-content");
    const resultContent = document.getElementById("result-content");
    if (tldrContent) {
      tldrContent.innerHTML = formattedReply.html;
    }
    if (resultContent) {
      resultContent.innerHTML = formattedReply.html;
    }

    const resultSection = document.getElementById("result-section");
    if (resultSection) {
      resultSection.style.display = "block";
    }

    const copyReplyButton = document.getElementById("copy-reply");
    const copyResultButton = document.getElementById("copy-result");
    if (copyReplyButton) {
      copyReplyButton.style.display = "inline-block";
    }
    if (copyResultButton) {
      copyResultButton.style.display = "none";
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    hideLoading();
  }
}

export { copyReply };

// --- helpers ----------------------------------------------------------------

function getSettingsSafe() {
  try {
    return getSettings();
  } catch (error) {
    console.error("Error getting settings for autorun:", error);
    return {};
  }
}
