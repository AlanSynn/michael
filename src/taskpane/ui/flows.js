/* global document, console, AbortController */

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

// --- request cancellation ---------------------------------------------------
// Only one generation flow runs at a time. Starting a new flow — a button
// click, or an autorun fired by switching emails — aborts the previous
// in-flight request. This prevents a late result from overwriting the current
// UI and releases the superseded flow's closures promptly. Under rapid email
// switching (autorun fires on every ItemChanged) unbounded in-flight requests
// were the single biggest source of memory growth + stale results.
let activeFlowController = null;

/** Abort any in-flight flow and return a fresh signal for the new one. */
function beginFlow() {
  if (activeFlowController) {
    activeFlowController.abort();
  }
  activeFlowController = new AbortController();
  return activeFlowController.signal;
}

/** True when the error is an AbortError from a superseded (cancelled) flow. */
function isCancelError(error) {
  return Boolean(error) && error.name === "AbortError";
}

/**
 * True when `signal` still belongs to the active flow. A superseded flow must
 * not touch the loading spinner or results — the newer flow owns the UI now,
 * so we bail before every DOM mutation.
 */
function isActiveFlow(signal) {
  return activeFlowController !== null && activeFlowController.signal === signal;
}

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

  const signal = beginFlow();
  showLoading("Summarizing email...");

  try {
    const emailContent = await getEmailContent();
    const subject = getSubject();
    const prompt = fillTemplate(requireTemplate("summarize", "Summarize"), {
      subject,
      content: emailContent,
    });

    if (getTldrModeSetting()) {
      const tldrContent = await generateTldrContent(
        prompt,
        apiKey,
        getLanguageText(getLanguage()),
        null,
        signal
      );
      if (!isActiveFlow(signal)) return;
      hideLoading();
      showResults(tldrContent, TYPES.SUMMARIZE);

      const fullContent = await generateContent(prompt, apiKey, null, false, signal);
      if (!isActiveFlow(signal)) return;
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      const summary = await generateContent(prompt, apiKey, null, false, signal);
      if (!isActiveFlow(signal)) return;
      showResults(summary, TYPES.SUMMARIZE);
    }
  } catch (error) {
    if (isCancelError(error)) return;
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    if (isActiveFlow(signal)) hideLoading();
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

  const signal = beginFlow();
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
      const tldrContent = await generateTldrContent(
        prompt,
        apiKey,
        getLanguageText(language),
        null,
        signal
      );
      if (!isActiveFlow(signal)) return;
      hideLoading();
      showResults(tldrContent, TYPES.TRANSLATE);

      const fullContent = await generateContent(prompt, apiKey, null, false, signal);
      if (!isActiveFlow(signal)) return;
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      const translation = await generateContent(prompt, apiKey, null, false, signal);
      if (!isActiveFlow(signal)) return;
      showResults(translation, TYPES.TRANSLATE);
    }
  } catch (error) {
    if (isCancelError(error)) return;
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    if (isActiveFlow(signal)) hideLoading();
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

  const signal = beginFlow();
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
      const tldrContent = await generateTldrContent(prompt, apiKey, undefined, null, signal);
      if (!isActiveFlow(signal)) return;
      hideLoading();
      showResults(tldrContent, TYPES.TRANSLATE_SUMMARIZE);

      const fullContent = await generateContent(prompt, apiKey, null, false, signal);
      if (!isActiveFlow(signal)) return;
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      const result = await generateContent(prompt, apiKey, null, false, signal);
      if (!isActiveFlow(signal)) return;
      showResults(result, TYPES.TRANSLATE_SUMMARIZE);
    }
  } catch (error) {
    if (isCancelError(error)) return;
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    if (isActiveFlow(signal)) hideLoading();
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

  const signal = beginFlow();
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

    const result = await generateContent(prompt, apiKey, replyModelOverride, false, signal);
    if (!isActiveFlow(signal)) return;
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
    if (isCancelError(error)) return;
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    if (isActiveFlow(signal)) hideLoading();
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
