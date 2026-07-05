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
import { getEmailContent, getSubject, hasSelectedItem } from "../mailbox.js";
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
  stampResultType,
  copyReply,
  getTldrModeSetting,
  getApiKey,
} from "./dom.js";
import { toggleSettingsView, getMissingApiKeyMessage } from "./settings-view.js";
import { beginFlow, isActiveFlow, isCancelError } from "./flow-control.js";

// Only one generation flow runs at a time. beginFlow() aborts the previous
// in-flight request; isActiveFlow(signal) guards every DOM mutation so a late
// result from a superseded flow cannot overwrite the current UI. The shared
// controller (flow-control.js) also covers the calendar parse, so an email
// flow and a calendar extraction cannot race on the result-section DOM.

/** Run the configured auto-action for the current item, if any. */
export function runAutorun() {
  // Stay quiet when nothing is selected — autorun fires on every ItemChanged,
  // including scrolling the mailbox list with no item body to read.
  if (!hasSelectedItem()) {
    return;
  }
  // Stay quiet when no API key is configured. The manual button flows open
  // Settings to guide the user, but doing that on every ItemChanged traps the
  // user in an oscillating panel + re-fires the missing-key toast.
  if (!getApiKey()) {
    return;
  }
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
      const tldrContent = await generateTldrContent({
        apiKey,
        emailContent,
        subject,
        language: getLanguageText(getLanguage()),
        signal,
      });
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
    // Bail on a cancel OR if this flow was superseded mid-flight: a late
    // non-Abort rejection from a cancelled flow must not toast over the
    // active flow's result. The isActiveFlow gate was already enforced before
    // every successful DOM mutation; this extends it to the error path.
    if (isCancelError(error) || !isActiveFlow(signal)) return;
    if (error && error.code === "NO_ITEM") {
      showNotification("Select an email first.", "info");
      return;
    }
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
      const tldrContent = await generateTldrContent({
        apiKey,
        emailContent,
        subject,
        language: getLanguageText(language),
        signal,
      });
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
    // Bail on a cancel OR if this flow was superseded mid-flight: a late
    // non-Abort rejection from a cancelled flow must not toast over the
    // active flow's result. The isActiveFlow gate was already enforced before
    // every successful DOM mutation; this extends it to the error path.
    if (isCancelError(error) || !isActiveFlow(signal)) return;
    if (error && error.code === "NO_ITEM") {
      showNotification("Select an email first.", "info");
      return;
    }
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
      const tldrContent = await generateTldrContent({
        apiKey,
        emailContent,
        subject,
        language,
        signal,
      });
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
    // Bail on a cancel OR if this flow was superseded mid-flight: a late
    // non-Abort rejection from a cancelled flow must not toast over the
    // active flow's result. The isActiveFlow gate was already enforced before
    // every successful DOM mutation; this extends it to the error path.
    if (isCancelError(error) || !isActiveFlow(signal)) return;
    if (error && error.code === "NO_ITEM") {
      showNotification("Select an email first.", "info");
      return;
    }
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

    // Stamp the result type + structured Subject/body so copyResult/copyReply
    // use the captured text instead of inferring "is this a reply?" from the
    // rendered markup (which false-positive'd on summaries quoting Subject:).
    stampResultType("reply", formattedReply.raw);

    const copyReplyButton = document.getElementById("copy-reply");
    const copyResultButton = document.getElementById("copy-result");
    if (copyReplyButton) {
      copyReplyButton.style.display = "inline-block";
    }
    if (copyResultButton) {
      copyResultButton.style.display = "none";
    }
  } catch (error) {
    // Bail on a cancel OR if this flow was superseded mid-flight: a late
    // non-Abort rejection from a cancelled flow must not toast over the
    // active flow's result. The isActiveFlow gate was already enforced before
    // every successful DOM mutation; this extends it to the error path.
    if (isCancelError(error) || !isActiveFlow(signal)) return;
    if (error && error.code === "NO_ITEM") {
      showNotification("Select an email first.", "info");
      return;
    }
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
