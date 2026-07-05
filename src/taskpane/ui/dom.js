/* global document, console, navigator, setTimeout */

// Low-level DOM helpers: notifications, loading/result sections, clipboard,
// asset paths, and the result-type enum shared with the flows. No Office, no
// generation. Reads settings (storage facade) only to render the result view.

import { marked } from "marked";
import DOMPurify from "dompurify";
import { getSettings } from "../storage.js";

/** Result kinds passed to showResults to toggle per-type buttons. */
export const TYPES = Object.freeze({
  SUMMARIZE: 0,
  TRANSLATE: 1,
  TRANSLATE_SUMMARIZE: 2,
  REPLY: 3,
});

/**
 * Read the configured Z.AI API key: the settings-panel input first (so an unsaved
 * typed key works), then saved Outlook add-in settings. Lives in the DOM layer
 * because it reads an input element; generation.js stays DOM-free.
 */
export function getApiKey() {
  const input = document.getElementById("dropdown-api-key");
  if (input && typeof input.value === "string" && input.value.trim()) {
    return input.value.trim();
  }
  const settings = getSettings();
  return typeof settings.apiKey === "string" ? settings.apiKey.trim() : "";
}

/** Resolve a packaged asset URL under ./assets (rewritten by webpack/html-loader). */
export function getAssetPath(fileName) {
  return `./assets/${fileName}`;
}

/**
 * Escape a string for safe insertion into an HTML text context.
 * Use this for any untrusted plain text (email body, LLM-parsed fields) that
 * is concatenated into markup.
 */
export function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Render untrusted markdown to sanitized HTML. LLM/email output reaches the DOM
 * only through this so script/event-handler payloads cannot execute (the taskpane
 * has access to roamingSettings, i.e. the API key). DOMPurify keeps safe tags.
 *
 * marked + DOMPurify are bundled (not CDN), so both are always available — no
 * fail-open path exists. This is the single trust boundary for rich rendering.
 */
export function renderMarkdown(content) {
  return DOMPurify.sanitize(marked.parse(content || ""));
}

/** Shorthand getElementById. Returns the element or null. */
export function $(id) {
  return document.getElementById(id);
}

/** Set the text content of an element if it exists. */
export function setText(id, message) {
  const el = $(id);
  if (el) {
    el.textContent = message;
  }
}

/**
 * Show a transient toast notification.
 * @param {string} message
 * @param {"info"|"success"|"warning"|"error"} [type]
 */
export function showNotification(message, type = "info") {
  const existing = $("notification");
  if (existing) {
    existing.remove();
  }

  const notification = document.createElement("div");
  notification.id = "notification";
  notification.className = `notification ${type}`;
  notification.textContent = message;

  document.body.appendChild(notification);
  // Force reflow so the slide-in animation plays.
  notification.offsetHeight;
  notification.style.animation = "slideInFromTop 0.3s ease-out";

  setTimeout(() => {
    notification.style.animation = "slideOutToTop 0.3s ease-out";
    setTimeout(() => notification.remove(), 300);
  }, 3000);
}

/** Show the loading section and hide landing/result sections. */
export function showLoading(message = "Loading...") {
  const loadingSection = $("loading");
  if (loadingSection) {
    loadingSection.style.display = "block";
  }
  setText("loading-message", message);

  const landing = $("landing-screen");
  if (landing) {
    landing.style.display = "none";
  }
  const result = $("result-section");
  if (result) {
    result.style.display = "none";
  }
}

/** Hide the loading section. */
export function hideLoading() {
  const loadingSection = $("loading");
  if (loadingSection) {
    loadingSection.style.display = "none";
  }
}

/** Apply a font-size token to the document root (small|medium|large). */
export function applyFontSize(size) {
  document.documentElement.setAttribute("data-font-size", size || "medium");
}

/** Toggle the "Generate reply" button visibility. */
export function updateReplyButtonVisibility(show) {
  const replyButton = $("generate-reply");
  if (replyButton) {
    replyButton.style.display = show ? "inline-block" : "none";
  }
}

/** Toggle the developer badge markers in the header/footer. */
export function updateDevBadges(show) {
  const devBadge = $("dev-badge");
  const footerDevBadge = $("footer-dev-badge");
  if (devBadge) {
    devBadge.style.display = show ? "block" : "none";
  }
  if (footerDevBadge) {
    footerDevBadge.style.display = show ? "block" : "none";
  }
}

/**
 * Flip the expand-button label/state once full content is ready (or reset).
 */
export function updateExpandButton(isFullContentVisible) {
  const expandButton = $("expand-content");
  if (!expandButton) {
    return;
  }

  expandButton.disabled = !isFullContentVisible;
  expandButton.classList.toggle("ms-Button--disabled", !isFullContentVisible);
  if (!isFullContentVisible) {
    expandButton.innerHTML = '<span class="ms-Button-label">Loading Full Content...</span>';
    expandButton.classList.remove("ms-Button--primary");
  } else {
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    expandButton.classList.add("ms-Button--primary");
  }
}

/**
 * Replace the loading spinner with full content and arm the expand button.
 */
export function updateResults(content) {
  const expandButton = $("expand-content");
  if (expandButton) {
    expandButton.disabled = false;
    expandButton.classList.remove("ms-Button--disabled");
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
  }

  const loadingContainer = $("loading-container");
  if (loadingContainer) {
    loadingContainer.remove();
  }

  const resultContent = $("result-content");
  if (resultContent) {
    resultContent.innerHTML = renderMarkdown(content);
  }
}

/** Expand/collapse the full-content container when the button is clicked. */
export function expandContent() {
  const expandButton = $("expand-content");
  if (!expandButton || expandButton.disabled) {
    return;
  }

  const fullContentContainer = $("full-content-container");
  if (!fullContentContainer) {
    return;
  }

  if (fullContentContainer.style.display === "none") {
    fullContentContainer.style.display = "block";
    expandButton.innerHTML = '<span class="ms-Button-label">Hide Full Content</span>';
    expandButton.classList.remove("ms-Button--primary");
  } else {
    fullContentContainer.style.display = "none";
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    expandButton.classList.add("ms-Button--primary");
  }
}

/**
 * Render the result panel for generated content.
 * In TL;DR mode the quick summary is shown first and the expand button armed
 * for the follow-up full content; otherwise full content is shown directly.
 */
export function showResults(content, type) {
  const resultContent = $("result-content");
  const tldrContent = $("tldr-content");
  if (resultContent) {
    resultContent.innerHTML = "";
  }
  if (tldrContent) {
    tldrContent.innerHTML = "";
  }

  // Reset the expand button to its disabled/loading state.
  const expandButton = $("expand-content");
  if (expandButton) {
    expandButton.disabled = true;
    expandButton.classList.add("ms-Button--disabled");
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    expandButton.classList.remove("ms-Button--primary");
  }

  const loading = $("loading");
  if (loading) {
    loading.style.display = "none";
  }
  const resultSection = $("result-section");
  if (resultSection) {
    resultSection.style.display = "block";
  }
  const landingScreen = $("landing-screen");
  if (landingScreen) {
    landingScreen.style.display = "none";
  }
  const appBody = $("app-body");
  if (appBody) {
    appBody.style.display = "block";
  }

  const tldrMode = getTldrModeSetting();

  if (tldrContent) {
    tldrContent.innerHTML = renderMarkdown(content);
  }

  if (tldrMode) {
    // Disable expand button and set to loading state until full content arrives.
    if (expandButton) {
      expandButton.disabled = true;
      expandButton.classList.add("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Loading Full Content...</span>';
    }
  } else {
    // Non-TL;DR: render full content immediately and arm the expand button.
    let fullEl = $("result-content");
    if (!fullEl) {
      fullEl = document.createElement("div");
      fullEl.id = "result-content";
      $("full-content-container")?.appendChild(fullEl);
    }
    fullEl.innerHTML = renderMarkdown(content);
    const fullContainer = $("full-content-container");
    if (fullContainer) {
      fullContainer.style.display = "block";
    }
    if (expandButton) {
      expandButton.disabled = false;
      expandButton.classList.remove("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
      expandButton.classList.add("ms-Button--primary");
    }
  }

  // Show/hide copy buttons based on the result type.
  const copyReplyButton = $("copy-reply");
  if (copyReplyButton) {
    copyReplyButton.style.display = type === TYPES.REPLY ? "inline-block" : "none";
  }
  const copyResultButton = $("copy-result");
  if (copyResultButton) {
    copyResultButton.style.display = type === TYPES.REPLY ? "none" : "inline-block";
  }
  const generateReplyButton = $("generate-reply");
  if (generateReplyButton) {
    const showReply = getSettings().showReply === "true";
    generateReplyButton.style.display =
      type === TYPES.REPLY ? "none" : showReply ? "inline-block" : "none";
  }

  // Re-apply font size + reply visibility from saved settings.
  try {
    const settings = getSettings();
    if (settings.fontSize) {
      applyFontSize(settings.fontSize);
    }
    if (settings.showReply) {
      updateReplyButtonVisibility(settings.showReply === "true");
    }
  } catch (error) {
    console.error("Error applying font size:", error);
  }

  if (resultContent) {
    resultContent.scrollTop = 0;
  }
  if (tldrContent) {
    tldrContent.scrollTop = 0;
  }
}

/** Copy the visible result to the clipboard, formatting replies correctly. */
export function copyResult() {
  let resultContent = "";

  const tldrContent = $("tldr-content");
  const fullContent = $("result-content");
  const tldrText = tldrContent ? tldrContent.innerText : "";
  const fullText = fullContent ? fullContent.innerText : "";

  const isReply = tldrText.includes("Subject:") || fullText.includes("Subject:");

  if (isReply) {
    let subject = "";
    let body = "";

    const headingMatch = /^(?:Subject:\s*)?(.+?)(?:\n|$)/i.exec(tldrText);
    if (headingMatch) {
      subject = headingMatch[1].trim();
    }

    const fullContainer = $("full-content-container");
    if (fullContainer && fullContainer.style.display !== "none") {
      body = fullText;
    } else if (headingMatch) {
      const lines = tldrText.split("\n");
      body = lines.slice(1).join("\n").trim();
    } else {
      body = tldrText;
    }

    resultContent = `Subject: ${subject}\n\n${body}`;
  } else {
    const fullContainer = $("full-content-container");
    if (fullContainer && fullContainer.style.display !== "none") {
      resultContent = fullText;
    } else {
      resultContent = tldrText;
    }
  }

  navigator.clipboard
    .writeText(resultContent)
    .then(() => {
      setText("copy-status", "Copied!");
      setTimeout(() => setText("copy-status", ""), 2000);
    })
    .catch((err) => {
      console.error("Could not copy text: ", err);
      showNotification("Failed to copy to clipboard", "error");
    });
}

/** Copy the formatted reply (Subject + body) to the clipboard. */
export function copyReply() {
  const tldrContent = $("tldr-content");
  const tldrText = tldrContent ? tldrContent.innerText : "";

  const subjectMatch = tldrText.match(/Subject:\s*(.+?)(?:\n|$)/i);
  const subject = subjectMatch ? subjectMatch[1].trim() : "";

  const bodyStart = tldrText.indexOf(subject) + subject.length;
  const body = tldrText.substring(bodyStart).trim();

  const replyContent = `Subject: ${subject}\n\n${body}`;

  navigator.clipboard
    .writeText(replyContent)
    .then(() => {
      setText("copy-reply-status", "Copied!");
      setTimeout(() => setText("copy-reply-status", ""), 2000);
    })
    .catch((err) => {
      console.error("Could not copy reply: ", err);
      showNotification("Failed to copy reply", "error");
    });
}

/**
 * Parse a raw reply into structured { subject, body, html, raw }.
 * Looks for an explicit SUBJECT: marker, else treats the first line as subject.
 */
export function formatReplyOutput(replyText) {
  let subject = "";
  let body = "";

  const subjectMatch = replyText.match(/^(?:SUBJECT:|Subject:)\s*(.+?)(?:\n|$)/m);
  if (subjectMatch) {
    subject = subjectMatch[1].trim();
    body = replyText.replace(/^(?:SUBJECT:|Subject:)\s*.+?\n+/m, "").trim();
  } else {
    const lines = replyText.trim().split("\n");
    if (lines.length > 0) {
      subject = lines[0].trim();
      body = lines.slice(1).join("\n").trim();
    } else {
      subject = "Re: Your email";
      body = replyText.trim();
    }
  }

  const formattedHtml = `
    <div class="reply-container">
      <div class="reply-subject">
        <span class="reply-label">Subject:</span>${escapeHtml(subject)}
      </div>
      <div class="reply-body">${renderMarkdown(body)}</div>
    </div>
  `;

  return {
    html: formattedHtml,
    subject,
    body,
    raw: `Subject: ${subject}\n\n${body}`,
  };
}

/** Read the TL;DR-mode flag (defaults to true when unset — original behavior). */
export function getTldrModeSetting() {
  try {
    const settings = getSettings();
    if (settings.tldrMode) {
      return settings.tldrMode === "true";
    }
  } catch (error) {
    console.error("Error getting TLDR mode setting:", error);
  }
  return true;
}
