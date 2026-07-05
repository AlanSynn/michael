/* global document, console, navigator, Office, AbortController */

// Calendar feature: detect whether an email describes an event, parse the
// event details via Z.AI, and open a prefilled Outlook appointment form.

import { generateContent, requireTemplate } from "../generation.js";
import { getCurrentDateContext } from "../date-context.js";
import { getEmailContent } from "../mailbox.js";
import { getLanguageText } from "../language.js";
import { fillTemplate } from "../prompts.js";
import { getSettings } from "../storage.js";
import { DEFAULT_SETTINGS } from "../prompt-templates.js";
import {
  showLoading,
  hideLoading,
  showNotification,
  escapeHtml,
  getApiKey,
  stampResultType,
} from "./dom.js";
import { toggleSettingsView, getMissingApiKeyMessage } from "./settings-view.js";
import { beginFlow, isActiveFlow, isCancelError } from "./flow-control.js";

function getEventTitleLanguage() {
  const settings = getSettings();
  return settings.eventTitleLanguage || DEFAULT_SETTINGS.eventTitleLanguage;
}

/** Build the language-specific instructions for the event-title prompt. */
function buildTitleLanguageInstructions(titleLanguage) {
  if (titleLanguage === "en") {
    return `Event title should be in English.
      If the event has a type or category, include it in square brackets ([]) at the beginning, then if there's a presenter and topic, write the presenter's name first, followed by a hyphen (-) and then the topic.`;
  }
  if (titleLanguage === "ko") {
    return `이벤트 제목은 한국어로 작성해주세요.
      이벤트 유형이나 카테고리가 있다면 대괄호([])로 먼저 표시하고, 발표자와 주제가 있다면 발표자 이름을 먼저 쓰고 하이픈(-) 후에 주제를 적어주세요.`;
  }
  if (titleLanguage === "ja") {
    return `イベントのタイトルは日本語で記載してください。
      イベントのタイプやカテゴリがある場合は、角括弧（[]）で囲んで最初に表示し、発表者とトピックがある場合は、発表者の名前を最初に書き、ハイフン（-）の後にトピックを書いてください。`;
  }
  if (titleLanguage === "zh_cn") {
    return `事件标题应该用中文书写。
      如果事件有类型或类别，请使用方括号（[]）将其括起来并放在开头，如果有演讲者和主题，请先写演讲者的名字，然后是连字符（-），再写主题。`;
  }
  // Default English instructions parametrized by language name.
  return `Event title should be in ${getLanguageText(titleLanguage)}.
      If the event has a type or category, include it in square brackets ([]) at the beginning, then if there's a presenter and topic, write the presenter's name first, followed by a hyphen (-) and then the topic.`;
}

/**
 * Accept the date/time strings Z.AI actually emits. Models frequently return
 * ISO 8601 with a trailing "Z", a "+09:00" offset, or fractional seconds
 * (".000"). The prior strict regex required a bare YYYY-MM-DDTHH:mm:ss and
 * rejected all of those, so valid extractions surfaced as "extraction failed".
 * Validate with `new Date(...)` so the downstream displayNewAppointmentForm
 * gets a real timestamp either way.
 */
function isValidDateTime(value) {
  if (typeof value !== "string" || !value) {
    return false;
  }
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?$/.test(value)) {
    return false;
  }
  return !isNaN(new Date(value).getTime());
}

/** Parse event details (JSON) from email content using Z.AI. */
async function parseEventDetailsWithZai(emailContent, signal = null) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      throw new Error(getMissingApiKeyMessage());
    }

    const langInstructions = buildTitleLanguageInstructions(getEventTitleLanguage());

    const prompt = fillTemplate(requireTemplate("calendarParse", "Calendar parse"), {
      currentDate: getCurrentDateContext(),
      languageInstructions: langInstructions,
      content: emailContent,
    });

    const result = await generateContent(prompt, apiKey, null, false, signal);

    // Extract only the JSON object from the response.
    let jsonText = result;
    const jsonMatch = result.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      jsonText = jsonMatch[0];
    }

    try {
      const eventDetails = JSON.parse(jsonText);

      if (!eventDetails.subject) {
        throw new Error("Event title not found.");
      }
      if (!eventDetails.start?.dateTime) {
        throw new Error("Event start time not found.");
      }
      if (!eventDetails.end?.dateTime) {
        throw new Error("Event end time not found.");
      }

      if (!isValidDateTime(eventDetails.start.dateTime)) {
        throw new Error("Invalid start time format. (Expected YYYY-MM-DDTHH:mm:ss)");
      }
      if (!isValidDateTime(eventDetails.end.dateTime)) {
        throw new Error("Invalid end time format. (Expected YYYY-MM-DDTHH:mm:ss)");
      }

      return eventDetails;
    } catch (parseError) {
      throw new Error("Failed to extract event information: " + parseError.message);
    }
  } catch (error) {
    console.error("Error extracting event information:", error);
    throw error;
  }
}

/** Open the Outlook "new appointment" form prefilled with the parsed event. */
function createCalendarEvent(eventDetails) {
  try {
    if (!eventDetails.subject || !eventDetails.start?.dateTime || !eventDetails.end?.dateTime) {
      throw new Error("Required event information is missing.");
    }

    const startDate = new Date(eventDetails.start.dateTime);
    const endDate = new Date(eventDetails.end.dateTime);

    const requiredAttendees = [];
    const optionalAttendees = [];
    if (eventDetails.attendees && eventDetails.attendees.length > 0) {
      eventDetails.attendees.forEach((attendee) => {
        if (attendee.emailAddress && attendee.emailAddress.address) {
          if (attendee.type === "optional") {
            optionalAttendees.push(attendee.emailAddress.address);
          } else {
            requiredAttendees.push(attendee.emailAddress.address);
          }
        }
      });
    }

    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees,
      optionalAttendees,
      start: startDate,
      end: endDate,
      location: eventDetails.location?.displayName || "",
      body: eventDetails.body?.content || "",
      subject: eventDetails.subject,
    });

    showNotification(`Event '${eventDetails.subject}' has been created.`, "info");
    return true;
  } catch (error) {
    showNotification(`Failed to create event: ${error.message}`, "error");
    console.error("Error creating calendar event:", error);
    throw error;
  }
}

/** Ask Z.AI whether the email content describes a calendar event. */
export async function checkIfCalendarEvent(emailContent, signal = null) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      console.error(getMissingApiKeyMessage());
      return false;
    }

    const prompt = fillTemplate(requireTemplate("calendarCheck", "Calendar check"), {
      currentDate: getCurrentDateContext(),
      content: emailContent,
    });

    const result = await generateContent(prompt, apiKey, null, true, signal);
    return result.toLowerCase().trim() === "true";
  } catch (error) {
    if (error && error.name === "AbortError") {
      throw error; // caller cancel — propagate, do not treat as "not an event"
    }
    console.error("Error checking if calendar event:", error);
    return false;
  }
}

/** Handle the calendar-event button: parse + open the appointment form. */
export async function handleCalendarEvent() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView();
    return;
  }

  // Join the shared single-flight controller so starting an email flow (or
  // vice-versa) cancels an in-flight calendar parse and they cannot race on
  // the result-section DOM.
  const signal = beginFlow();
  showLoading("Creating calendar event...");

  try {
    const emailContent = await getEmailContent();
    if (!isActiveFlow(signal)) return;

    const landingScreen = document.getElementById("landing-screen");
    const resultSection = document.getElementById("result-section");
    if (landingScreen) {
      landingScreen.style.display = "none";
    }
    if (resultSection) {
      resultSection.style.display = "block";
    }
    // Calendar output is neither a reply nor a copyable "result" — stamp the
    // type so copyResult/copyReply do not read a stale "reply" stamp from a
    // prior run, and hide the reply-only copy button.
    stampResultType("calendar");
    const copyReplyBtn = document.getElementById("copy-reply");
    if (copyReplyBtn) {
      copyReplyBtn.style.display = "none";
    }

    const tldrContent = document.getElementById("tldr-content");
    if (tldrContent) {
      tldrContent.innerHTML = `
        <div class="email-content-header">
          <h3>Email Content</h3>
          <button id="copy-email-content" class="ms-Button ms-Button--primary">
            <span class="ms-Button-label">Copy to Clipboard</span>
          </button>
        </div>
        <div class="email-content-body">
          <pre style="white-space: pre-wrap; word-break: break-word;">${escapeHtml(emailContent)}</pre>
        </div>
      `;
    }

    const copyEmailBtn = document.getElementById("copy-email-content");
    if (copyEmailBtn) {
      copyEmailBtn.addEventListener("click", () => {
        navigator.clipboard
          .writeText(emailContent)
          .then(() => showNotification("Email content copied to clipboard", "info"))
          .catch((err) => {
            console.error("Error copying to clipboard:", err);
            showNotification("Failed to copy to clipboard", "error");
          });
      });
    }

    try {
      const eventDetails = await parseEventDetailsWithZai(emailContent, signal);
      if (!isActiveFlow(signal)) return;

      const resultContent = document.getElementById("result-content");
      if (resultContent) {
        resultContent.innerHTML = `
          <div class="event-details-header">
            <h3>Extracted Event Details</h3>
          </div>
          <div class="event-details-body">
            <p><strong>Subject:</strong> ${escapeHtml(eventDetails.subject || "Not found")}</p>
            <p><strong>Start:</strong> ${escapeHtml(eventDetails.start?.dateTime || "Not found")}</p>
            <p><strong>End:</strong> ${escapeHtml(eventDetails.end?.dateTime || "Not found")}</p>
            <p><strong>Location:</strong> ${escapeHtml(eventDetails.location?.displayName || "Not found")}</p>
          </div>
        `;
      }

      await createCalendarEvent(eventDetails);
    } catch (extractionError) {
      // A superseded parse surfaces an AbortError — do NOT render "extraction
      // failed" for a cancel the user initiated by starting another flow. The
      // isActiveFlow check also covers a late non-Abort rejection from a
      // cancelled parse (e.g. body.getAsync rejecting after the item changed).
      if (isCancelError(extractionError) || !isActiveFlow(signal)) {
        return;
      }
      console.error("Event extraction error:", extractionError);
      const errorMessage =
        extractionError && extractionError.message
          ? extractionError.message
          : String(extractionError);
      const cleanedMessage = errorMessage.includes("Failed to extract event information:")
        ? errorMessage.split("Failed to extract event information:")[1].trim()
        : errorMessage;

      const resultContent = document.getElementById("result-content");
      if (resultContent) {
        resultContent.innerHTML = `
          <div class="event-details-header">
            <h3>Event Extraction Failed</h3>
          </div>
          <div class="event-details-body">
            <p class="error-message">${escapeHtml(cleanedMessage)}</p>
          </div>
        `;
      }

      showNotification(`Event extraction failed: ${cleanedMessage}`, "error");
    }
  } catch (error) {
    if (isCancelError(error) || !isActiveFlow(signal)) {
      return;
    }
    console.error("Calendar event handling error:", error);
    const message = error && error.message ? error.message : String(error);
    showNotification(`Error: ${message}`, "error");
  } finally {
    // Only the active flow owns the loading spinner + expand button. A
    // superseded calendar parse must not reset them and clobber the newer
    // flow's UI.
    if (isActiveFlow(signal)) {
      hideLoading();
      updateCalendarButtonState();

      const expandButton = document.getElementById("expand-content");
      if (expandButton) {
        expandButton.disabled = false;
        expandButton.classList.remove("ms-Button--disabled");
        expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
        expandButton.classList.add("ms-Button--primary");

        const fullContentContainer = document.getElementById("full-content-container");
        if (fullContentContainer) {
          fullContentContainer.style.display = "block";
        }
      }
    }
  }
}

// --- calendar-event auto-detection -----------------------------------------
// Detecting whether an email is a calendar event calls Z.AI (a paid
// chat-completion that ships the full email body). Doing that on every email
// open + bootstrap is a cost, privacy, and stale-UI firehose. So detection is
// OPT-IN (autoDetectCalendarEvents setting, default off): when off the button
// stays enabled and the user triggers extraction on demand; when on we probe,
// memoize the result per itemId, and cancel the in-flight probe on ItemChanged.

const calendarProbeCache = new Map(); // itemId -> boolean
const CALENDAR_PROBE_CACHE_MAX = 50;
let calendarProbeController = null;

function isAutoDetectEnabled() {
  const settings = getSettings();
  return settings.autoDetectCalendarEvents === "true";
}

function getCurrentItemId() {
  try {
    return (Office.context.mailbox.item && Office.context.mailbox.item.itemId) || null;
  } catch {
    return null;
  }
}

function applyCalendarButtonState(calendarBtn, isCalendarEvent) {
  if (isCalendarEvent) {
    calendarBtn.disabled = false;
    calendarBtn.classList.remove("action-button--disabled");
    calendarBtn.classList.add("action-button--primary");
  } else {
    calendarBtn.disabled = true;
    calendarBtn.classList.add("action-button--disabled");
    calendarBtn.classList.remove("action-button--primary");
  }
}

/**
 * Enable/disable the calendar button based on whether the email is an event.
 * When auto-detect is off (default), the button is simply enabled. When on,
 * the result is memoized per itemId and the probe is aborted on re-entry so
 * rapid email switching cannot pile up requests or overwrite the button state.
 */
export async function updateCalendarButtonState() {
  const calendarBtn = document.getElementById("calendar-event");
  if (!calendarBtn) {
    return;
  }

  if (!isAutoDetectEnabled()) {
    applyCalendarButtonState(calendarBtn, true);
    return;
  }

  const itemId = getCurrentItemId();
  if (itemId && calendarProbeCache.has(itemId)) {
    applyCalendarButtonState(calendarBtn, calendarProbeCache.get(itemId));
    return;
  }

  // Abort any prior in-flight probe (e.g. the user switched emails mid-probe).
  if (calendarProbeController) {
    calendarProbeController.abort();
  }
  calendarProbeController = new AbortController();
  const { signal } = calendarProbeController;

  try {
    const emailContent = await getEmailContent();
    if (signal.aborted) return;
    const isCalendarEvent = await checkIfCalendarEvent(emailContent, signal);
    if (signal.aborted) return;
    if (itemId) {
      if (calendarProbeCache.size >= CALENDAR_PROBE_CACHE_MAX) {
        calendarProbeCache.delete(calendarProbeCache.keys().next().value);
      }
      calendarProbeCache.set(itemId, isCalendarEvent);
    }
    applyCalendarButtonState(calendarBtn, isCalendarEvent);
  } catch (error) {
    if (error && error.name === "AbortError") return;
    // No selected item is benign (user is on the mailbox list), not a probe
    // failure — stay quiet, mirroring the flows' NO_ITEM handling.
    if (error && error.code === "NO_ITEM") return;
    console.error("Error updating calendar button state:", error);
  }
}
