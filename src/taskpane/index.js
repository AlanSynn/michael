/* global document, Office, console, window, setTimeout, clearTimeout */

// Taskpane entry point. Owns Office.onReady + DOM listener wiring + the single
// flat ItemChanged handler. (Earlier builds nested addHandlerAsync/register
// chains that were dead; they were collapsed into one flat registration.)

import { migrateSettingsKeys } from "./storage.js";
import {
  onItemChanged,
  onSettingsChanged as registerSettingsChanged,
  clearEmailContentCache,
} from "./mailbox.js";
import {
  loadDropdownSettings,
  updateAuthenticationStatus,
  refreshModelCatalog,
  applyCurrentTheme,
  onSettingsChanged,
  initializeSettingsTabs,
  initSettingsInteractions,
  toggleSettingsView,
  saveDropdownSettings,
  resetAllSettings,
  resetTemplates,
  loadBuiltInTemplateDefaults,
  saveCurrentTemplatesAsDefaults,
  clearTemplates,
  copyAllTemplatesToClipboard,
  exportTemplatesAsMarkdown,
} from "./ui/settings-view.js";
import {
  summarizeEmail,
  translateEmail,
  translateAndSummarizeEmail,
  generateReply,
  runAutorun,
} from "./ui/flows.js";
import { handleCalendarEvent, updateCalendarButtonState } from "./ui/calendar.js";
import { expandContent, copyResult, copyReply } from "./ui/dom.js";

/** Attach a listener only if the element exists. */
function on(id, eventName, handler) {
  const el = document.getElementById(id);
  if (el) {
    el.addEventListener(eventName, handler);
  }
}

/** Wire every button/input listener used by the taskpane. Idempotent: a
 *  bootstrap retry (e.g. both Office.initialize and Office.onReady firing)
 *  must not double-register handlers. */
let listenersWired = false;
function wireListeners() {
  if (listenersWired) return;
  listenersWired = true;

  on("summarize", "click", summarizeEmail);
  on("translate", "click", translateEmail);
  on("translate-summarize", "click", translateAndSummarizeEmail);
  on("calendar-event", "click", handleCalendarEvent);
  on("settings-toggle", "click", toggleSettingsView);
  on("close-settings-view", "click", toggleSettingsView);

  on("dropdown-save-settings", "click", saveDropdownSettings);
  on("dropdown-reset-all", "click", resetAllSettings);

  // Prompt-template toolbar (lives in the Templates tab).
  on("dropdown-reset-templates", "click", resetTemplates);
  on("dropdown-load-builtins", "click", loadBuiltInTemplateDefaults);
  on("dropdown-save-template-defaults", "click", saveCurrentTemplatesAsDefaults);
  on("dropdown-clear-templates", "click", clearTemplates);
  on("dropdown-copy-templates", "click", copyAllTemplatesToClipboard);
  on("dropdown-export-markdown", "click", exportTemplatesAsMarkdown);
  on("dropdown-refresh-models", "click", () => refreshModelCatalog({ force: true }));

  // Auth status updates per keystroke; a debounced live model refresh after
  // the user settles on a key invalidates the cached catalog (which was served
  // from the prior/empty key) so the dropdowns reflect the new credentials.
  let keyRefreshTimer = 0;
  on("dropdown-api-key", "input", () => {
    updateAuthenticationStatus();
    clearTimeout(keyRefreshTimer);
    keyRefreshTimer = setTimeout(() => refreshModelCatalog({ force: true, silent: true }), 600);
  });

  // Dev-mode toggle reveals the dev-server field group.
  on("dropdown-dev-mode", "change", (event) => {
    const value = event.target.value;
    const devServerGroup = document.getElementById("dev-server-group");
    if (devServerGroup) {
      devServerGroup.style.display = value === "true" ? "block" : "none";
    }
  });

  on("expand-content", "click", expandContent);
  on("copy-reply", "click", copyReply);
  on("copy-result", "click", copyResult);
  on("generate-reply", "click", generateReply);
}

function bootstrapOutlook() {
  migrateSettingsKeys();

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  wireListeners();
  loadDropdownSettings();
  updateAuthenticationStatus();
  refreshModelCatalog({ silent: true });
  applyCurrentTheme();
  initializeSettingsTabs();
  initSettingsInteractions();

  // Re-apply the theme when Office settings change (only when following Office).
  registerSettingsChanged(onSettingsChanged);

  // ONE flat ItemChanged handler: clear the per-item body cache, autorun the
  // configured action, and refresh the calendar-button state for the newly
  // selected item.
  onItemChanged(() => {
    clearEmailContentCache();
    runAutorun();
    updateCalendarButtonState();
  });

  // Initial state for the current item.
  runAutorun();
  updateCalendarButtonState();
}

// --- Office initialization ---------------------------------------------------
// Office.onReady is the modern init handshake. We log every step because a
// silent failure here — host mismatch, a thrown bootstrap, or office.js never
// resolving — otherwise leaves the pane stuck on the sideload screen with no
// clue in the console. Known silent-init failure modes: office-js #4735
// (onReady never fires), #3330 (slow CDN blocks init), and a blocked
// MicrosoftAjax.js dependency (theofficecontext 2025). All produce a dead pane
// whose only signal is the console.
let bootstrapped = false;

function runBootstrap(source) {
  if (bootstrapped) return;
  try {
    bootstrapOutlook();
    bootstrapped = true;
    console.info("[Michael] bootstrap complete via " + source);
  } catch (err) {
    console.error("[Michael] bootstrap FAILED via " + source + ":", err);
    // Surface the failure on the pane itself so a dead init is not invisible.
    const msg = document.getElementById("sideload-msg");
    if (msg) {
      msg.textContent =
        "Michael failed to start: " +
        (err && err.message ? err.message : String(err)) +
        ". Open Safari → Develop → [Mac] → Outlook to inspect the console.";
    }
  }
}

// Fallback init handshake. Some classic hosts fire Office.initialize but never
// resolve Office.onReady (office-js #4735, a blocked MicrosoftAjax.js). The
// bootstrapped guard above makes whichever fires first authoritative and the
// other a no-op, so registering both is safe.
if (typeof Office !== "undefined") {
  // Deliberate resilience fallback for hosts where onReady never fires
  // (office-js #4735); onReady alone is the documented preference, but
  // registering initialize too is safe given the bootstrapped guard above.
  // eslint-disable-next-line office-addins/no-office-initialize
  Office.initialize = () => runBootstrap("initialize");
}

// Guard onReady: if office.js loaded only partially (e.g. a blocked dependency
// left Office defined but Office.onReady missing), calling it would throw and
// leave the pane dead. Fall back to the Office.initialize path above.
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    const expectedHost = Office.HostType && Office.HostType.Outlook;
    console.info("[Michael] Office.onReady fired:", {
      host: info && info.host,
      platform: info && info.platform,
      expectedOutlookHost: expectedHost,
    });
    if (info && expectedHost !== undefined && info.host === expectedHost) {
      runBootstrap("Office.onReady");
    } else {
      const detectedHost = info && info.host ? info.host : "unknown";
      console.warn(
        "[Michael] Office.onReady fired but host is not Outlook (detected: " +
          detectedHost +
          ") — staying on the sideload screen."
      );
      // Surface an actionable message instead of the generic sideload logo so a
      // misconfigured launch (e.g. opened in a plain browser tab) is not a dead end.
      const msg = document.getElementById("sideload-msg");
      if (msg) {
        msg.textContent =
          "Michael can only run inside Outlook. Close this pane and reopen it from Outlook" +
          (detectedHost !== "unknown" ? " (detected host: " + detectedHost + ")." : ".");
      }
    }
  });
} else {
  console.warn("[Michael] Office.onReady unavailable — relying on the Office.initialize fallback.");
}

// Surface any uncaught error or rejected promise so a dead pane always leaves
// a trace in the Web Inspector console.
window.addEventListener("error", (event) => {
  console.error("[Michael] uncaught error:", event.error || event.message);
});
window.addEventListener("unhandledrejection", (event) => {
  console.error("[Michael] unhandled promise rejection:", event.reason);
});
