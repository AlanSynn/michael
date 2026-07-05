/* global document, Office */

// Taskpane entry point. Owns Office.onReady + DOM listener wiring + the single
// flat ItemChanged handler (R3: the prior nested addHandlerAsync/register
// chains were dead). Dead initializeApp/loadSettings/setTheme/setFontSize were
// dropped (R9) — every section already has the correct initial inline state.

import { migrateSettingsKeys } from "./storage.js";
import { onItemChanged, onSettingsChanged as registerSettingsChanged } from "./mailbox.js";
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

/** Wire every button/input listener used by the taskpane. */
function wireListeners() {
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
  on("dropdown-refresh-models", "click", () => refreshModelCatalog());
  on("dropdown-api-key", "input", updateAuthenticationStatus);

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

  // ONE flat ItemChanged handler (R3): autorun the configured action and refresh
  // the calendar-button state for the newly selected item.
  onItemChanged(() => {
    runAutorun();
    updateCalendarButtonState();
  });

  // Initial state for the current item.
  runAutorun();
  updateCalendarButtonState();
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    bootstrapOutlook();
  }
});
