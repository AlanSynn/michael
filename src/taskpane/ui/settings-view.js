/* global document, console, navigator, setTimeout, URL, Blob, Event, Office */

// Settings panel controller: tab switching, form load/save, model-catalog UI
// glue, theme application (resolveTheme is the single fallback source), and
// the prompt-template toolbar. Reaches into storage + generation facades.

import { fetchAvailableModels, getDefaultZaiModels } from "../../shared/zai.js";
import {
  createBlankPromptTemplates,
  createBlankSettings,
  createDefaultPromptTemplates,
  DEFAULT_SETTINGS,
  PROVIDER_MESSAGES,
  TEMPLATE_KEYS,
} from "../prompt-templates.js";
import { resolveTheme, DEFAULT_THEME_FALLBACK } from "../theme.js";
import {
  getSettings,
  saveSettingsToStore,
  roamingStorage,
  saveRoamingStoreAsync,
  getSavedTemplateDefaults,
  saveTemplateDefaults,
  readCachedModelCatalog,
  writeCachedModelCatalog,
  SETTINGS_KEY,
  TEMPLATE_DEFAULTS_KEY,
  MODEL_CATALOG_CACHE_KEY,
} from "../storage.js";
import {
  getCatalog,
  setCatalog,
  getCachedModelCatalog as resolveCachedCatalog,
} from "../model-catalog.js";
import {
  getAssetPath,
  applyFontSize,
  updateReplyButtonVisibility,
  updateDevBadges,
  showNotification,
  setText,
  getApiKey,
  $,
} from "./dom.js";

const PROMPT_FIELD_MAP = Object.freeze({
  summarize: "dropdown-summarize-template",
  translate: "dropdown-translate-template",
  translateSummarize: "dropdown-translate-summarize-template",
  reply: "dropdown-reply-template",
  commandTranslate: "dropdown-command-translate-template",
  tldrPrompt: "dropdown-tldr-template",
  calendarParse: "dropdown-calendar-parse-template",
  calendarCheck: "dropdown-calendar-check-template",
});

/** Read the model catalog from cache, resolving to defaults when empty. */
function getCachedCatalog() {
  return resolveCachedCatalog(readCachedModelCatalog());
}

/** Persist a model list to the in-memory catalog + roaming cache. */
function persistModelCatalog(models) {
  setCatalog(models);
  writeCachedModelCatalog(getCatalog());
}

// --- Settings view visibility + tabs ----------------------------------------

/** Toggle the settings panel over the main content area. */
export function toggleSettingsView() {
  const settingsView = $("settings-view");
  const appBody = $("app-body");

  if (!settingsView || !appBody) {
    return;
  }

  const isVisible = settingsView.style.display === "block";
  if (isVisible) {
    settingsView.style.display = "none";
    appBody.style.display = "flex";
  } else {
    settingsView.style.display = "block";
    appBody.style.display = "none";
    loadDropdownSettings();
    refreshModelCatalog({ silent: true });
  }
}

/** Wire up the Settings tab buttons and activate the first tab. */
export function initializeSettingsTabs() {
  const tabButtons = document.querySelectorAll(".settings-tabs .settings-tab-button");
  const tabContents = document.querySelectorAll(".settings-content .settings-tab-content");

  if (!tabButtons.length || !tabContents.length) {
    return;
  }

  const activate = (button) => {
    tabButtons.forEach((btn) => btn.classList.remove("active"));
    tabContents.forEach((content) => content.classList.remove("active"));
    button.classList.add("active");
    const target = $(button.getAttribute("data-tab"));
    if (target) {
      target.classList.add("active");
    } else {
      console.error("Target tab content not found:", button.getAttribute("data-tab"));
    }
  };

  tabButtons.forEach((button) => {
    button.addEventListener("click", () => activate(button));
  });

  // Ensure the first tab is active on initialization.
  const firstTabButton = tabButtons[0];
  if (firstTabButton) {
    activate(firstTabButton);
  }
}

// --- Segmented toggles + accordions -----------------------------------------

/**
 * Sync every segmented toggle's active button from its bound <select>. The
 * <select> stays the single source of truth for load/save; the toggle is a
 * visual layer on top.
 */
export function syncSegmentedToggles() {
  document.querySelectorAll("[data-segmented]").forEach((group) => {
    const selectId = group.getAttribute("data-segmented");
    const select = selectId ? $(selectId) : null;
    if (!select) {
      return;
    }
    const value = select.value;
    group.querySelectorAll(".segmented-btn").forEach((btn) => {
      const isActive = btn.getAttribute("data-value") === value;
      btn.classList.toggle("active", isActive);
      btn.setAttribute("aria-pressed", String(isActive));
    });
  });
  updateAutorunOptionState();
}

/** Dim + disable the auto-run action row whenever auto-run is off. */
function updateAutorunOptionState() {
  const autorun = $("dropdown-autorun");
  const row = $("autorun-option-row");
  if (autorun && row) {
    row.classList.toggle("is-disabled", autorun.value !== "true");
  }
}

/**
 * Wire the segmented toggles (write back to their <select> + emit `change` so
 * existing listeners — e.g. dev-mode revealing the dev-server group — still
 * fire) and the template accordions. Call once during bootstrap.
 */
export function initSettingsInteractions() {
  document.querySelectorAll(".segmented").forEach((group) => {
    group.addEventListener("click", (event) => {
      const btn = event.target.closest(".segmented-btn");
      if (!btn || !group.contains(btn)) {
        return;
      }
      const selectId = group.getAttribute("data-segmented");
      const select = selectId ? $(selectId) : null;
      if (!select) {
        return;
      }
      select.value = btn.getAttribute("data-value");
      select.dispatchEvent(new Event("change", { bubbles: true }));
      syncSegmentedToggles();
    });
  });

  document.querySelectorAll("[data-accordion]").forEach((accordion) => {
    accordion.addEventListener("click", (event) => {
      const header = event.target.closest("[data-accordion-toggle]");
      if (!header || !accordion.contains(header)) {
        return;
      }
      header.closest("[data-accordion-item]").classList.toggle("open");
    });
  });
}

// --- Prompt template form glue ----------------------------------------------

function getPromptFieldValue(templateKey) {
  const fieldId = PROMPT_FIELD_MAP[templateKey];
  const field = fieldId ? $(fieldId) : null;
  return field ? field.value : "";
}

function setPromptFieldValue(templateKey, value) {
  const fieldId = PROMPT_FIELD_MAP[templateKey];
  const field = fieldId ? $(fieldId) : null;
  if (field) {
    field.value = value || "";
  }
}

function collectPromptTemplatesFromForm() {
  return TEMPLATE_KEYS.reduce((templates, key) => {
    templates[key] = getPromptFieldValue(key);
    return templates;
  }, {});
}

function applyPromptTemplatesToForm(templates) {
  TEMPLATE_KEYS.forEach((key) => {
    setPromptFieldValue(key, templates[key] || "");
  });
}

// --- Auth + model catalog UI ------------------------------------------------

export function getMissingApiKeyMessage() {
  return PROVIDER_MESSAGES.missingApiKey;
}

export function updateAuthenticationStatus() {
  if (getApiKey()) {
    setText("dropdown-auth-status", "Authentication source: saved Outlook add-in settings.");
    return;
  }

  setText(
    "dropdown-auth-status",
    "Authentication source: empty. Enter the API key in this screen and save settings."
  );
}

function updateModelSelectOptions(selectId, models, selectedValue) {
  const select = $(selectId);
  if (!select) {
    return;
  }

  const catalog = Array.isArray(models) && models.length ? models : getDefaultZaiModels();
  const normalizedSelection = typeof selectedValue === "string" ? selectedValue.trim() : "";
  const nextValue =
    normalizedSelection && catalog.includes(normalizedSelection) ? normalizedSelection : "";

  select.innerHTML = "";
  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = "Select a model";
  select.appendChild(placeholderOption);
  catalog.forEach((modelName) => {
    const option = document.createElement("option");
    option.value = modelName;
    option.textContent = modelName;
    select.appendChild(option);
  });

  select.value = nextValue || "";
}

function syncModelDropdowns() {
  const settings = getSettings();
  const models = getCatalog();

  updateModelSelectOptions("dropdown-model", models, settings.model || "");
  updateModelSelectOptions("dropdown-reply-model", models, settings.replyModel || "");
}

/** Refresh the Z.AI model catalog (live when an API key is present, else cached). */
export async function refreshModelCatalog(options = {}) {
  const silent = options.silent === true;
  const apiKey = getApiKey();

  updateAuthenticationStatus();

  if (!apiKey) {
    setCatalog(getCachedCatalog());
    syncModelDropdowns();
    setText(
      "dropdown-model-status",
      "Using cached/fallback Z.AI models. Enter an API key in Settings > General to enable live refresh."
    );
    return getCatalog();
  }

  setText("dropdown-model-status", "Refreshing available Z.AI models...");

  try {
    const liveModels = await fetchAvailableModels({ apiKey });
    persistModelCatalog(liveModels);
    syncModelDropdowns();
    setText("dropdown-model-status", `Loaded ${liveModels.length} models from Z.AI.`);
    return getCatalog();
  } catch (error) {
    console.error("Error refreshing Z.AI model catalog:", error);
    setCatalog(getCachedCatalog());
    syncModelDropdowns();
    setText(
      "dropdown-model-status",
      `Live refresh failed. Using cached/fallback models (${error.message}).`
    );
    if (!silent) {
      showNotification(`Model refresh failed: ${error.message}`, "warning");
    }
    return getCatalog();
  }
}

// --- Settings load / save / reset -------------------------------------------

/** Load saved settings into the Settings form fields. */
export function loadDropdownSettings() {
  try {
    const settings = getSettings();
    const sessionDefaults = getSavedTemplateDefaults();

    $("dropdown-api-key").value = settings.apiKey || "";
    if (settings.defaultLanguage) $("dropdown-language").value = settings.defaultLanguage;
    if (settings.eventTitleLanguage)
      $("dropdown-event-title-language").value = settings.eventTitleLanguage;
    if (settings.theme) $("dropdown-theme").value = settings.theme;
    if (settings.fontSize) $("dropdown-font-size").value = settings.fontSize;
    if (settings.tldrMode) $("dropdown-tldr-mode").value = settings.tldrMode;
    if (settings.showReply) $("dropdown-show-reply").value = settings.showReply;
    if (settings.devServer !== undefined) $("dropdown-dev-server").value = settings.devServer;
    if (settings.autorun) $("dropdown-autorun").value = settings.autorun;
    if (settings.autorunOption) $("dropdown-autorun-option").value = settings.autorunOption;
    if (settings.devMode) $("dropdown-dev-mode").value = settings.devMode;

    const devServerGroup = $("dev-server-group");
    if (devServerGroup) {
      devServerGroup.style.display = settings.devMode === "true" ? "block" : "none";
    }

    updateDevBadges(settings.devMode === "true");
    applyFontSize(settings.fontSize || DEFAULT_SETTINGS.fontSize);
    updateReplyButtonVisibility(settings.showReply === "true");

    const templates =
      settings.templates && Object.keys(settings.templates).length
        ? settings.templates
        : sessionDefaults;
    applyPromptTemplatesToForm({
      ...createBlankPromptTemplates(),
      ...templates,
    });

    setCatalog(getCachedCatalog());
    syncModelDropdowns();
    updateAuthenticationStatus();
    syncSegmentedToggles();
  } catch (error) {
    console.error("Error loading dropdown settings:", error);
    applyPromptTemplatesToForm(createBlankPromptTemplates());
  }
}

/** Persist the Settings form into roaming settings and re-apply live state. */
export async function saveDropdownSettings() {
  try {
    const settings = getSettings();

    const apiKey = $("dropdown-api-key").value.trim();
    const model = $("dropdown-model").value;
    const language = $("dropdown-language").value;
    const eventTitleLanguage = $("dropdown-event-title-language").value;
    const theme = $("dropdown-theme").value;
    const fontSize = $("dropdown-font-size").value;
    const tldrMode = $("dropdown-tldr-mode").value;
    const showReply = $("dropdown-show-reply").value;
    const replyModelEl = $("dropdown-reply-model");
    const replyModel = replyModelEl ? replyModelEl.value : undefined;
    const autorun = $("dropdown-autorun").value;
    const autorunOption = $("dropdown-autorun-option").value;
    const devMode = $("dropdown-dev-mode").value;
    const devServer = $("dropdown-dev-server").value;
    const templates = collectPromptTemplatesFromForm();

    settings.apiKey = apiKey;
    settings.model = model;
    settings.defaultLanguage = language;
    settings.eventTitleLanguage = eventTitleLanguage;
    settings.theme = theme;
    settings.fontSize = fontSize;
    settings.tldrMode = tldrMode;
    settings.showReply = showReply;
    settings.replyModel = replyModel || "";
    settings.autorun = autorun;
    settings.autorunOption = autorunOption;
    settings.devMode = devMode;
    settings.devServer = devServer;
    settings.templates = {
      ...createBlankPromptTemplates(),
      ...templates,
    };

    saveSettingsToStore(settings);
    roamingStorage.setItem("theme", theme);
    await saveRoamingStoreAsync();

    applyCurrentTheme();
    applyFontSize(fontSize);
    updateReplyButtonVisibility(showReply === "true");
    updateDevBadges(devMode === "true");
    updateAuthenticationStatus();

    showNotification("All settings saved successfully");
    toggleSettingsView();
  } catch (error) {
    console.error("Error saving Outlook settings:", error);
    showNotification(`Failed to save settings: ${error.message}`, "error");
  }
}

/** Reset template fields to the saved session defaults. */
export function resetTemplates() {
  const currentSettings = getSettings();
  const sessionDefaults = getSavedTemplateDefaults();
  currentSettings.templates = {
    ...createBlankPromptTemplates(),
    ...sessionDefaults,
  };
  saveSettingsToStore(currentSettings);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving template reset:", error);
  });

  applyPromptTemplatesToForm(currentSettings.templates);
  showNotification(PROVIDER_MESSAGES.templatesReset);
}

/** Clear all template fields. */
export function clearTemplates() {
  const currentSettings = getSettings();
  currentSettings.templates = createBlankPromptTemplates();
  saveSettingsToStore(currentSettings);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving cleared templates:", error);
  });
  applyPromptTemplatesToForm(currentSettings.templates);
  showNotification(PROVIDER_MESSAGES.templatesCleared);
}

/** Load the built-in prompt defaults into the form (not yet saved). */
export function loadBuiltInTemplateDefaults() {
  const builtInDefaults = createDefaultPromptTemplates();
  applyPromptTemplatesToForm(builtInDefaults);
  showNotification("Built-in prompt defaults loaded into the form.");
}

/** Save the current form templates as the saved session defaults. */
export function saveCurrentTemplatesAsDefaults() {
  const templates = collectPromptTemplatesFromForm();
  const currentSettings = getSettings();
  currentSettings.templates = {
    ...createBlankPromptTemplates(),
    ...templates,
  };

  saveTemplateDefaults(templates);
  saveSettingsToStore(currentSettings);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving template defaults:", error);
  });
  showNotification(PROVIDER_MESSAGES.sessionDefaultsSaved);
}

/** Copy all templates from the form to the clipboard as markdown. */
export async function copyAllTemplatesToClipboard() {
  try {
    const templates = collectPromptTemplatesFromForm();
    const lines = [
      "# Michael Prompt Templates",
      "",
      ...Object.entries(templates).flatMap(([key, value]) => [
        `## ${key}`,
        "```",
        value || "",
        "```",
        "",
      ]),
    ];

    await navigator.clipboard.writeText(lines.join("\n"));
    showNotification("All prompt templates copied to clipboard.", "success");
  } catch (error) {
    console.error("Error copying prompt templates:", error);
    showNotification("Failed to copy prompt templates.", "error");
  }
}

/** Export the saved templates + general settings as a downloadable markdown file. */
export function exportTemplatesAsMarkdown() {
  try {
    const settings = getSettings();
    const templates = settings.templates || createBlankPromptTemplates();
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);

    let userInfo = "";
    try {
      if (
        typeof Office !== "undefined" &&
        Office.context.mailbox &&
        Office.context.mailbox.userProfile
      ) {
        const user = Office.context.mailbox.userProfile;
        userInfo = `\n\n*Exported by: ${user.displayName} (${user.emailAddress})*`;
      }
    } catch {
      console.log("User profile not available");
    }

    let markdownContent = `# Michael Prompt Templates\n\n`;
    markdownContent += `*Exported on: ${now.toLocaleString()}*${userInfo}\n\n`;
    markdownContent += `## General Settings\n\n`;
    markdownContent += `- **Provider**: Z.AI GLM Coding Plan\n`;
    markdownContent += `- **Authentication**: Outlook add-in saved setting\n`;
    markdownContent += `- **Model**: ${settings.model || "(empty)"}\n`;
    markdownContent += `- **Reply Model**: ${settings.replyModel || "(empty)"}\n`;
    markdownContent += `- **Default Language**: ${settings.defaultLanguage || DEFAULT_SETTINGS.defaultLanguage}\n`;
    markdownContent += `- **Event Title Language**: ${settings.eventTitleLanguage || DEFAULT_SETTINGS.eventTitleLanguage}\n\n`;
    markdownContent += `## Prompt Templates\n\n`;

    const exportedSections = [
      ["Summarize Template", templates.summarize],
      ["Translate Template", templates.translate],
      ["Translate & Summarize Template", templates.translateSummarize],
      ["Reply Template", templates.reply],
      ["Quick Translate Command Template", templates.commandTranslate],
      ["TL;DR Template", templates.tldrPrompt],
      ["Calendar Parse Template", templates.calendarParse],
      ["Calendar Check Template", templates.calendarCheck],
    ];

    exportedSections.forEach(([title, value]) => {
      markdownContent += `### ${title}\n\n\`\`\`\n${value || ""}\n\`\`\`\n\n`;
    });

    const blob = new Blob([markdownContent], { type: "text/markdown" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `michael-templates-${dateStr}.md`;

    document.body.appendChild(a);
    a.click();

    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);

    showNotification("Templates exported successfully", "success");
  } catch (error) {
    console.error("Error exporting templates:", error);
    showNotification("Failed to export templates", "error");
  }
}

/** Reset every saved Outlook add-in setting to defaults. */
export async function resetAllSettings() {
  try {
    const blankSettings = createBlankSettings();

    $("dropdown-api-key").value = "";
    $("dropdown-model").value = blankSettings.model;
    $("dropdown-language").value = blankSettings.defaultLanguage;
    $("dropdown-event-title-language").value = blankSettings.eventTitleLanguage;
    $("dropdown-theme").value = blankSettings.theme;
    $("dropdown-font-size").value = blankSettings.fontSize;
    $("dropdown-tldr-mode").value = blankSettings.tldrMode;
    $("dropdown-show-reply").value = blankSettings.showReply;
    $("dropdown-reply-model").value = blankSettings.replyModel;
    $("dropdown-autorun").value = blankSettings.autorun;
    $("dropdown-autorun-option").value = blankSettings.autorunOption;
    $("dropdown-dev-mode").value = blankSettings.devMode;
    $("dropdown-dev-server").value = blankSettings.devServer;
    $("dev-server-group").style.display = "none";

    applyPromptTemplatesToForm(blankSettings.templates);

    roamingStorage.removeItem(SETTINGS_KEY);
    roamingStorage.removeItem(TEMPLATE_DEFAULTS_KEY);
    roamingStorage.removeItem(MODEL_CATALOG_CACHE_KEY);
    roamingStorage.removeItem("theme");
    await saveRoamingStoreAsync();
    persistModelCatalog(getDefaultZaiModels());
    syncModelDropdowns();
    updateAuthenticationStatus();
    setText(
      "dropdown-model-status",
      "Saved settings cleared. Model catalog reset to fallback defaults."
    );

    applyCurrentTheme();
    applyFontSize(blankSettings.fontSize);
    updateReplyButtonVisibility(blankSettings.showReply === "true");
    updateDevBadges(false);
    syncSegmentedToggles();

    showNotification("All saved Outlook settings cleared", "success");
    initializeSettingsTabs();
  } catch (error) {
    console.error("Error resetting Outlook settings:", error);
    showNotification(`Failed to reset settings: ${error.message}`, "error");
  }
}

// --- Theme application (resolveTheme is the single fallback source) --------

/**
 * Apply the current theme to <body>, switching logos for contrast.
 * Resolution is delegated to theme.resolveTheme so the dark/light fallback
 * lives in exactly one place.
 */
export function applyCurrentTheme() {
  const savedTheme = getSettings().theme || DEFAULT_THEME_FALLBACK;
  const officeTheme =
    typeof Office !== "undefined" && Office.context ? Office.context.officeTheme : undefined;
  const resolved = resolveTheme(savedTheme, officeTheme);

  document.body.setAttribute("data-theme", resolved);
  document.body.classList.toggle("dark-theme", resolved === "dark");

  // ----- Logo switching (white on dark, black on light) -----
  const sideloadLogo = $("sideload-logo");
  const landingLogo = $("landing-logo-main");
  const brandLogo = $("brand-logo");
  const isDark = resolved === "dark";

  if (sideloadLogo) {
    sideloadLogo.src = isDark
      ? getAssetPath("meet-michael-white.png")
      : getAssetPath("meet-michael-black.png");
  }
  if (landingLogo) {
    landingLogo.src = isDark
      ? getAssetPath("meet-michael-white.png")
      : getAssetPath("meet-michael-black.png");
  }
  if (brandLogo) {
    brandLogo.src = isDark ? getAssetPath("michael-white.png") : getAssetPath("michael-black.png");
  }
}

/** Re-apply the theme when Office settings change (only if following Office). */
export function onSettingsChanged() {
  const savedTheme = getSettings().theme;
  if (savedTheme !== "light" && savedTheme !== "dark") {
    applyCurrentTheme();
  }
}
