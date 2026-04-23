/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Blob, console, document, localStorage, marked, navigator, Office, sessionStorage, setTimeout, URL, window */

import { fetchAvailableModels, generateText, getDefaultZaiModels } from "../shared/zai";
import {
  createBlankPromptTemplates,
  createBlankSettings,
  createDefaultPromptTemplates,
  DEFAULT_SETTINGS,
  PROVIDER_MESSAGES,
  TEMPLATE_KEYS,
} from "./prompt-templates";

const TYPES = {
  SUMMARIZE: 0,
  TRANSLATE: 1,
  TRANSLATE_SUMMARIZE: 2,
  REPLY: 3,
  CALENDAR: 4,
};

const MODEL_CATALOG_CACHE_KEY = "michael_zai_model_catalog";
const TEMPLATE_DEFAULTS_KEY = "michael_template_defaults";
const SETTINGS_KEY = "michael_settings";
let availableModelCatalog = getDefaultZaiModels();

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

function getAssetPath(fileName) {
  return `./assets/${fileName}`;
}

function getRoamingStore() {
  return Office?.context?.roamingSettings || null;
}

function saveRoamingStoreAsync() {
  const store = getRoamingStore();
  if (!store) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    store.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error?.message || "Failed to save Outlook add-in settings."));
        return;
      }

      resolve();
    });
  });
}

const roamingStorage = Object.freeze({
  getItem(key) {
    try {
      const store = getRoamingStore();
      const value = store ? store.get(key) : null;
      return typeof value === "string" ? value : null;
    } catch (error) {
      console.error(`Error reading Outlook add-in setting ${key}:`, error);
      return null;
    }
  },
  setItem(key, value) {
    try {
      const store = getRoamingStore();
      if (!store) {
        return;
      }

      store.set(key, value);
    } catch (error) {
      console.error(`Error writing Outlook add-in setting ${key}:`, error);
    }
  },
  removeItem(key) {
    try {
      const store = getRoamingStore();
      if (!store) {
        return;
      }

      store.remove(key);
    } catch (error) {
      console.error(`Error removing Outlook add-in setting ${key}:`, error);
    }
  },
});

/**
 * Migrate legacy browser storage keys into Outlook add-in settings.
 */
function migrateSettingsKeys() {
  try {
    const legacySettings = sessionStorage.getItem("my_sidekick_michael_settings");
    const currentSettings = roamingStorage.getItem(SETTINGS_KEY);
    if (!currentSettings && legacySettings) {
      roamingStorage.setItem(SETTINGS_KEY, legacySettings);
    }
    sessionStorage.removeItem("my_sidekick_michael_settings");

    const previousLocalSettings = localStorage.getItem("michael_settings");
    if (!currentSettings && previousLocalSettings) {
      roamingStorage.setItem(SETTINGS_KEY, previousLocalSettings);
    }

    const legacyTemplateDefaults = sessionStorage.getItem("michael_session_template_defaults");
    const currentTemplateDefaults = roamingStorage.getItem(TEMPLATE_DEFAULTS_KEY);
    if (!currentTemplateDefaults && legacyTemplateDefaults) {
      roamingStorage.setItem(TEMPLATE_DEFAULTS_KEY, legacyTemplateDefaults);
    }

    const legacyModelCatalog = sessionStorage.getItem(MODEL_CATALOG_CACHE_KEY);
    const currentModelCatalog = roamingStorage.getItem(MODEL_CATALOG_CACHE_KEY);
    if (!currentModelCatalog && legacyModelCatalog) {
      roamingStorage.setItem(MODEL_CATALOG_CACHE_KEY, legacyModelCatalog);
    }

    saveRoamingStoreAsync().catch((error) => {
      console.error("Error saving migrated Outlook add-in settings:", error);
    });
  } catch {
    // no-op
  }
}

/**
 * Toggle the visibility of the settings view and handle main content visibility.
 */
function toggleSettingsView() {
  const settingsView = document.getElementById("settings-view");
  const appBody = document.getElementById("app-body"); // Main content area

  if (settingsView && appBody) {
    const isSettingsVisible = settingsView.style.display === "block";
    if (isSettingsVisible) {
      settingsView.style.display = "none";
      appBody.style.display = "flex"; // Show main content (assuming flex is default)
    } else {
      settingsView.style.display = "block";
      appBody.style.display = "none"; // Hide main content
      // Optional: Load settings when view is opened
      loadDropdownSettings();
      refreshModelCatalog({ silent: true });
      // Activate the first tab by default if needed, handled by initTabs
    }
  }
}

/**
 * Initialize tab switching logic for the settings view.
 */
function initializeSettingsTabs() {
  const tabButtons = document.querySelectorAll(".settings-tabs .settings-tab-button");
  const tabContents = document.querySelectorAll(".settings-content .settings-tab-content");

  if (!tabButtons.length || !tabContents.length) return;

  tabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      // Get target tab content ID from button's data attribute
      const targetTabId = button.getAttribute("data-tab");

      // Deactivate all buttons and contents
      tabButtons.forEach((btn) => btn.classList.remove("active"));
      tabContents.forEach((content) => content.classList.remove("active"));

      // Activate the clicked button and corresponding content
      button.classList.add("active");
      const targetContent = document.getElementById(targetTabId);
      if (targetContent) {
        targetContent.classList.add("active");
      } else {
        console.error("Target tab content not found:", targetTabId);
      }
    });
  });

  // Ensure the first tab is active on initialization
  const firstTabButton = tabButtons[0];
  const firstTabContentId = firstTabButton?.getAttribute("data-tab");
  const firstTabContent = document.getElementById(firstTabContentId);

  tabButtons.forEach((btn) => btn.classList.remove("active"));
  tabContents.forEach((content) => content.classList.remove("active"));

  if (firstTabButton && firstTabContent) {
    firstTabButton.classList.add("active");
    firstTabContent.classList.add("active");
  }
}

function readJsonStorage(key, fallbackValue) {
  try {
    const rawValue = roamingStorage.getItem(key);
    return rawValue ? JSON.parse(rawValue) : fallbackValue;
  } catch (error) {
    console.error(`Error reading ${key}:`, error);
    return fallbackValue;
  }
}

function writeJsonStorage(key, value) {
  roamingStorage.setItem(key, JSON.stringify(value));
}

function getSettings() {
  return readJsonStorage(SETTINGS_KEY, createBlankSettings());
}

function saveSettingsToStore(settings) {
  writeJsonStorage(SETTINGS_KEY, settings);
}

function getSavedTemplateDefaults() {
  return readJsonStorage(TEMPLATE_DEFAULTS_KEY, createBlankPromptTemplates());
}

function saveTemplateDefaults(templates) {
  writeJsonStorage(TEMPLATE_DEFAULTS_KEY, templates);
}

function getPromptFieldValue(templateKey) {
  const fieldId = PROMPT_FIELD_MAP[templateKey];
  const field = fieldId ? document.getElementById(fieldId) : null;
  return field ? field.value : "";
}

function setPromptFieldValue(templateKey, value) {
  const fieldId = PROMPT_FIELD_MAP[templateKey];
  const field = fieldId ? document.getElementById(fieldId) : null;
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

function getCachedModelCatalog() {
  const cachedModels = readJsonStorage(MODEL_CATALOG_CACHE_KEY, []);
  return Array.isArray(cachedModels) && cachedModels.length ? cachedModels : getDefaultZaiModels();
}

function persistModelCatalog(models) {
  availableModelCatalog =
    Array.isArray(models) && models.length ? [...models] : getDefaultZaiModels();
  roamingStorage.setItem(MODEL_CATALOG_CACHE_KEY, JSON.stringify(availableModelCatalog));
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving model catalog:", error);
  });
}

function setSettingsStatus(elementId, message) {
  const element = document.getElementById(elementId);
  if (element) {
    element.textContent = message;
  }
}

function updateAuthenticationStatus() {
  if (getApiKey()) {
    setSettingsStatus(
      "dropdown-auth-status",
      "Authentication source: saved Outlook add-in settings."
    );
    return;
  }

  setSettingsStatus(
    "dropdown-auth-status",
    "Authentication source: empty. Enter the API key in this screen and save settings."
  );
}

function getMissingApiKeyMessage() {
  return PROVIDER_MESSAGES.missingApiKey;
}

function updateModelSelectOptions(selectId, models, selectedValue) {
  const select = document.getElementById(selectId);
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
  const models = availableModelCatalog.length ? availableModelCatalog : getCachedModelCatalog();

  updateModelSelectOptions("dropdown-model", models, settings.model || "");
  updateModelSelectOptions("dropdown-reply-model", models, settings.replyModel || "");
}

async function refreshModelCatalog(options = {}) {
  const silent = options.silent === true;
  const apiKey = getApiKey();

  updateAuthenticationStatus();

  if (!apiKey) {
    availableModelCatalog = getCachedModelCatalog();
    syncModelDropdowns();
    setSettingsStatus(
      "dropdown-model-status",
      "Using cached/fallback Z.AI models. Enter an API key in Settings > General to enable live refresh."
    );
    return availableModelCatalog;
  }

  setSettingsStatus("dropdown-model-status", "Refreshing available Z.AI models...");

  try {
    const liveModels = await fetchAvailableModels({ apiKey });
    persistModelCatalog(liveModels);
    syncModelDropdowns();
    setSettingsStatus("dropdown-model-status", `Loaded ${liveModels.length} models from Z.AI.`);
    return availableModelCatalog;
  } catch (error) {
    console.error("Error refreshing Z.AI model catalog:", error);
    availableModelCatalog = getCachedModelCatalog();
    syncModelDropdowns();
    setSettingsStatus(
      "dropdown-model-status",
      `Live refresh failed. Using cached/fallback models (${error.message}).`
    );
    if (!silent) {
      showNotification(`Model refresh failed: ${error.message}`, "warning");
    }
    return availableModelCatalog;
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    migrateSettingsKeys();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Initialize the application
    initializeApp();

    // Check if autorun is enabled and get the selected option
    let autorunEnabled = false;
    let selectedOption = null;
    try {
      const settings = getSettings();
      autorunEnabled = settings.autorun === "true";
      selectedOption = settings.autorunOption;
    } catch (error) {
      console.error("Error getting autorun settings:", error);
    }

    // If autorun is enabled and an option is selected, execute it
    if (autorunEnabled && selectedOption) {
      switch (selectedOption) {
        case "summarize":
          summarizeEmail();
          break;
        case "translate":
          translateEmail();
          break;
        case "translateAndSummarize":
          translateAndSummarizeEmail();
          break;
        case "reply":
          generateReply();
          break;
      }
    }

    // Add event listeners for the application buttons
    document.getElementById("summarize").addEventListener("click", summarizeEmail);
    document.getElementById("translate").addEventListener("click", translateEmail);
    document
      .getElementById("translate-summarize")
      .addEventListener("click", translateAndSummarizeEmail);
    document.getElementById("calendar-event").addEventListener("click", handleCalendarEvent);
    document.getElementById("settings-toggle").addEventListener("click", toggleSettingsView); // Updated listener
    document.getElementById("close-settings-view").addEventListener("click", toggleSettingsView); // Listener for new close button
    document
      .getElementById("dropdown-save-settings")
      .addEventListener("click", saveDropdownSettings);
    document.getElementById("dropdown-reset-all").addEventListener("click", resetAllSettings);
    // Note: Template specific buttons are now inside the template tab HTML
    const resetTemplatesBtn = document.getElementById("dropdown-reset-templates");
    if (resetTemplatesBtn) {
      resetTemplatesBtn.addEventListener("click", resetTemplates);
    }
    const loadBuiltInsBtn = document.getElementById("dropdown-load-builtins");
    if (loadBuiltInsBtn) {
      loadBuiltInsBtn.addEventListener("click", loadBuiltInTemplateDefaults);
    }
    const saveTemplateDefaultsBtn = document.getElementById("dropdown-save-template-defaults");
    if (saveTemplateDefaultsBtn) {
      saveTemplateDefaultsBtn.addEventListener("click", saveCurrentTemplatesAsDefaults);
    }
    const clearTemplatesBtn = document.getElementById("dropdown-clear-templates");
    if (clearTemplatesBtn) {
      clearTemplatesBtn.addEventListener("click", clearTemplates);
    }
    const copyTemplatesBtn = document.getElementById("dropdown-copy-templates");
    if (copyTemplatesBtn) {
      copyTemplatesBtn.addEventListener("click", copyAllTemplatesToClipboard);
    }
    const exportMarkdownBtn = document.getElementById("dropdown-export-markdown");
    if (exportMarkdownBtn) {
      exportMarkdownBtn.addEventListener("click", exportTemplatesAsMarkdown);
    }
    const refreshModelsBtn = document.getElementById("dropdown-refresh-models");
    if (refreshModelsBtn) {
      refreshModelsBtn.addEventListener("click", () => {
        refreshModelCatalog();
      });
    }
    const apiKeyInput = document.getElementById("dropdown-api-key");
    if (apiKeyInput) {
      apiKeyInput.addEventListener("input", updateAuthenticationStatus);
    }

    // Add dev mode toggle listener
    document.getElementById("dropdown-dev-mode").addEventListener("change", function () {
      const devServerGroup = document.getElementById("dev-server-group");
      devServerGroup.style.display = this.value === "true" ? "block" : "none";
    });

    // Expand button listener
    document.getElementById("expand-content").addEventListener("click", expandContent);
    // Copy reply listener if present
    const copyReplyBtn = document.getElementById("copy-reply");
    if (copyReplyBtn) copyReplyBtn.addEventListener("click", copyReply);

    // Copy buttons listeners
    document.getElementById("copy-result").addEventListener("click", copyResult);
    document.getElementById("generate-reply").addEventListener("click", generateReply);

    // Load saved settings if any
    loadDropdownSettings();
    updateAuthenticationStatus();
    refreshModelCatalog({ silent: true });

    // Apply current theme
    applyCurrentTheme();

    // Register for theme change events
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.SettingsChanged, onSettingsChanged);
    }

    // Add event handler for email selection
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      function () {
        // Check if autorun is enabled and get the selected option
        let autorunEnabled = false;
        let selectedOption = null;
        try {
          const settings = getSettings();
          autorunEnabled = settings.autorun === "true";
          selectedOption = settings.autorunOption;
        } catch (error) {
          console.error("Error getting autorun settings:", error);
        }

        // If autorun is enabled and an option is selected, execute it
        if (autorunEnabled && selectedOption) {
          switch (selectedOption) {
            case "summarize":
              summarizeEmail();
              break;
            case "translate":
              translateEmail();
              break;
            case "translateAndSummarize":
              translateAndSummarizeEmail();
              break;
            case "reply":
              generateReply();
              break;
          }
        }
      },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          result.value.addHandlerAsync(
            Office.EventType.ItemChanged,
            function () {
              // Check if autorun is enabled and get the selected option
              let autorunEnabled = false;
              let selectedOption = null;
              try {
                const settings = getSettings();
                autorunEnabled = settings.autorun === "true";
                selectedOption = settings.autorunOption;
              } catch (error) {
                console.error("Error getting autorun settings:", error);
              }

              // If autorun is enabled and an option is selected, execute it
              if (autorunEnabled && selectedOption) {
                switch (selectedOption) {
                  case "summarize":
                    summarizeEmail();
                    break;
                  case "translate":
                    translateEmail();
                    break;
                  case "translateAndSummarize":
                    translateAndSummarizeEmail();
                    break;
                  case "reply":
                    generateReply();
                    break;
                }
              }
            },
            function (result) {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                result.value.register();
              }
            }
          );
        }
      }
    );

    // Update calendar button state when email changes
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      function () {
        updateCalendarButtonState();
      },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          result.value.addHandlerAsync(Office.EventType.ItemChanged, function () {
            updateCalendarButtonState();
          });
        }
      }
    );

    // Initial calendar button state update
    updateCalendarButtonState();

    // Initialize Settings Tabs
    initializeSettingsTabs();
  }
});

/**
 * Apply the current theme based on user preference or Office theme
 */
function applyCurrentTheme() {
  // Default to dark theme instead of system
  const savedTheme = getSettings().theme || "dark";

  if (savedTheme === "light") {
    document.body.setAttribute("data-theme", "light");
    document.body.classList.remove("dark-theme");
  } else if (savedTheme === "dark") {
    document.body.setAttribute("data-theme", "dark");
    document.body.classList.add("dark-theme");
  } else {
    // Use Office theme
    if (Office.context.officeTheme) {
      const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
      // Only call isDarkTheme if bodyBackgroundColor exists
      if (bodyBackgroundColor) {
        if (isDarkTheme(bodyBackgroundColor)) {
          document.body.setAttribute("data-theme", "dark");
          document.body.classList.add("dark-theme");
        } else {
          document.body.setAttribute("data-theme", "light");
          document.body.classList.remove("dark-theme");
        }
      } else {
        // Default to dark theme if no color information
        document.body.setAttribute("data-theme", "dark");
        document.body.classList.add("dark-theme");
      }
    }
  }

  // ----- Logo Switching Logic Start -----
  const sideloadLogo = document.getElementById("sideload-logo");
  const landingLogo = document.getElementById("landing-logo-main");
  const brandLogo = document.getElementById("brand-logo"); // Get new brand logo
  const currentThemeIsDark = document.body.classList.contains("dark-theme");

  // Set sideload logo (White on Dark, Black on Light)
  if (sideloadLogo) {
    sideloadLogo.src = currentThemeIsDark
      ? getAssetPath("meet-michael-white.png")
      : getAssetPath("meet-michael-black.png");
  }

  // Set landing page logo (White on Dark, Black on Light - Corrected)
  if (landingLogo) {
    landingLogo.src = currentThemeIsDark
      ? getAssetPath("meet-michael-white.png")
      : getAssetPath("meet-michael-black.png");
  }

  // Set brand logo (White on Dark, Black on Light)
  if (brandLogo) {
    brandLogo.src = currentThemeIsDark
      ? getAssetPath("michael-white.png")
      : getAssetPath("michael-black.png");
  }
  // ----- Logo Switching Logic End -----
}

/**
 * Determine if a color is dark by converting it to RGB and calculating perceived brightness
 * @param {string} color - Hex color code
 * @returns {boolean} - True if the color is dark
 */
function isDarkTheme(color) {
  // If color is undefined or not a string, default to light theme
  if (!color || typeof color !== "string") {
    return false;
  }

  try {
    // Convert hex to RGB
    color = color.replace("#", "");
    const r = parseInt(color.substr(0, 2), 16);
    const g = parseInt(color.substr(2, 2), 16);
    const b = parseInt(color.substr(4, 2), 16);

    // Check if we got valid RGB values
    if (isNaN(r) || isNaN(g) || isNaN(b)) {
      return false;
    }

    // Calculate perceived brightness using the formula: (0.299*R + 0.587*G + 0.114*B)
    const brightness = 0.299 * r + 0.587 * g + 0.114 * b;

    // If brightness is less than 128, consider it dark
    return brightness < 128;
  } catch (error) {
    console.error("Error calculating theme brightness:", error);
    return false;
  }
}

/**
 * Handle Office theme change event
 */
function onSettingsChanged() {
  const savedTheme = getSettings().theme || "system";
  if (savedTheme === "system") {
    applyCurrentTheme();
  }
}

/**
 * Show notification message
 */
function showNotification(message, type = "info") {
  // Remove any existing notification
  const existingNotification = document.getElementById("notification");
  if (existingNotification) {
    existingNotification.remove();
  }

  // Create new notification element
  const notification = document.createElement("div");
  notification.id = "notification";
  notification.className = `notification ${type}`;
  notification.textContent = message;

  // Add to document
  document.body.appendChild(notification);

  // Force reflow to ensure animation plays
  notification.offsetHeight;

  // Add slide-in animation
  notification.style.animation = "slideInFromTop 0.3s ease-out";

  // Set timeout to remove notification
  setTimeout(() => {
    // Add slide-out animation
    notification.style.animation = "slideOutToTop 0.3s ease-out";

    // Remove element after animation
    setTimeout(() => {
      notification.remove();
    }, 300);
  }, 3000);
}

/**
 * Save settings from the dropdown menu
 */
async function saveDropdownSettings() {
  try {
    const settings = getSettings();

    // Update with values from dropdown form
    const apiKey = document.getElementById("dropdown-api-key").value.trim();
    const model = document.getElementById("dropdown-model").value;
    const language = document.getElementById("dropdown-language").value;
    const eventTitleLanguage = document.getElementById("dropdown-event-title-language").value;
    const theme = document.getElementById("dropdown-theme").value;
    const fontSize = document.getElementById("dropdown-font-size").value;
    const tldrMode = document.getElementById("dropdown-tldr-mode").value;
    const showReply = document.getElementById("dropdown-show-reply").value;
    const replyModel = document.getElementById("dropdown-reply-model")
      ? document.getElementById("dropdown-reply-model").value
      : undefined;
    const autorun = document.getElementById("dropdown-autorun").value;
    const autorunOption = document.getElementById("dropdown-autorun-option").value;
    const devMode = document.getElementById("dropdown-dev-mode").value;
    const devServer = document.getElementById("dropdown-dev-server").value;
    const templates = collectPromptTemplatesFromForm();

    // Update settings object
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

/**
 * Reset template fields to defaults
 */
function resetTemplates() {
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

function clearTemplates() {
  const currentSettings = getSettings();
  currentSettings.templates = createBlankPromptTemplates();
  saveSettingsToStore(currentSettings);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving cleared templates:", error);
  });
  applyPromptTemplatesToForm(currentSettings.templates);
  showNotification(PROVIDER_MESSAGES.templatesCleared);
}

function loadBuiltInTemplateDefaults() {
  const builtInDefaults = createDefaultPromptTemplates();
  applyPromptTemplatesToForm(builtInDefaults);
  showNotification("Built-in prompt defaults loaded into the form.");
}

function saveCurrentTemplatesAsDefaults() {
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

async function copyAllTemplatesToClipboard() {
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

/**
 * Load saved settings to the dropdown form fields
 */
function loadDropdownSettings() {
  try {
    const settings = getSettings();
    const sessionDefaults = getSavedTemplateDefaults();

    document.getElementById("dropdown-api-key").value = settings.apiKey || "";
    if (settings.defaultLanguage)
      document.getElementById("dropdown-language").value = settings.defaultLanguage;
    if (settings.eventTitleLanguage)
      document.getElementById("dropdown-event-title-language").value = settings.eventTitleLanguage;
    if (settings.theme) document.getElementById("dropdown-theme").value = settings.theme;
    if (settings.fontSize) document.getElementById("dropdown-font-size").value = settings.fontSize;
    if (settings.tldrMode) document.getElementById("dropdown-tldr-mode").value = settings.tldrMode;
    if (settings.showReply)
      document.getElementById("dropdown-show-reply").value = settings.showReply;
    if (settings.devServer !== undefined)
      document.getElementById("dropdown-dev-server").value = settings.devServer;
    if (settings.autorun) document.getElementById("dropdown-autorun").value = settings.autorun;
    if (settings.autorunOption)
      document.getElementById("dropdown-autorun-option").value = settings.autorunOption;
    if (settings.devMode) document.getElementById("dropdown-dev-mode").value = settings.devMode;

    const devServerGroup = document.getElementById("dev-server-group");
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

    availableModelCatalog = getCachedModelCatalog();
    syncModelDropdowns();
    updateAuthenticationStatus();
  } catch (error) {
    console.error("Error loading dropdown settings:", error);
    applyPromptTemplatesToForm(createBlankPromptTemplates());
  }
}

/**
 * Apply the selected font size to result content
 * @param {string} size - The font size to apply (small, medium, large)
 */
function applyFontSize(size) {
  document.documentElement.setAttribute("data-font-size", size || "medium");
}

// Get email content
async function getEmailContent() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email content"));
      }
    });
  });
}

// Generate content using Z.AI GLM Coding Plan
async function generateContent(prompt, apiKey, modelOverride = null, isTldr = false) {
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
    // Hide loading spinner only if this is not a TL;DR request
    if (!isTldr) {
      document.getElementById("loading").style.display = "none";
    }
  }
}

// Generate TL;DR content
async function generateTldrContent(prompt, apiKey, language = "Korean", modelOverride = null) {
  const subject = Office.context.mailbox.item.subject;
  const emailContent = await getEmailContent();

  const tldrPrompt = requireTemplate("tldrPrompt", "TL;DR")
    .replace("{subject}", subject)
    .replace("{content}", emailContent)
    .replace("{language}", language);

  const content = await generateContent(tldrPrompt, apiKey, modelOverride, true);
  return content;
}

// Get language display text
function getLanguageText(languageCode) {
  switch (languageCode) {
    case "es":
      return "Spanish";
    case "fr":
      return "French";
    case "de":
      return "German";
    case "it":
      return "Italian";
    case "ja":
      return "Japanese";
    case "ko":
      return "Korean";
    case "zh_cn":
      return "Chinese";
    default:
      return "English";
  }
}

// Get the configured Z.AI API key from Outlook add-in settings
function getApiKey() {
  const input = document.getElementById("dropdown-api-key");
  if (input && typeof input.value === "string" && input.value.trim()) {
    return input.value.trim();
  }

  const settings = getSettings();
  return typeof settings.apiKey === "string" ? settings.apiKey.trim() : "";
}

// Get language from settings
function getLanguage() {
  const settings = getSettings();
  return settings.defaultLanguage || DEFAULT_SETTINGS.defaultLanguage;
}

function getTemplateValue(templateKey) {
  const settings = getSettings();
  return settings.templates && typeof settings.templates[templateKey] === "string"
    ? settings.templates[templateKey].trim()
    : "";
}

function requireTemplate(templateKey, label) {
  const template = getTemplateValue(templateKey);
  if (!template) {
    throw new Error(`${label} prompt is empty. Configure it in Settings > Templates.`);
  }

  return template;
}

function requireModel(settingKey, label) {
  const settings = getSettings();
  const model = typeof settings[settingKey] === "string" ? settings[settingKey].trim() : "";
  if (!model) {
    throw new Error(`${label} is empty. Select a model in Settings > General.`);
  }

  return model;
}

// Summarize email
async function summarizeEmail() {
  const apiKey = getApiKey();

  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView(); // Open settings dropdown to prompt for API key
    return;
  }

  // Show loading UI
  showLoading("Summarizing email...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;

    const template = requireTemplate("summarize", "Summarize");

    // Replace placeholders in template
    const prompt = template.replace("{subject}", subject).replace("{content}", emailContent);

    // Check for TL;DR mode
    let tldrMode = true;
    try {
      const settings = getSettings();
      if (settings.tldrMode) {
        tldrMode = settings.tldrMode === "true";
      }
    } catch (error) {
      console.error("Error getting TLDR mode setting:", error);
    }

    if (tldrMode) {
      // Generate TL;DR first
      const tldrContent = await generateTldrContent(prompt, apiKey, getLanguageText(getLanguage()));
      hideLoading();
      showResults(tldrContent, TYPES.SUMMARIZE);

      // Then generate full content in the background
      const fullContent = await generateContent(prompt, apiKey);

      // display notification of full content
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      // Generate full content only
      const summary = await generateContent(prompt, apiKey);
      showResults(summary, TYPES.SUMMARIZE);
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  }
}

// Translate email
async function translateEmail() {
  const apiKey = getApiKey();
  const language = getLanguage();

  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView(); // Open settings dropdown to prompt for API key
    return;
  }

  // Show loading UI
  showLoading("Translating to " + getLanguageText(language) + "...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;

    const template = requireTemplate("translate", "Translate");

    // Replace placeholders in template
    const prompt = template.replace("{subject}", subject).replace("{content}", emailContent);

    // Check for TL;DR mode
    let tldrMode = true;
    try {
      const settings = getSettings();
      if (settings.tldrMode) {
        tldrMode = settings.tldrMode === "true";
      }
    } catch (error) {
      console.error("Error getting TLDR mode setting:", error);
    }

    if (tldrMode) {
      // Generate TL;DR first
      const tldrContent = await generateTldrContent(prompt, apiKey, getLanguageText(language));
      hideLoading();
      showResults(tldrContent, TYPES.TRANSLATE);

      // Then generate full content in the background
      const fullContent = await generateContent(prompt, apiKey);

      // display notification of full content
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      // Generate full content only
      const translation = await generateContent(prompt, apiKey);
      showResults(translation, TYPES.TRANSLATE);
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  }
}

/**
 * Update the expand button text and style based on the full content display state
 */
function updateExpandButton(isFullContentVisible) {
  const expandButton = document.getElementById("expand-content");

  if (expandButton) {
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
}

// Copy result to clipboard
function copyResult() {
  // Check if this is a reply (both TLDR and full content need to be combined)
  let resultContent = "";

  // Get the TLDR content (might contain the subject)
  const tldrContent = document.getElementById("tldr-content").innerText;

  // Get the full content (might contain the body)
  const fullContent = document.getElementById("result-content").innerText;

  // Check if this is a reply format
  const isReply = tldrContent.includes("Subject:") || fullContent.includes("Subject:");

  if (isReply) {
    // If it's a reply, try to extract and format subject and body properly
    let subject = "";
    let body = "";

    // Try to get subject from headings
    const headingMatch = /^(?:Subject:\s*)?(.+?)(?:\n|$)/i.exec(tldrContent);
    if (headingMatch) {
      subject = headingMatch[1].trim();
    }

    // Get the body text (prioritize full content if visible, otherwise use TLDR minus subject)
    if (document.getElementById("full-content-container").style.display !== "none") {
      body = fullContent;
    } else {
      // Attempt to remove subject line from TLDR if present
      if (headingMatch) {
        const lines = tldrContent.split("\n");
        body = lines.slice(1).join("\n").trim();
      } else {
        body = tldrContent;
      }
    }

    // Format as email reply
    resultContent = `Subject: ${subject}\n\n${body}`;
  } else {
    // For normal content, get what's visible (either TLDR or full content)
    if (document.getElementById("full-content-container").style.display !== "none") {
      resultContent = fullContent;
    } else {
      resultContent = tldrContent;
    }
  }

  navigator.clipboard
    .writeText(resultContent)
    .then(() => {
      const copyStatus = document.getElementById("copy-status");
      copyStatus.textContent = "Copied!";
      setTimeout(() => {
        copyStatus.textContent = "";
      }, 2000);
    })
    .catch((err) => {
      console.error("Could not copy text: ", err);
      showNotification("Failed to copy to clipboard", "error");
    });
}

// Translate and Summarize email
async function translateAndSummarizeEmail() {
  const apiKey = getApiKey();

  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView(); // Open settings dropdown to prompt for API key
    return;
  }

  // Show loading UI
  showLoading("Translating and summarizing...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;
    let language = "English";

    let template = requireTemplate("translateSummarize", "Translate & Summarize");
    try {
      const settings = getSettings();
      if (settings.defaultLanguage) language = getLanguageText(settings.defaultLanguage);
    } catch (error) {
      console.error("Error getting template:", error);
    }

    // Replace placeholders in template
    const prompt = template
      .replace("{subject}", subject)
      .replace("{content}", emailContent)
      .replace("{language}", language);

    // Check for TL;DR mode
    let tldrMode = true;
    try {
      const settings = getSettings();
      if (settings.tldrMode) {
        tldrMode = settings.tldrMode === "true";
      }
    } catch (error) {
      console.error("Error getting TLDR mode setting:", error);
    }

    if (tldrMode) {
      // Generate TL;DR first
      const tldrContent = await generateTldrContent(prompt, apiKey);
      hideLoading();
      showResults(tldrContent, TYPES.TRANSLATE_SUMMARIZE);

      // Then generate full content in the background
      const fullContent = await generateContent(prompt, apiKey);

      // display notification of full content
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      // Generate full content only
      const result = await generateContent(prompt, apiKey);
      showResults(result, TYPES.TRANSLATE_SUMMARIZE);
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  }
}

/**
 * Show loading indicator with message
 */
function showLoading(message = "Loading...") {
  // Show loading section
  const loadingSection = document.getElementById("loading");
  loadingSection.style.display = "block";

  // Update loading message
  const loadingMessage = document.getElementById("loading-message");
  loadingMessage.textContent = message;

  // Hide other sections
  document.getElementById("landing-screen").style.display = "none";
  document.getElementById("result-section").style.display = "none";
}

/**
 * Hide loading indicator
 */
function hideLoading() {
  const loadingSection = document.getElementById("loading");
  if (loadingSection) {
    loadingSection.style.display = "none";
  }
}

// Function to show the results
function showResults(content, type) {
  // Reset the full result content
  document.getElementById("result-content").innerHTML = "";
  // Reset the tldr content
  document.getElementById("tldr-content").innerHTML = "";

  // Reset the expand button
  const expandButton = document.getElementById("expand-content");
  if (expandButton) {
    expandButton.disabled = true;
    expandButton.classList.add("ms-Button--disabled");
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    expandButton.classList.remove("ms-Button--primary");
  }

  // Hide loading section
  document.getElementById("loading").style.display = "none";

  // Show result section
  document.getElementById("result-section").style.display = "block";

  // Hide landing screen
  document.getElementById("landing-screen").style.display = "none";

  // Show the app body
  document.getElementById("app-body").style.display = "block";

  // Check for TL;DR mode
  let tldrMode = true;
  try {
    const settings = getSettings();
    if (settings.tldrMode) {
      tldrMode = settings.tldrMode === "true";
    }
  } catch (error) {
    console.error("Error getting TLDR mode setting:", error);
  }

  // Update content based on TL;DR mode
  if (tldrMode) {
    // For TL;DR mode, show the quick summary first
    document.getElementById("tldr-content").innerHTML = marked.parse(content);

    // // Show loading spinner in full content section
    // const fullContentContainer = document.getElementById("full-content-container");
    // fullContentContainer.style.display = "block";
    // fullContentContainer.innerHTML = `
    //     <div class="loading-container" style="display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 20px;">
    //         <div class="spinner"></div>
    //         <p style="margin-top: 10px; color: var(--text-secondary);">Generating full content...</p>
    //     </div>
    // `;

    // Disable expand button and set to loading state
    const expandButton = document.getElementById("expand-content");
    if (expandButton) {
      expandButton.disabled = true;
      expandButton.classList.add("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Loading Full Content...</span>';
    }
  } else {
    // For non-TL;DR mode, show the full content
    document.getElementById("tldr-content").innerHTML = marked.parse(content);

    // Ensure result-content element exists
    let resultContent = document.getElementById("result-content");
    if (!resultContent) {
      resultContent = document.createElement("div");
      resultContent.id = "result-content";
      document.getElementById("full-content-container").appendChild(resultContent);
    }
    resultContent.innerHTML = marked.parse(content);
    document.getElementById("full-content-container").style.display = "block";

    // Enable expand button and set to normal state
    const expandButton = document.getElementById("expand-content");
    if (expandButton) {
      expandButton.disabled = false;
      expandButton.classList.remove("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
      expandButton.classList.add("ms-Button--primary");
    }
  }

  // Show/hide copy reply button based on type
  const copyReplyButton = document.getElementById("copy-reply");
  if (copyReplyButton) {
    copyReplyButton.style.display = type === TYPES.REPLY ? "inline-block" : "none";
  }

  // Show/hide copy result button based on type
  const copyResultButton = document.getElementById("copy-result");
  if (copyResultButton) {
    copyResultButton.style.display = type === TYPES.REPLY ? "none" : "inline-block";
  }

  // Show/hide generate reply button based on type and settings
  const generateReplyButton = document.getElementById("generate-reply");
  if (generateReplyButton) {
    const showReply = getSettings().showReply === "true";
    generateReplyButton.style.display =
      type === TYPES.REPLY ? "none" : showReply ? "inline-block" : "none";
  }

  // Apply font size from settings
  try {
    const settings = getSettings();
    if (settings.fontSize) {
      applyFontSize(settings.fontSize);
    }
    // Update reply button visibility
    if (settings.showReply) {
      updateReplyButtonVisibility(settings.showReply === "true");
    }
  } catch (error) {
    console.error("Error applying font size:", error);
  }

  // Scroll to top of result content if elements exist
  const resultContent = document.getElementById("result-content");
  const tldrContent = document.getElementById("tldr-content");
  if (resultContent) {
    resultContent.scrollTop = 0;
  }
  if (tldrContent) {
    tldrContent.scrollTop = 0;
  }
}

function updateResults(content) {
  // change the button status to show full content
  const expandButton = document.getElementById("expand-content");
  if (expandButton) {
    expandButton.disabled = false;
    expandButton.classList.remove("ms-Button--disabled");
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
  }

  // Remove the loading container
  const loadingContainer = document.getElementById("loading-container");
  if (loadingContainer) {
    loadingContainer.remove();
  }

  // Redo the result content
  const resultContent = document.getElementById("result-content");
  if (resultContent) {
    resultContent.innerHTML = marked.parse(content);
  }
}

/**
 * Update reply button visibility based on settings
 */
function updateReplyButtonVisibility(show) {
  const replyButton = document.getElementById("generate-reply");
  if (replyButton) {
    replyButton.style.display = show ? "inline-block" : "none";
  }
}

/**
 * Expand the full content when the expand button is clicked
 */
function expandContent() {
  const expandButton = document.getElementById("expand-content");
  if (expandButton.disabled) {
    return; // Don't do anything if button is disabled
  }

  const fullContentContainer = document.getElementById("full-content-container");

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

// Format a reply with clear subject and body sections
function formatReplyOutput(replyText) {
  // Extract subject and body
  let subject = "";
  let body = "";

  // Check for a SUBJECT: line first
  const subjectMatch = replyText.match(/^(?:SUBJECT:|Subject:)\s*(.+?)(?:\n|$)/m);
  if (subjectMatch) {
    subject = subjectMatch[1].trim();

    // Remove the subject line from the text to get the body
    body = replyText.replace(/^(?:SUBJECT:|Subject:)\s*.+?\n+/m, "").trim();
  } else {
    // If no explicit subject marker, check for first line as subject
    const lines = replyText.trim().split("\n");
    if (lines.length > 0) {
      subject = lines[0].trim();
      body = lines.slice(1).join("\n").trim();
    } else {
      // Fallback if no clear structure
      subject = "Re: Your email";
      body = replyText.trim();
    }
  }

  // Create formatted HTML
  const formattedHtml = `
    <div class="reply-container">
      <div class="reply-subject">
        <span class="reply-label">Subject:</span>${subject}
      </div>
      <div class="reply-body">${body}</div>
    </div>
  `;

  return {
    html: formattedHtml,
    subject: subject,
    body: body,
    raw: `Subject: ${subject}\n\n${body}`,
  };
}

// Generate a reply based on the current content
async function generateReply() {
  const apiKey = getApiKey();

  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView();
    return;
  }

  // Show loading UI
  showLoading("Generating reply...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;

    const template = requireTemplate("reply", "Reply");

    // Get language (assuming you still want language for reply)
    const language = getLanguage();

    // Replace placeholders in template
    const prompt = template
      .replace("{subject}", subject)
      .replace("{content}", emailContent)
      .replace("{language}", getLanguageText(language)); // Ensure language name is used if placeholder exists

    // Get reply model from settings
    let replyModelOverride = null;
    try {
      replyModelOverride = requireModel("replyModel", "Reply model");
    } catch (error) {
      console.error("Error getting reply model:", error);
    }

    const result = await generateContent(prompt, apiKey, replyModelOverride);
    let formattedReply = formatReplyOutput(result);

    // Display in TLDR and full content sections
    document.getElementById("tldr-content").innerHTML = formattedReply.html;
    document.getElementById("result-content").innerHTML = formattedReply.html;

    // Show the result section
    document.getElementById("result-section").style.display = "block";

    // Show the copy reply button and hide the regular copy button
    document.getElementById("copy-reply").style.display = "inline-block";
    document.getElementById("copy-result").style.display = "none";
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    hideLoading();
  }
}

// Copy the reply to clipboard
function copyReply() {
  // Get the formatted reply content
  const tldrContent = document.getElementById("tldr-content").innerText;

  // Extract subject and body using regex
  const subjectMatch = tldrContent.match(/Subject:\s*(.+?)(?:\n|$)/i);
  const subject = subjectMatch ? subjectMatch[1].trim() : "";

  // Get the body (everything after the subject)
  const bodyStart = tldrContent.indexOf(subject) + subject.length;
  const body = tldrContent.substring(bodyStart).trim();

  // Format as email reply
  const replyContent = `Subject: ${subject}\n\n${body}`;

  navigator.clipboard
    .writeText(replyContent)
    .then(() => {
      const copyStatus = document.getElementById("copy-reply-status");
      copyStatus.textContent = "Copied!";
      setTimeout(() => {
        copyStatus.textContent = "";
      }, 2000);
    })
    .catch((err) => {
      console.error("Could not copy reply: ", err);
      showNotification("Failed to copy reply", "error");
    });
}

/**
 * Initialize the application
 */
async function initializeApp() {
  // DOM 요소들을 먼저 찾아서 변수에 저장
  const settingsSection = document.getElementById("settings-section");
  const resultSection = document.getElementById("result-section");
  const tldrSection = document.getElementById("tldr-section");

  // 요소가 존재하는지 확인 후 스타일 적용
  if (settingsSection) {
    settingsSection.style.display = "none";
  }

  if (resultSection) {
    resultSection.style.display = "none";
  }

  if (tldrSection) {
    tldrSection.style.display = "none";
  }

  // 설정 로드
  const settings = loadSettings();
  if (settings) {
    // 테마 설정
    if (settings.theme) {
      setTheme(settings.theme);
    }

    // 폰트 크기 설정
    if (settings.fontSize) {
      setFontSize(settings.fontSize);
    }
  }

  // API 키 확인
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "warning");
  }
}

/**
 * Load settings from Outlook add-in settings
 */
function loadSettings() {
  try {
    return getSettings();
  } catch (error) {
    console.error("Error loading settings:", error);
    return createBlankSettings();
  }
}

/**
 * Set theme based on selection
 */
function setTheme(theme) {
  const root = document.documentElement;

  if (theme === "system") {
    // Use system preference
    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) {
      root.setAttribute("data-theme", "dark");
    } else {
      root.setAttribute("data-theme", "light");
    }
  } else {
    // Use explicit theme
    root.setAttribute("data-theme", theme);
  }
}

/**
 * Set font size for results
 */
function setFontSize(size) {
  const root = document.documentElement;
  root.style.setProperty("--result-font-size", getFontSizeValue(size));
}

/**
 * Get font size value based on size name
 */
function getFontSizeValue(size) {
  switch (size) {
    case "small":
      return "0.875rem";
    case "medium":
      return "1rem";
    case "large":
      return "1.125rem";
    default:
      return "1rem";
  }
}

/**
 * Reset all settings stored in Outlook add-in settings.
 */
async function resetAllSettings() {
  try {
    const blankSettings = createBlankSettings();

    document.getElementById("dropdown-api-key").value = "";
    document.getElementById("dropdown-model").value = blankSettings.model;
    document.getElementById("dropdown-language").value = blankSettings.defaultLanguage;
    document.getElementById("dropdown-event-title-language").value =
      blankSettings.eventTitleLanguage;
    document.getElementById("dropdown-theme").value = blankSettings.theme;
    document.getElementById("dropdown-font-size").value = blankSettings.fontSize;
    document.getElementById("dropdown-tldr-mode").value = blankSettings.tldrMode;
    document.getElementById("dropdown-show-reply").value = blankSettings.showReply;
    document.getElementById("dropdown-reply-model").value = blankSettings.replyModel;
    document.getElementById("dropdown-autorun").value = blankSettings.autorun;
    document.getElementById("dropdown-autorun-option").value = blankSettings.autorunOption;
    document.getElementById("dropdown-dev-mode").value = blankSettings.devMode;
    document.getElementById("dropdown-dev-server").value = blankSettings.devServer;
    document.getElementById("dev-server-group").style.display = "none";

    applyPromptTemplatesToForm(blankSettings.templates);

    roamingStorage.removeItem(SETTINGS_KEY);
    roamingStorage.removeItem(TEMPLATE_DEFAULTS_KEY);
    roamingStorage.removeItem(MODEL_CATALOG_CACHE_KEY);
    roamingStorage.removeItem("theme");
    await saveRoamingStoreAsync();
    persistModelCatalog(getDefaultZaiModels());
    syncModelDropdowns();
    updateAuthenticationStatus();
    setSettingsStatus(
      "dropdown-model-status",
      "Saved settings cleared. Model catalog reset to fallback defaults."
    );

    applyCurrentTheme();
    applyFontSize(blankSettings.fontSize);
    updateReplyButtonVisibility(blankSettings.showReply === "true");
    updateDevBadges(false);

    showNotification("All saved Outlook settings cleared", "success");
    initializeSettingsTabs();
  } catch (error) {
    console.error("Error resetting Outlook settings:", error);
    showNotification(`Failed to reset settings: ${error.message}`, "error");
  }
}

/**
 * Update dev badges visibility
 */
function updateDevBadges(show) {
  const devBadge = document.getElementById("dev-badge");
  const footerDevBadge = document.getElementById("footer-dev-badge");

  if (devBadge) {
    devBadge.style.display = show ? "block" : "none";
  }
  if (footerDevBadge) {
    footerDevBadge.style.display = show ? "block" : "none";
  }
}

// Get event title language from settings
function getEventTitleLanguage() {
  const settings = getSettings();
  return settings.eventTitleLanguage || DEFAULT_SETTINGS.eventTitleLanguage;
}

// Helper function to parse event details using Z.AI
async function parseEventDetailsWithZai(emailContent) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      throw new Error(getMissingApiKeyMessage());
    }

    // Get event title language from settings
    const titleLanguage = getEventTitleLanguage();
    let langInstructions = "";

    // Set language-specific instructions
    if (titleLanguage === "en") {
      langInstructions = `Event title should be in English.
      If the event has a type or category, include it in square brackets ([]) at the beginning, then if there's a presenter and topic, write the presenter's name first, followed by a hyphen (-) and then the topic.`;
    } else if (titleLanguage === "ko") {
      langInstructions = `이벤트 제목은 한국어로 작성해주세요.
      이벤트 유형이나 카테고리가 있다면 대괄호([])로 먼저 표시하고, 발표자와 주제가 있다면 발표자 이름을 먼저 쓰고 하이픈(-) 후에 주제를 적어주세요.`;
    } else if (titleLanguage === "ja") {
      langInstructions = `イベントのタイトルは日本語で記載してください。
      イベントのタイプやカテゴリがある場合は、角括弧（[]）で囲んで最初に表示し、発表者とトピックがある場合は、発表者の名前を最初に書き、ハイフン（-）の後にトピックを書いてください。`;
    } else if (titleLanguage === "zh_cn") {
      langInstructions = `事件标题应该用中文书写。
      如果事件有类型或类别，请使用方括号（[]）将其括起来并放在开头，如果有演讲者和主题，请先写演讲者的名字，然后是连字符（-），再写主题。`;
    } else {
      // Default English instructions for other languages
      langInstructions = `Event title should be in ${getLanguageText(titleLanguage)}.
      If the event has a type or category, include it in square brackets ([]) at the beginning, then if there's a presenter and topic, write the presenter's name first, followed by a hyphen (-) and then the topic.`;
    }

    const prompt = requireTemplate("calendarParse", "Calendar parse")
      .replace("{languageInstructions}", langInstructions)
      .replace("{content}", emailContent);

    console.log("Sending prompt to Z.AI");
    const result = await generateContent(prompt, apiKey, null, false);
    console.log("Received response from Z.AI");

    // Extract only the JSON part from the result (remove any explanations or comments)
    let jsonText = result;

    // Extract text that starts with { and ends with } (JSON only)
    const jsonMatch = result.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      jsonText = jsonMatch[0];
    }

    console.log("Extracted JSON text:", jsonText);

    // Try to parse JSON
    try {
      const eventDetails = JSON.parse(jsonText);
      console.log("Successfully parsed event details:", eventDetails);

      // Validate required fields
      if (!eventDetails.subject) {
        throw new Error("Event title not found.");
      }

      if (!eventDetails.start?.dateTime) {
        throw new Error("Event start time not found.");
      }

      if (!eventDetails.end?.dateTime) {
        throw new Error("Event end time not found.");
      }

      // Validate date format
      const dateRegex = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/;
      if (!dateRegex.test(eventDetails.start.dateTime)) {
        throw new Error("Invalid start time format. (Should be YYYY-MM-DDTHH:mm:ss format)");
      }

      if (!dateRegex.test(eventDetails.end.dateTime)) {
        throw new Error("Invalid end time format. (Should be YYYY-MM-DDTHH:mm:ss format)");
      }

      return eventDetails;
    } catch (parseError) {
      console.error("JSON parsing error:", parseError);
      console.error("Original response:", result);
      throw new Error("Failed to extract event information: " + parseError.message);
    }
  } catch (error) {
    console.error("Error extracting event information:", error);
    throw error;
  }
}

// Create calendar event using Outlook Add-in API
async function createCalendarEvent(eventDetails) {
  try {
    // Validate required fields
    if (!eventDetails.subject || !eventDetails.start?.dateTime || !eventDetails.end?.dateTime) {
      throw new Error("Required event information is missing.");
    }

    // Convert to Outlook Add-in API format
    // Create Date objects from ISO 8601 strings
    const startDate = new Date(eventDetails.start.dateTime);
    const endDate = new Date(eventDetails.end.dateTime);

    // Convert attendee list to string arrays
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

    // Create appointment format - according to Outlook API
    const appointmentData = {
      requiredAttendees: requiredAttendees,
      optionalAttendees: optionalAttendees,
      start: startDate,
      end: endDate,
      location: eventDetails.location?.displayName || "",
      body: eventDetails.body?.content || "",
      subject: eventDetails.subject,
    };

    console.log("Creating appointment with data:", JSON.stringify(appointmentData));

    // Display appointment form - using new appointment creation method
    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees: requiredAttendees,
      optionalAttendees: optionalAttendees,
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

// Check if email content is a calendar event
async function checkIfCalendarEvent(emailContent) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      console.error(getMissingApiKeyMessage());
      return false;
    }

    const prompt = requireTemplate("calendarCheck", "Calendar check").replace(
      "{content}",
      emailContent
    );

    const result = await generateContent(prompt, apiKey, null, true);
    return result.toLowerCase().trim() === "true";
  } catch (error) {
    console.error("Error checking if calendar event:", error);
    return false;
  }
}

// Handle calendar event button click
async function handleCalendarEvent() {
  const apiKey = getApiKey();

  if (!apiKey) {
    showNotification(getMissingApiKeyMessage(), "error");
    toggleSettingsView();
    return;
  }

  showLoading("Creating calendar event...");

  try {
    const emailContent = await getEmailContent();

    // Display result section
    document.getElementById("landing-screen").style.display = "none";
    document.getElementById("result-section").style.display = "block";

    // Display email content and prepare copy button
    document.getElementById("tldr-content").innerHTML = `
      <div class="email-content-header">
        <h3>Email Content</h3>
        <button id="copy-email-content" class="ms-Button ms-Button--primary">
          <span class="ms-Button-label">Copy to Clipboard</span>
        </button>
      </div>
      <div class="email-content-body">
        <pre style="white-space: pre-wrap; word-break: break-word;">${emailContent}</pre>
      </div>
    `;

    // Add event listener to clipboard copy button
    document.getElementById("copy-email-content").addEventListener("click", function () {
      navigator.clipboard
        .writeText(emailContent)
        .then(() => {
          showNotification("Email content copied to clipboard", "info");
        })
        .catch((err) => {
          console.error("Error copying to clipboard:", err);
          showNotification("Failed to copy to clipboard", "error");
        });
    });

    try {
      // Parse event details with Z.AI
      const eventDetails = await parseEventDetailsWithZai(emailContent);

      // Display extracted event details
      document.getElementById("result-content").innerHTML = `
        <div class="event-details-header">
          <h3>Extracted Event Details</h3>
        </div>
        <div class="event-details-body">
          <p><strong>Subject:</strong> ${eventDetails.subject || "Not found"}</p>
          <p><strong>Start:</strong> ${eventDetails.start?.dateTime || "Not found"}</p>
          <p><strong>End:</strong> ${eventDetails.end?.dateTime || "Not found"}</p>
          <p><strong>Location:</strong> ${eventDetails.location?.displayName || "Not found"}</p>
        </div>
      `;

      // Create calendar event
      await createCalendarEvent(eventDetails);
    } catch (extractionError) {
      console.error("Event extraction error:", extractionError);
      // Clean up error message by removing the prefix if present
      const errorMessage = extractionError.message;
      const cleanedMessage = errorMessage.includes("Failed to extract event information:")
        ? errorMessage.split("Failed to extract event information:")[1].trim()
        : errorMessage;

      document.getElementById("result-content").innerHTML = `
        <div class="event-details-header">
          <h3>Event Extraction Failed</h3>
        </div>
        <div class="event-details-body">
          <p class="error-message">${cleanedMessage}</p>
        </div>
      `;

      showNotification(`Event extraction failed: ${cleanedMessage}`, "error");
    }
  } catch (error) {
    console.error("Calendar event handling error:", error);
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    // Hide loading and update button state
    hideLoading();
    updateCalendarButtonState();

    // Enable show full content button
    const expandButton = document.getElementById("expand-content");
    if (expandButton) {
      expandButton.disabled = false;
      expandButton.classList.remove("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
      expandButton.classList.add("ms-Button--primary");

      // Show full content container
      document.getElementById("full-content-container").style.display = "block";
    }
  }
}

// Update calendar button state based on email content
async function updateCalendarButtonState() {
  try {
    const emailContent = await getEmailContent();
    const isCalendarEvent = await checkIfCalendarEvent(emailContent);

    const calendarBtn = document.getElementById("calendar-event");
    if (calendarBtn) {
      if (isCalendarEvent) {
        calendarBtn.disabled = false;
        calendarBtn.classList.remove("action-button--disabled");
        calendarBtn.classList.add("action-button--primary");
        console.log("Calendar event detected, button enabled");
      } else {
        calendarBtn.disabled = true;
        calendarBtn.classList.add("action-button--disabled");
        calendarBtn.classList.remove("action-button--primary");
        console.log("Not a calendar event, button disabled");
      }
    }
  } catch (error) {
    console.error("Error updating calendar button state:", error);
  }
}

/**
 * Export current templates as markdown file
 */
function exportTemplatesAsMarkdown() {
  try {
    const settings = getSettings();
    const templates = settings.templates || createBlankPromptTemplates();
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);

    // Get current user information if available
    let userInfo = "";
    try {
      if (Office.context.mailbox && Office.context.mailbox.userProfile) {
        const user = Office.context.mailbox.userProfile;
        userInfo = `\n\n*Exported by: ${user.displayName} (${user.emailAddress})*`;
      }
    } catch {
      console.log("User profile not available");
    }

    // Create markdown content
    let markdownContent = `# Michael Prompt Templates\n\n`;
    markdownContent += `*Exported on: ${now.toLocaleString()}*${userInfo}\n\n`;

    // Add model information
    markdownContent += `## General Settings\n\n`;
    markdownContent += `- **Provider**: Z.AI GLM Coding Plan\n`;
    markdownContent += `- **Authentication**: Outlook add-in saved setting\n`;
    markdownContent += `- **Model**: ${settings.model || "(empty)"}\n`;
    markdownContent += `- **Reply Model**: ${settings.replyModel || "(empty)"}\n`;
    markdownContent += `- **Default Language**: ${settings.defaultLanguage || DEFAULT_SETTINGS.defaultLanguage}\n`;
    markdownContent += `- **Event Title Language**: ${settings.eventTitleLanguage || DEFAULT_SETTINGS.eventTitleLanguage}\n\n`;

    // Add prompts
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

    // Create download link
    const blob = new Blob([markdownContent], {
      type: "text/markdown",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `michael-templates-${dateStr}.md`;

    // Append to body, click, and remove
    document.body.appendChild(a);
    a.click();

    // Clean up
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);

    showNotification("Templates exported successfully", "success");
  } catch (error) {
    console.error("Error exporting templates:", error);
    showNotification("Failed to export templates", "error");
  }
}
