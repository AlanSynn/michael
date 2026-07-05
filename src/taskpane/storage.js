/* global console, localStorage, sessionStorage, Office */

// The single Office roamingSettings storage boundary.
// All Office.context.roamingSettings access lives here; everything else calls
// these helpers so the Office coupling stays in one testable place.

import { createBlankPromptTemplates, createBlankSettings } from "./prompt-templates.js";

export const SETTINGS_KEY = "michael_settings";
export const TEMPLATE_DEFAULTS_KEY = "michael_template_defaults";
export const MODEL_CATALOG_CACHE_KEY = "michael_zai_model_catalog";

export function getRoamingStore() {
  return (typeof Office !== "undefined" && Office?.context?.roamingSettings) || null;
}

export function saveRoamingStoreAsync() {
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

export const roamingStorage = Object.freeze({
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
export function migrateSettingsKeys() {
  try {
    let migrated = false;
    // Track whether settings were already sourced from ANY legacy path. The
    // sessionStorage blob below is the "primary" legacy store; without this
    // flag, the localStorage branch would reuse the now-stale `currentSettings`
    // snapshot (still null) and silently OVERWRITE the just-migrated
    // sessionStorage value with the older website-era blob — losing every
    // field that existed only in the add-in's session settings (templates etc.).
    let settingsMigrated = false;
    const currentSettings = roamingStorage.getItem(SETTINGS_KEY);
    if (currentSettings) {
      settingsMigrated = true;
    }

    const legacySettings = sessionStorage.getItem("my_sidekick_michael_settings");
    if (!settingsMigrated && legacySettings) {
      roamingStorage.setItem(SETTINGS_KEY, legacySettings);
      migrated = true;
      settingsMigrated = true;
    }
    sessionStorage.removeItem("my_sidekick_michael_settings");

    // Legacy localStorage on the alansynn.com origin carried the API key.
    // Copy once (only if no settings came from the primary session path), then
    // clear it so the key is not left sitting in browser storage after
    // migration to Office-managed roamingSettings.
    const previousLocalSettings = localStorage.getItem("michael_settings");
    if (!settingsMigrated && previousLocalSettings) {
      roamingStorage.setItem(SETTINGS_KEY, previousLocalSettings);
      migrated = true;
      settingsMigrated = true;
    }
    if (previousLocalSettings) {
      localStorage.removeItem("michael_settings");
    }

    const legacyTemplateDefaults = sessionStorage.getItem("michael_session_template_defaults");
    const currentTemplateDefaults = roamingStorage.getItem(TEMPLATE_DEFAULTS_KEY);
    if (!currentTemplateDefaults && legacyTemplateDefaults) {
      roamingStorage.setItem(TEMPLATE_DEFAULTS_KEY, legacyTemplateDefaults);
      migrated = true;
    }

    const legacyModelCatalog = sessionStorage.getItem(MODEL_CATALOG_CACHE_KEY);
    const currentModelCatalog = roamingStorage.getItem(MODEL_CATALOG_CACHE_KEY);
    if (!currentModelCatalog && legacyModelCatalog) {
      roamingStorage.setItem(MODEL_CATALOG_CACHE_KEY, legacyModelCatalog);
      migrated = true;
    }

    // Only round-trip to Exchange when something actually moved; previously
    // this fired on every bootstrap even when no migration occurred.
    if (migrated) {
      saveRoamingStoreAsync().catch((error) => {
        console.error("Error saving migrated Outlook add-in settings:", error);
      });
    }
  } catch {
    // no-op
  }
}

export function readJsonStorage(key, fallbackValue) {
  try {
    const rawValue = roamingStorage.getItem(key);
    return rawValue ? JSON.parse(rawValue) : fallbackValue;
  } catch (error) {
    console.error(`Error reading ${key}:`, error);
    return fallbackValue;
  }
}

export function writeJsonStorage(key, value) {
  roamingStorage.setItem(key, JSON.stringify(value));
}

export function getSettings() {
  return readJsonStorage(SETTINGS_KEY, createBlankSettings());
}

export function saveSettingsToStore(settings) {
  writeJsonStorage(SETTINGS_KEY, settings);
}

export function getSavedTemplateDefaults() {
  return readJsonStorage(TEMPLATE_DEFAULTS_KEY, createBlankPromptTemplates());
}

export function saveTemplateDefaults(templates) {
  writeJsonStorage(TEMPLATE_DEFAULTS_KEY, templates);
}

/** Read the persisted model catalog cache (raw; resolution lives in model-catalog). */
export function readCachedModelCatalog() {
  return readJsonStorage(MODEL_CATALOG_CACHE_KEY, []);
}

/** Persist the model catalog cache to roaming settings and flush. */
export function writeCachedModelCatalog(models) {
  writeJsonStorage(MODEL_CATALOG_CACHE_KEY, models);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving model catalog:", error);
  });
}
