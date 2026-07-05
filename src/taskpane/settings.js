// Settings-shape logic over a plain object — no storage, no DOM.
// getAutorunAction consolidates the autorun switch that was duplicated 3x
// across Office.onReady and the ItemChanged handlers.

import { DEFAULT_SETTINGS } from "./prompt-templates.js";

export { DEFAULT_SETTINGS };

/** Scalar settings keys (templates are nested separately). */
export const SETTINGS_FIELD_KEYS = Object.freeze([
  "model",
  "replyModel",
  "defaultLanguage",
  "eventTitleLanguage",
  "theme",
  "fontSize",
  "tldrMode",
  "showReply",
  "autorun",
  "autorunOption",
  "devMode",
  "devServer",
  "apiKey",
]);

const AUTORUN_ACTIONS = Object.freeze(["summarize", "translate", "translateAndSummarize", "reply"]);

/**
 * Merge a partial settings object over DEFAULT_SETTINGS.
 * @param {Record<string, *>} partial
 * @returns {Record<string, *>}
 */
export function mergeWithDefaults(partial = {}) {
  return { ...DEFAULT_SETTINGS, ...partial };
}

/**
 * Keep only recognized settings fields (+ a shallow templates copy).
 * @param {Record<string, *>} input
 * @returns {Record<string, *>}
 */
export function normalizeSettings(input = {}) {
  const out = {};
  for (const key of SETTINGS_FIELD_KEYS) {
    if (input[key] !== undefined) {
      out[key] = input[key];
    }
  }
  if (input.templates && typeof input.templates === "object") {
    out.templates = { ...input.templates };
  }
  return out;
}

export function isAutorunEnabled(settings = {}) {
  return settings.autorun === "true";
}

/**
 * Resolve the autorun flow to invoke, or null when autorun is off / option invalid.
 * @param {Record<string, *>} settings
 * @returns {"summarize"|"translate"|"translateAndSummarize"|"reply"|null}
 */
export function getAutorunAction(settings = {}) {
  if (!isAutorunEnabled(settings)) {
    return null;
  }
  const action = settings.autorunOption;
  return AUTORUN_ACTIONS.includes(action) ? action : null;
}
