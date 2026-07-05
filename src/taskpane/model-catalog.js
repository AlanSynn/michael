// In-memory model catalog state + pure cache-resolution helper.
// Persistence lives in storage.js; this module owns only the live copy so it
// can be unit-tested without Office.

import { getDefaultZaiModels } from "../shared/zai.js";

// Upper bound on the catalog. Z.AI returns ~30 glm-* models; a malformed
// provider response or a stale cache blob should never be able to balloon the
// in-memory catalog (and the model <select> DOM) into the thousands.
const MAX_MODEL_CATALOG = 200;

let catalog = getDefaultZaiModels();

/**
 * Coerce a candidate list into a bounded, non-empty array of strings, falling
 * back to the default Z.AI models when the input is empty/invalid.
 * @param {unknown} models
 * @returns {string[]}
 */
function boundModels(models) {
  if (!Array.isArray(models) || models.length === 0) {
    return getDefaultZaiModels();
  }
  return models.slice(0, MAX_MODEL_CATALOG);
}

/** @returns {string[]} a defensive copy of the live catalog */
export function getCatalog() {
  return catalog.length ? [...catalog] : getDefaultZaiModels();
}

/** @param {string[]} models @returns {string[]} */
export function setCatalog(models) {
  catalog = boundModels(models);
  return getCatalog();
}

/**
 * Resolve the catalog from a persisted JSON payload, falling back to the
 * Z.AI default model list when the cache is empty/invalid.
 * @param {unknown} cached
 * @returns {string[]}
 */
export function getCachedModelCatalog(cached) {
  return boundModels(cached);
}
