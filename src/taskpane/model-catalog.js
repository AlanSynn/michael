// In-memory model catalog state + pure cache-resolution helper.
// Persistence lives in storage.js; this module owns only the live copy so it
// can be unit-tested without Office.

import { getDefaultZaiModels } from "../shared/zai.js";

let catalog = getDefaultZaiModels();

/** @returns {string[]} a defensive copy of the live catalog */
export function getCatalog() {
  return catalog.length ? [...catalog] : getDefaultZaiModels();
}

/** @param {string[]} models @returns {string[]} */
export function setCatalog(models) {
  catalog = Array.isArray(models) && models.length ? [...models] : getDefaultZaiModels();
  return getCatalog();
}

/**
 * Resolve the catalog from a persisted JSON payload, falling back to the
 * Z.AI default model list when the cache is empty/invalid.
 * @param {unknown} cached
 * @returns {string[]}
 */
export function getCachedModelCatalog(cached) {
  return Array.isArray(cached) && cached.length ? [...cached] : getDefaultZaiModels();
}
