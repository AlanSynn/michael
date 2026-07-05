import { test } from "node:test";
import assert from "node:assert/strict";
import { getCatalog, setCatalog, getCachedModelCatalog } from "./model-catalog.js";
import { getDefaultZaiModels } from "../shared/zai.js";

test("getCatalog returns a copy (mutating result does not affect state)", () => {
  const before = getCatalog();
  const copy = getCatalog();
  copy.push("glm-fake");
  assert.deepEqual(getCatalog(), before);
});

test("setCatalog stores a copy and fills empty input with defaults", () => {
  setCatalog(["glm-a", "glm-b"]);
  assert.deepEqual(getCatalog(), ["glm-a", "glm-b"]);
  setCatalog([]);
  assert.deepEqual(getCatalog(), getDefaultZaiModels());
  setCatalog(undefined);
  assert.deepEqual(getCatalog(), getDefaultZaiModels());
});

test("getCachedModelCatalog falls back when cache is empty/invalid", () => {
  assert.deepEqual(getCachedModelCatalog([]), getDefaultZaiModels());
  assert.deepEqual(getCachedModelCatalog(null), getDefaultZaiModels());
  assert.deepEqual(getCachedModelCatalog("nope"), getDefaultZaiModels());
});

test("getCachedModelCatalog uses a valid cache", () => {
  assert.deepEqual(getCachedModelCatalog(["glm-x", "glm-y"]), ["glm-x", "glm-y"]);
});

test("setCatalog dedupes duplicate model names (case-insensitive)", () => {
  // Regression: a stale cache blob or repeated provider entry previously
  // rendered duplicate <option>s because boundModels skipped dedupe.
  setCatalog(["glm-4.5", "glm-4.5", "GLM-4.5", "glm-4.6"]);
  assert.deepEqual(getCatalog(), ["glm-4.5", "glm-4.6"]);
});

test("getCachedModelCatalog dedupes duplicate model names", () => {
  assert.deepEqual(
    getCachedModelCatalog(["glm-x", "glm-x", "glm-y", "GLM-X"]),
    ["glm-x", "glm-y"]
  );
});

test("catalog is capped at the max bound (overflow truncated)", () => {
  const big = Array.from({ length: 250 }, (_, i) => `glm-model-${i}`);
  setCatalog(big);
  const catalog = getCatalog();
  // MAX_MODEL_CATALOG = 200. Cap is enforced AFTER dedupe, so 250 unique
  // entries collapse to the first 200.
  assert.equal(catalog.length, 200);
  assert.equal(catalog[0], "glm-model-0");
  assert.equal(catalog[199], "glm-model-199");
});
