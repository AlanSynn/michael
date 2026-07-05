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
