import { test } from "node:test";
import assert from "node:assert/strict";
import {
  DEFAULT_SETTINGS,
  mergeWithDefaults,
  normalizeSettings,
  isAutorunEnabled,
  getAutorunAction,
} from "./settings.js";

test("mergeWithDefaults fills missing keys", () => {
  const merged = mergeWithDefaults({ theme: "dark", apiKey: "k" });
  assert.equal(merged.theme, "dark");
  assert.equal(merged.apiKey, "k");
  assert.equal(merged.fontSize, DEFAULT_SETTINGS.fontSize);
  assert.equal(merged.defaultLanguage, DEFAULT_SETTINGS.defaultLanguage);
});

test("isAutorunEnabled is strictly the string 'true'", () => {
  assert.equal(isAutorunEnabled({ autorun: "true" }), true);
  assert.equal(isAutorunEnabled({ autorun: "false" }), false);
  assert.equal(isAutorunEnabled({ autorun: true }), false); // boolean, not string
  assert.equal(isAutorunEnabled({}), false);
});

test("getAutorunAction returns null when disabled", () => {
  assert.equal(getAutorunAction({ autorun: "false", autorunOption: "summarize" }), null);
  assert.equal(getAutorunAction({}), null);
});

test("getAutorunAction returns the action when enabled and valid", () => {
  assert.equal(getAutorunAction({ autorun: "true", autorunOption: "summarize" }), "summarize");
  assert.equal(getAutorunAction({ autorun: "true", autorunOption: "reply" }), "reply");
  assert.equal(
    getAutorunAction({ autorun: "true", autorunOption: "translateAndSummarize" }),
    "translateAndSummarize"
  );
});

test("getAutorunAction returns null when option is unknown", () => {
  assert.equal(getAutorunAction({ autorun: "true", autorunOption: "bogus" }), null);
  assert.equal(getAutorunAction({ autorun: "true" }), null);
});

test("normalizeSettings drops unknown keys, keeps templates", () => {
  const out = normalizeSettings({ theme: "dark", bogus: 1, templates: { summarize: "s" } });
  assert.equal("bogus" in out, false);
  assert.equal(out.theme, "dark");
  assert.deepEqual(out.templates, { summarize: "s" });
});
