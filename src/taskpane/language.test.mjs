import { test } from "node:test";
import assert from "node:assert/strict";
import { getLanguageText, LANGUAGE_OPTIONS } from "./language.js";

test("getLanguageText maps known codes to display names", () => {
  assert.equal(getLanguageText("ko"), "Korean");
  assert.equal(getLanguageText("en"), "English");
  assert.equal(getLanguageText("ja"), "Japanese");
  assert.equal(getLanguageText("zh_cn"), "Chinese");
  assert.equal(getLanguageText("es"), "Spanish");
  assert.equal(getLanguageText("fr"), "French");
  assert.equal(getLanguageText("de"), "German");
  assert.equal(getLanguageText("it"), "Italian");
});

test("getLanguageText defaults to English for unknown codes", () => {
  assert.equal(getLanguageText("xyz"), "English");
  assert.equal(getLanguageText(undefined), "English");
  assert.equal(getLanguageText(""), "English");
});

test("LANGUAGE_OPTIONS values are unique", () => {
  const values = LANGUAGE_OPTIONS.map((o) => o.value);
  assert.equal(new Set(values).size, values.length);
});
