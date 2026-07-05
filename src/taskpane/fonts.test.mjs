import { test } from "node:test";
import assert from "node:assert/strict";
import { getFontSizeValue, FONT_SIZE_OPTIONS } from "./fonts.js";

test("getFontSizeValue maps tokens to rem values", () => {
  assert.equal(getFontSizeValue("small"), "0.875rem");
  assert.equal(getFontSizeValue("medium"), "1rem");
  assert.equal(getFontSizeValue("large"), "1.125rem");
});

test("getFontSizeValue defaults to 1rem for unknown/empty", () => {
  assert.equal(getFontSizeValue("huge"), "1rem");
  assert.equal(getFontSizeValue(undefined), "1rem");
  assert.equal(getFontSizeValue(""), "1rem");
});

test("FONT_SIZE_OPTIONS has small/medium/large", () => {
  const values = FONT_SIZE_OPTIONS.map((o) => o.value);
  assert.deepEqual(values, ["small", "medium", "large"]);
});
