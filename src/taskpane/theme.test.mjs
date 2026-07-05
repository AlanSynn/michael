import { test } from "node:test";
import assert from "node:assert/strict";
import { isDarkTheme, resolveTheme, DEFAULT_THEME_FALLBACK } from "./theme.js";

test("isDarkTheme detects dark and light colors", () => {
  assert.equal(isDarkTheme("#000000"), true);
  assert.equal(isDarkTheme("#252423"), true);
  assert.equal(isDarkTheme("#ffffff"), false);
  assert.equal(isDarkTheme("#f3f2f1"), false);
});

test("isDarkTheme tolerates # prefix and is case-insensitive on digits", () => {
  assert.equal(isDarkTheme("000000"), true);
  assert.equal(isDarkTheme("FFFFFF"), false);
});

test("isDarkTheme returns false for invalid/empty input", () => {
  assert.equal(isDarkTheme(undefined), false);
  assert.equal(isDarkTheme(""), false);
  assert.equal(isDarkTheme("nothex"), false);
  assert.equal(isDarkTheme("#12"), false);
});

test("resolveTheme honors explicit light/dark", () => {
  assert.equal(resolveTheme("light"), "light");
  assert.equal(resolveTheme("dark"), "dark");
});

test("resolveTheme follows Office background color when system/empty", () => {
  assert.equal(resolveTheme("system", { bodyBackgroundColor: "#252423" }), "dark");
  assert.equal(resolveTheme("system", { bodyBackgroundColor: "#ffffff" }), "light");
  assert.equal(resolveTheme(undefined, { bodyBackgroundColor: "#000000" }), "dark");
});

test("resolveTheme falls back when no theme info available", () => {
  assert.equal(resolveTheme(undefined, undefined), DEFAULT_THEME_FALLBACK);
  assert.equal(resolveTheme("system", undefined), DEFAULT_THEME_FALLBACK);
  assert.equal(resolveTheme("", {}), DEFAULT_THEME_FALLBACK);
});
