import { test } from "node:test";
import assert from "node:assert/strict";
import {
  DEFAULT_SETTINGS,
  DEFAULT_PROMPT_TEMPLATES,
  TEMPLATE_KEYS,
  createDefaultSettings,
  createBlankSettings,
  createDefaultPromptTemplates,
} from "./prompt-templates.js";

test("DEFAULT_PROMPT_TEMPLATES has all template keys non-empty", () => {
  for (const key of TEMPLATE_KEYS) {
    assert.equal(typeof DEFAULT_PROMPT_TEMPLATES[key], "string");
    assert.ok(DEFAULT_PROMPT_TEMPLATES[key].length > 0, `template "${key}" is empty`);
  }
});

test("factory fns return distinct, well-shaped objects", () => {
  const a = createDefaultSettings();
  const b = createDefaultSettings();
  assert.notEqual(a, b);
  assert.deepEqual(a, b);
  assert.equal(typeof a.model, "string");
  assert.ok(a.templates && typeof a.templates === "object");
});

test("blank settings keep DEFAULT_SETTINGS scalars but blank templates", () => {
  const blank = createBlankSettings();
  assert.equal(blank.model, DEFAULT_SETTINGS.model);
  for (const key of TEMPLATE_KEYS) {
    assert.equal(blank.templates[key], "");
  }
});

test("createDefaultPromptTemplates covers every key", () => {
  const t = createDefaultPromptTemplates();
  for (const key of TEMPLATE_KEYS) {
    assert.equal(t[key], DEFAULT_PROMPT_TEMPLATES[key]);
  }
});
