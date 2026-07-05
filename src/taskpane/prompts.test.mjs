import { test } from "node:test";
import assert from "node:assert/strict";
import { fillTemplate } from "./prompts.js";

test("fills subject/content/language placeholders", () => {
  const out = fillTemplate("S:{subject} C:{content} L:{language}", {
    subject: "hi",
    content: "body",
    language: "Korean",
  });
  assert.equal(out, "S:hi C:body L:Korean");
});

test("null/undefined vars become empty string", () => {
  assert.equal(fillTemplate("[{subject}]", {}), "[]");
  assert.equal(fillTemplate("[{subject}]", { subject: null }), "[]");
  assert.equal(fillTemplate("[{subject}]", { subject: undefined }), "[]");
});

test("does not cascade — a value containing a placeholder literal is not re-expanded", () => {
  const out = fillTemplate("{content}", { content: "{subject}" });
  assert.equal(out, "{subject}");
});

test("values containing $ are treated literally (no $-pattern expansion)", () => {
  const out = fillTemplate("{content}", { content: "$100 $$ &" });
  assert.equal(out, "$100 $$ &");
});

test("languageInstructions placeholder is supported", () => {
  const out = fillTemplate("rule: {languageInstructions}", { languageInstructions: "use Korean" });
  assert.equal(out, "rule: use Korean");
});

test("non-string template returns empty string", () => {
  assert.equal(fillTemplate(undefined, {}), "");
  assert.equal(fillTemplate(null, {}), "");
});
