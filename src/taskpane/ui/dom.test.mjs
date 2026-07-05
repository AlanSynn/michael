import { test } from "node:test";
import assert from "node:assert/strict";

// dom.js imports only pure modules at module scope (storage -> prompt-templates);
// DOM/CDN globals (document, marked, DOMPurify) are touched inside function
// bodies, so escapeHtml is testable in plain Node.
const { escapeHtml } = await import("./dom.js");

test("escapeHtml returns empty for null/undefined", () => {
  assert.equal(escapeHtml(null), "");
  assert.equal(escapeHtml(undefined), "");
});

test("escapeHtml neutralizes the obvious XSS payloads", () => {
  assert.equal(
    escapeHtml("<script>alert(1)</script>"),
    "&lt;script&gt;alert(1)&lt;/script&gt;"
  );
  assert.equal(
    escapeHtml(`<img src=x onerror="alert(1)">`),
    "&lt;img src=x onerror=&quot;alert(1)&quot;&gt;"
  );
});

test("escapeHtml escapes every HTML-special character", () => {
  assert.equal(escapeHtml(`a & b < c > d "e" 'f'`), "a &amp; b &lt; c &gt; d &quot;e&quot; &#39;f&#39;");
});

test("escapeHtml coerces non-strings", () => {
  assert.equal(escapeHtml(42), "42");
  assert.equal(escapeHtml(["a", "b"]), "a,b");
});

test("renderMarkdown sanitizes marked output via DOMPurify", async () => {
  // Stub the CDN globals so renderMarkdown is exercised end-to-end.
  globalThis.marked = { parse: (src) => `<p>${src}</p>` };
  let sanitizedInput = null;
  globalThis.DOMPurify = { sanitize: (html) => ((sanitizedInput = html), "CLEAN") };

  const { renderMarkdown } = await import("./dom.js");

  assert.equal(renderMarkdown("**hi**"), "CLEAN");
  assert.equal(sanitizedInput, "<p>**hi**</p>");
  assert.equal(renderMarkdown(null), "CLEAN"); // null -> "" -> parsed

  delete globalThis.marked;
  delete globalThis.DOMPurify;
});
