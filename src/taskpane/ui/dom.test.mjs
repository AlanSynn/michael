import { test } from "node:test";
import assert from "node:assert/strict";
import { JSDOM } from "jsdom";

// dom.js now bundles marked + DOMPurify (real imports, not CDN globals), so the
// test environment must provide a DOM before importing it. DOMPurify reads
// window/document at import time. navigator is a getter on Node >=21, so define
// it explicitly.
const dom = new JSDOM("<!DOCTYPE html><html><body></body></html>");
globalThis.window = dom.window;
globalThis.document = dom.window.document;
Object.defineProperty(globalThis, "navigator", {
  value: dom.window.navigator,
  configurable: true,
});

const { escapeHtml, renderMarkdown } = await import("./dom.js");

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
  assert.equal(
    escapeHtml(`a & b < c > d "e" 'f'`),
    "a &amp; b &lt; c &gt; d &quot;e&quot; &#39;f&#39;"
  );
});

test("escapeHtml coerces non-strings", () => {
  assert.equal(escapeHtml(42), "42");
  assert.equal(escapeHtml(["a", "b"]), "a,b");
});

test("renderMarkdown keeps safe markdown formatting", () => {
  const out = renderMarkdown("**bold** and `code`");
  assert.match(out, /<strong>bold<\/strong>/);
  assert.match(out, /<code>code<\/code>/);
});

test("renderMarkdown strips script injection and event handlers", () => {
  const out = renderMarkdown(
    "hi <script>alert(1)</scr" + "ipt><img src=x onerror=alert(1)>"
  );
  assert.ok(!/<script/i.test(out), `<script> leaked: ${out}`);
  assert.ok(!/onerror/i.test(out), `onerror leaked: ${out}`);
  assert.match(out, /hi/); // benign text preserved
});

test("renderMarkdown null/empty yields empty string (no crash)", () => {
  assert.equal(renderMarkdown(null), "");
  assert.equal(renderMarkdown(""), "");
});
