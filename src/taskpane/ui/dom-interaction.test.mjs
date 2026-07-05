import { test } from "node:test";
import assert from "node:assert/strict";
import { JSDOM } from "jsdom";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

// Load the REAL taskpane.html so DOM-interaction tests run against the actual
// element tree (ids, classes, structure) the production UI uses. Scripts are
// not executed (no runScripts), so Office.js/taskpane.js do not run here.
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = readFileSync(path.resolve(__dirname, "..", "taskpane.html"), "utf8");
const dom = new JSDOM(html, { url: "https://localhost/" });
globalThis.window = dom.window;
globalThis.document = dom.window.document;
Object.defineProperty(globalThis, "navigator", {
  value: dom.window.navigator,
  configurable: true,
});

// jsdom has no clipboard; stub writeText so copyResult/copyReply are testable.
let clipboardText = null;
Object.defineProperty(dom.window.navigator, "clipboard", {
  value: {
    writeText: async (text) => {
      clipboardText = text;
    },
  },
  configurable: true,
});

const {
  showNotification,
  showLoading,
  hideLoading,
  showResults,
  applyFontSize,
  expandContent,
  updateExpandButton,
  updateReplyButtonVisibility,
  stampResultType,
  copyResult,
  copyReply,
  formatReplyOutput,
  TYPES,
} = await import("./dom.js");

test("showNotification injects a #notification node with the message and type", () => {
  document.body.innerHTML = document.body.innerHTML; // ensure clean
  showNotification("saved ok", "success");
  const n = document.getElementById("notification");
  assert.ok(n, "notification node created");
  assert.ok(n.className.includes("success"), `class: ${n.className}`);
  // Message lives in a dedicated span (a manual close button is also rendered).
  const text = n.querySelector(".notification-text");
  assert.ok(text, "notification-text span present");
  assert.equal(text.textContent, "saved ok");
});

test("showNotification replaces any prior notification (no duplicates)", () => {
  showNotification("first", "info");
  showNotification("second", "error");
  const notes = document.querySelectorAll("#notification");
  assert.equal(notes.length, 1);
  assert.equal(notes[0].querySelector(".notification-text").textContent, "second");
});

test("showLoading shows #loading and hides landing/result sections", () => {
  showLoading("working");
  assert.equal(document.getElementById("loading").style.display, "block");
  assert.equal(document.getElementById("landing-screen").style.display, "none");
  assert.equal(document.getElementById("result-section").style.display, "none");
  assert.equal(document.getElementById("loading-message").textContent, "working");
});

test("hideLoading hides #loading", () => {
  showLoading("working");
  hideLoading();
  assert.equal(document.getElementById("loading").style.display, "none");
});

test("applyFontSize sets data-font-size on the document root", () => {
  applyFontSize("large");
  assert.equal(document.documentElement.getAttribute("data-font-size"), "large");
  applyFontSize("");
  assert.equal(document.documentElement.getAttribute("data-font-size"), "medium");
});

test("showResults renders sanitized markdown into #tldr-content and reveals the section", () => {
  showResults("**hello** <script>alert(1)</scr" + "ipt>", TYPES.SUMMARIZE);
  const tldr = document.getElementById("tldr-content");
  assert.match(tldr.innerHTML, /<strong>hello<\/strong>/);
  assert.ok(!/<script/i.test(tldr.innerHTML), `script leaked: ${tldr.innerHTML}`);
  assert.equal(document.getElementById("result-section").style.display, "block");
});

test("updateExpandButton toggles disabled state and label", () => {
  updateExpandButton(false);
  const btn = document.getElementById("expand-content");
  assert.equal(btn.disabled, true);
  assert.match(btn.innerHTML, /Loading Full Content/);
  updateExpandButton(true);
  assert.equal(btn.disabled, false);
  assert.match(btn.innerHTML, /Show Full Content/);
});

test("expandContent toggles the full-content container visibility", () => {
  updateExpandButton(true); // arm the button
  const container = document.getElementById("full-content-container");
  container.style.display = "none";
  expandContent();
  assert.equal(container.style.display, "block");
  expandContent();
  assert.equal(container.style.display, "none");
});

test("updateReplyButtonVisibility toggles #generate-reply display", () => {
  updateReplyButtonVisibility(true);
  assert.equal(document.getElementById("generate-reply").style.display, "inline-block");
  updateReplyButtonVisibility(false);
  assert.equal(document.getElementById("generate-reply").style.display, "none");
});

test("stampResultType records the type (+ optional reply raw) on #result-section", () => {
  stampResultType("reply", "Subject: Hi\n\nBody text");
  const section = document.getElementById("result-section");
  assert.equal(section.getAttribute("data-result-type"), "reply");
  assert.equal(section.getAttribute("data-reply-raw"), "Subject: Hi\n\nBody text");

  stampResultType("summarize");
  assert.equal(section.getAttribute("data-result-type"), "summarize");
  assert.equal(section.hasAttribute("data-reply-raw"), false);
});

test("copyResult uses the structured reply text (not Subject-substring inference)", () => {
  // Regression: a summary quoting "Subject:" used to be misread as a reply.
  clipboardText = null;
  stampResultType("reply", "Subject: Re: Hello\n\nReal reply body");
  copyResult();
  assert.equal(clipboardText, "Subject: Re: Hello\n\nReal reply body");
});

test("copyResult copies visible TL;DR text for a non-reply result", async () => {
  clipboardText = null;
  stampResultType("summarize");
  document.getElementById("full-content-container").style.display = "none";
  document.getElementById("tldr-content").innerText = "Subject: quoted line in a summary\nDetail";
  copyResult();
  // Give the .then() callback a tick to run.
  await new Promise((r) => setTimeout(r, 0));
  assert.equal(
    clipboardText,
    "Subject: quoted line in a summary\nDetail",
    "summary is NOT misread as a reply"
  );
});

test("copyReply reads the structured reply raw; falls back when absent", async () => {
  clipboardText = null;
  stampResultType("reply", "Subject: Hi\n\nBody");
  copyReply();
  await new Promise((r) => setTimeout(r, 0));
  assert.equal(clipboardText, "Subject: Hi\n\nBody");
});

test("formatReplyOutput: single-line reply yields an empty body (no marker echo)", () => {
  const out = formatReplyOutput("Subject: Hi");
  assert.equal(out.subject, "Hi");
  assert.equal(out.body, "");
});

test("formatReplyOutput: empty reply falls back to a placeholder subject", () => {
  const out = formatReplyOutput("   ");
  assert.equal(out.subject, "Re: Your email");
  assert.equal(out.body, "");
});
