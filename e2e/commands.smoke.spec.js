// E2e smoke test for the function-file host page (dist/commands.html). This is
// the headless page Outlook loads to execute ribbon commands such as Quick
// Translate — it has no UI, just office.js + commands.js. It must:
//   1. Serve HTTP 200 (a broken chunk reference here 404s and bricks ribbon
//      commands silently).
//   2. Load EXACTLY two scripts: the office.js CDN + the bundled commands.js
//      entry. A prior build emitted a phantom hashed chunk here (html-loader
//      treating a manual <script src="commands.js"> as a relative asset) whose
//      filename was never produced, 404-ing the function-file. This pins that.
//   3. Have an empty body (it is a function-file host, not a UI page).
//   4. Load with no console errors / failed requests.
import { test, expect } from "@playwright/test";

const violations = (page) => {
  const list = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") list.push(`console.error: ${msg.text()}`);
  });
  page.on("pageerror", (err) => list.push(`pageerror: ${err.message}`));
  page.on("requestfailed", (req) => {
    const url = req.url();
    if (!url.endsWith("/favicon.ico")) {
      list.push(`requestfailed: ${url} (${req.failure()?.errorText || "unknown"})`);
    }
  });
  return list;
};

test("commands.html serves 200 with only office.js + commands.js (no phantom chunk)", async ({
  page,
}) => {
  const errors = violations(page);
  const response = await page.goto("/commands.html", { waitUntil: "load" });
  expect(response?.status(), "commands.html must serve 200").toBe(200);
  // office.js fires a console.info on load; give async errors a beat to surface.
  await page.waitForTimeout(300);

  const srcs = await page.$$eval("script[src]", (els) =>
    els.map((e) => e.getAttribute("src"))
  );

  // The bundled entry is injected exactly once.
  const entry = srcs.filter((s) => s === "commands.js");
  expect(entry, `commands.js entry scripts: ${srcs.join(", ")}`).toHaveLength(1);

  // office.js legitimately loads additional scripts from its own CDN at runtime
  // (office.js + o15apptofilemappingtable.js). The regression we pin here is a
  // PHANTOM CHUNK — a hashed filename the build referenced but never emitted,
  // which 404s the function-file silently. Guard: every script src must be
  // either the bundled "commands.js" entry or belong to the office.js CDN
  // family. Any other src (e.g. commands.<hash>.js) is the bug.
  const unexpected = srcs.filter(
    (s) => s !== "commands.js" && !s.includes("appsforoffice.microsoft.com")
  );
  expect(
    unexpected,
    `unexpected script srcs (phantom chunk?): ${unexpected.join(", ")}`
  ).toEqual([]);

  // Function-file host page renders no UI.
  await expect(page.locator("body")).toBeEmpty();

  expect(errors, errors.join("\n")).toEqual([]);
});
