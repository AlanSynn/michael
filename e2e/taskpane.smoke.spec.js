// E2e smoke tests for the built taskpane. Loads dist/taskpane.html in a real
// headless browser and asserts:
//   1. No console errors / CSP violations / failed requests (catches the exact
//      SRI/CDN/CSP class of bug that broke production DOMPurify).
//   2. office.js loads and taskpane.js executes.
//   3. The static UI renders (brand text, key buttons present).
//   4. With a mocked Office host injected before taskpane.js runs, bootstrap
//      wires listeners and the settings toggle opens the settings panel.
//
// Office.onReady never fires outside the real Office host, so the unmocked
// tests are scoped to static smoke + CSP. The mocked test verifies wiring.
import { test, expect } from "@playwright/test";

const violations = (page) => {
  const list = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") list.push(`console.error: ${msg.text()}`);
  });
  page.on("pageerror", (err) => list.push(`pageerror: ${err.message}`));
  page.on("requestfailed", (req) => {
    const url = req.url();
    // office.js occasionally logs a benign 304/abort under headless load; only
    // flag hard failures (non-2xx that aren't the expected favicon).
    if (!url.endsWith("/favicon.ico")) {
      list.push(`requestfailed: ${url} (${req.failure()?.errorText || "unknown"})`);
    }
  });
  return list;
};

test.describe("static smoke (no Office host)", () => {
  test("loads without console errors, CSP violations, or failed requests", async ({ page }) => {
    const errors = violations(page);
    await page.goto("/taskpane.html", { waitUntil: "load" });
    // Give async script-load errors a chance to surface.
    await page.waitForTimeout(500);
    expect(errors, errors.join("\n")).toEqual([]);
  });

  test("office.js is loaded onto window", async ({ page }) => {
    violations(page);
    await page.goto("/taskpane.html", { waitUntil: "load" });
    const hasOffice = await page.evaluate(() => typeof window.Office === "object");
    expect(hasOffice).toBeTruthy();
  });

  test("static UI renders brand and primary buttons", async ({ page }) => {
    violations(page);
    await page.goto("/taskpane.html", { waitUntil: "load" });
    await expect(page.locator(".brand-name-text")).toContainText("Michael");
    // Primary action buttons exist in the DOM (click handlers are wired only
    // after Office.onReady, which is mocked in the integration test below).
    await expect(page.locator("#summarize")).toBeVisible();
    await expect(page.locator("#translate")).toBeVisible();
  });

  test("taskpane.js bundle is injected (sanitizer is bundled, no CDN script)", async ({ page }) => {
    violations(page);
    await page.goto("/taskpane.html", { waitUntil: "load" });
    // taskpane.js is the bundled entry; marked/dompurify live inside it, so no
    // separate <script src> for them should be present.
    const cdnScripts = await page.locator("script[src*='dompurify'], script[src*='marked']").count();
    expect(cdnScripts).toBe(0);
    const bundle = await page.locator("script[src*='taskpane.js']").count();
    expect(bundle).toBeGreaterThanOrEqual(1);
  });
});

// Intercepts the office.js CDN request and fulfills it with a stub that defines
// window.Office, so Office.onReady fires synchronously and bootstrapOutlook wires
// its listeners. We must route the CDN script (rather than set window.Office in an
// initScript) because the real office.js loads after initScripts and would
// overwrite our mock.
const OFFICE_STUB = `
(function () {
  var store = {};
  var item = {
    itemId: "mock-item-1",
    subject: "Q3 roadmap review",
    body: {
      getAsync: function (coercion, cb) {
        cb({ status: "succeeded", value: "Team, please review the attached Q3 roadmap before Friday." });
      },
    },
  };
  window.Office = {
    HostType: { Outlook: "Outlook" },
    PlatformType: { Mac: "Mac", PC: "PC" },
    EventType: { ItemChanged: "itemChanged", SettingsChanged: "settingsChanged" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Text: "Text", Html: "Html" },
    context: {
      mailbox: {
        item: item,
        addHandlerAsync: function () {},
        removeHandlerAsync: function () {},
        userProfile: { displayName: "Test User", emailAddress: "test@example.com" },
      },
      roamingSettings: {
        get: function (key) { return key in store ? store[key] : null; },
        set: function (key, val) { store[key] = val; },
        remove: function (key) { delete store[key]; },
        saveAsync: function (cb) { if (cb) cb({ status: "succeeded" }); },
      },
      officeTheme: { bodyBackgroundColor: "#ffffff" },
      requirements: { isSetSupported: function () { return true; } },
    },
    onReady: function (cb) { cb({ host: "Outlook", platform: "Mac" }); },
    initialize: function () {},
    actions: { associate: function () {} },
  };
  // Z.AI fetch will be attempted on bootstrap (model catalog refresh with
  // silent:true). Stub it to avoid real network + 401 noise.
  window.fetch = function () {
    return Promise.resolve({
      ok: true,
      status: 200,
      json: function () { return Promise.resolve({ data: [] }); },
      text: function () { return Promise.resolve(""); },
    });
  };
})();
`;

test.describe("integration (mocked Office host)", () => {
  test("bootstrap wires the settings toggle and reveals the settings panel", async ({ page }) => {
    const errors = violations(page);
    await page.route("**/office.js", (route) =>
      route.fulfill({ contentType: "text/javascript", body: OFFICE_STUB })
    );
    await page.goto("/taskpane.html", { waitUntil: "load" });
    // bootstrapOutlook() flips #app-body to visible (and hides #sideload-msg)
    // once Office.onReady fires — wait on that deterministic signal instead of
    // a fixed timeout so the test does not flake on a loaded CI runner.
    await expect(page.locator("#app-body")).toBeVisible({ timeout: 5_000 });

    // Settings toggle button should now have a click handler wired by bootstrap.
    await page.locator("#settings-toggle").click({ timeout: 5_000 });

    // Settings panel becomes visible.
    await expect(page.locator("#settings-view")).toBeVisible({ timeout: 5_000 });
    // The no-API-key warning is expected app output (no key is provisioned in
    // this mock), not a load/CSP defect — filter it before asserting clean.
    const real = errors.filter((e) => !e.includes("Please enter your Z.AI API key"));
    expect(real, real.join("\n")).toEqual([]);
  });
});
