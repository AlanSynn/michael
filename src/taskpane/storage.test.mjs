import { test } from "node:test";
import assert from "node:assert/strict";
import { JSDOM } from "jsdom";

// storage.js uses bare localStorage/sessionStorage globals + Office.roamingSettings.
// Provide a DOM so those globals are real, then stub the roaming store.
const dom = new JSDOM("<!DOCTYPE html><html><body></body></html>", { url: "https://localhost/" });
globalThis.window = dom.window;
globalThis.localStorage = dom.window.localStorage;
globalThis.sessionStorage = dom.window.sessionStorage;

const {
  migrateSettingsKeys,
  SETTINGS_KEY,
  TEMPLATE_DEFAULTS_KEY,
  MODEL_CATALOG_CACHE_KEY,
} = await import("./storage.js");

let store;
let saveCalls;

function installOffice() {
  store = {};
  saveCalls = 0;
  globalThis.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: {
      roamingSettings: {
        get: (k) => (k in store ? store[k] : null),
        set: (k, v) => {
          store[k] = v;
        },
        remove: (k) => {
          delete store[k];
        },
        saveAsync: (cb) => {
          saveCalls += 1;
          if (cb) cb({ status: "succeeded" });
        },
      },
    },
  };
}

function resetBrowserStorage() {
  localStorage.clear();
  sessionStorage.clear();
}

test("migrate: no-op when no legacy keys present — saveAsync must not fire", async () => {
  installOffice();
  resetBrowserStorage();
  await migrateSettingsKeys();
  assert.equal(saveCalls, 0, "saveAsync must not round-trip when nothing migrated");
});

test("migrate: copies legacy localStorage key into roaming, clears it, saves once", async () => {
  installOffice();
  resetBrowserStorage();
  localStorage.setItem("michael_settings", JSON.stringify({ apiKey: "legacy" }));

  await migrateSettingsKeys();

  assert.deepEqual(JSON.parse(store[SETTINGS_KEY]), { apiKey: "legacy" });
  assert.equal(localStorage.getItem("michael_settings"), null);
  assert.equal(saveCalls, 1);
});

test("migrate: does not overwrite an existing roaming key, but still clears legacy", async () => {
  installOffice();
  resetBrowserStorage();
  store[SETTINGS_KEY] = JSON.stringify({ apiKey: "existing" });
  localStorage.setItem("michael_settings", JSON.stringify({ apiKey: "legacy" }));

  await migrateSettingsKeys();

  assert.deepEqual(JSON.parse(store[SETTINGS_KEY]), { apiKey: "existing" });
  assert.equal(localStorage.getItem("michael_settings"), null);
});

test("migrate: copies legacy sessionStorage settings key (primary legacy path)", async () => {
  installOffice();
  resetBrowserStorage();
  sessionStorage.setItem("my_sidekick_michael_settings", JSON.stringify({ apiKey: "ss" }));

  await migrateSettingsKeys();

  assert.deepEqual(JSON.parse(store[SETTINGS_KEY]), { apiKey: "ss" });
  assert.equal(sessionStorage.getItem("my_sidekick_michael_settings"), null);
});

test("migrate: copies legacy template defaults + model catalog in a single save", async () => {
  installOffice();
  resetBrowserStorage();
  sessionStorage.setItem("michael_session_template_defaults", "{}");
  sessionStorage.setItem(MODEL_CATALOG_CACHE_KEY, "[]");

  await migrateSettingsKeys();

  assert.equal(store[TEMPLATE_DEFAULTS_KEY], "{}");
  assert.equal(store[MODEL_CATALOG_CACHE_KEY], "[]");
  assert.equal(saveCalls, 1);
});

test("migrate: sessionStorage wins when BOTH legacy settings keys exist (no clobber)", async () => {
  // Regression guard: the localStorage branch must not overwrite a value just
  // migrated from the (primary) sessionStorage path. Previously the second
  // branch reused a stale `currentSettings` snapshot and silently lost every
  // field that existed only in the session blob (templates, etc.).
  installOffice();
  resetBrowserStorage();
  sessionStorage.setItem(
    "my_sidekick_michael_settings",
    JSON.stringify({ apiKey: "ss", templates: { x: 1 } })
  );
  localStorage.setItem("michael_settings", JSON.stringify({ apiKey: "ls" }));

  await migrateSettingsKeys();

  const migrated = JSON.parse(store[SETTINGS_KEY]);
  assert.equal(migrated.apiKey, "ss", "sessionStorage (primary) must win");
  assert.deepEqual(migrated.templates, { x: 1 }, "session-only fields must survive");
  assert.equal(sessionStorage.getItem("my_sidekick_michael_settings"), null);
  assert.equal(localStorage.getItem("michael_settings"), null);
  assert.equal(saveCalls, 1);
});
