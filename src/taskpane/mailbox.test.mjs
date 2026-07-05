import { test } from "node:test";
import assert from "node:assert/strict";

// mailbox.js reaches into Office.context.mailbox.item via helpers that re-read
// the global on every call (no import-time capture), so swapping globalThis.Office
// between tests is sufficient — no fresh import needed.

function setOffice(item) {
  globalThis.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Text: "Text" },
    EventType: { ItemChanged: "itemChanged", SettingsChanged: "settingsChanged" },
    context: { mailbox: { item } },
  };
}

const { hasSelectedItem, getEmailContent } = await import("./mailbox.js");

test("hasSelectedItem: true when item exposes body.getAsync", () => {
  setOffice({ body: { getAsync() {} } });
  assert.equal(hasSelectedItem(), true);
});

test("hasSelectedItem: false when no item is selected", () => {
  setOffice(null);
  assert.equal(hasSelectedItem(), false);
});

test("hasSelectedItem: false when item has no readable body", () => {
  setOffice({ subject: "no body here" });
  assert.equal(hasSelectedItem(), false);
});

test("getEmailContent: rejects with NO_ITEM when nothing is selected", async () => {
  setOffice(null);
  await assert.rejects(getEmailContent(), (err) => err.code === "NO_ITEM");
});

test("getEmailContent: rejects with NO_ITEM when item has no body", async () => {
  setOffice({ subject: "x" });
  await assert.rejects(getEmailContent(), (err) => err.code === "NO_ITEM");
});

test("getEmailContent: resolves the body text on success (coerces to text)", async () => {
  setOffice({
    body: {
      getAsync(coercion, cb) {
        assert.equal(coercion, "text");
        cb({ status: "succeeded", value: "hello body" });
      },
    },
  });
  const text = await getEmailContent();
  assert.equal(text, "hello body");
});

test("getEmailContent: rejects when getAsync reports failure", async () => {
  setOffice({
    body: {
      getAsync(_coercion, cb) {
        cb({ status: "failed", error: { message: "boom" } });
      },
    },
  });
  await assert.rejects(getEmailContent(), /Failed to get email content/);
});
