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

const { hasSelectedItem, getEmailContent, clearEmailContentCache } = await import("./mailbox.js");

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

test("getEmailContent: collapses concurrent calls for the same itemId into one getAsync", async () => {
  clearEmailContentCache();
  let getAsyncCalls = 0;
  setOffice({
    itemId: "item-A",
    body: {
      getAsync(_coercion, cb) {
        getAsyncCalls += 1;
        // Resolve asynchronously so the second concurrent caller observes the
        // in-flight promise (the documented autorun+probe collision case).
        setTimeout(() => cb({ status: "succeeded", value: "shared body" }), 0);
      },
    },
  });
  const [a, b] = await Promise.all([getEmailContent(), getEmailContent()]);
  assert.equal(a, "shared body");
  assert.equal(b, "shared body");
  assert.equal(getAsyncCalls, 1, "concurrent calls share one host round-trip");
});

test("getEmailContent: sequential calls on the same itemId hit the cache", async () => {
  clearEmailContentCache();
  let getAsyncCalls = 0;
  setOffice({
    itemId: "item-B",
    body: {
      getAsync(_coercion, cb) {
        getAsyncCalls += 1;
        cb({ status: "succeeded", value: "cached body" });
      },
    },
  });
  await getEmailContent();
  await getEmailContent();
  assert.equal(getAsyncCalls, 1, "second call served from cache");
});

test("getEmailContent: a failed fetch is dropped from the cache (retries next call)", async () => {
  clearEmailContentCache();
  let getAsyncCalls = 0;
  setOffice({
    itemId: "item-C",
    body: {
      getAsync(_coercion, cb) {
        getAsyncCalls += 1;
        cb({ status: "failed", error: { message: "transient" } });
      },
    },
  });
  await assert.rejects(getEmailContent(), /Failed to get email content/);
  await assert.rejects(getEmailContent(), /Failed to get email content/);
  assert.equal(getAsyncCalls, 2, "rejected promise evicted so the call retries");
});
