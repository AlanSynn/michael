import { test } from "node:test";
import assert from "node:assert/strict";
import { beginFlow, isActiveFlow, isCancelError } from "./flow-control.js";

test("beginFlow returns a fresh, un-aborted signal each call", () => {
  const signal = beginFlow();
  assert.equal(signal.aborted, false);
  assert.equal(isActiveFlow(signal), true);
});

test("starting a new flow aborts the previous flow's signal", () => {
  const first = beginFlow();
  const second = beginFlow();
  // The newer flow owns the UI; the older signal is aborted + no longer active.
  assert.equal(first.aborted, true);
  assert.equal(isActiveFlow(first), false);
  assert.equal(isActiveFlow(second), true);
});

test("isCancelError recognizes AbortError and rejects other values", () => {
  const abortErr = new Error("cancelled");
  abortErr.name = "AbortError";
  assert.equal(isCancelError(abortErr), true);
  assert.equal(isCancelError(new Error("boom")), false);
  assert.equal(isCancelError(null), false);
  assert.equal(isCancelError(undefined), false);
});
