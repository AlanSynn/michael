import { test } from "node:test";
import assert from "node:assert/strict";
import { getCurrentDateContext } from "./date-context.js";

test("getCurrentDateContext emits the local date, weekday, time, and a UTC offset", () => {
  // new Date(2026, 6, 5, 14, 30) = 2026-07-05 14:30 in the RUNTIME's local tz.
  // Assert structure, not the exact offset/weekday, so the test is tz-stable.
  const ctx = getCurrentDateContext(new Date(2026, 6, 5, 14, 30));

  assert.match(ctx, /Today is \w+, 2026-07-05\./);
  assert.match(ctx, /Current local time: 14:30/);
  assert.match(ctx, /UTC[+-]\d{2}:\d{2}/);
  // The tz name sits in parentheses (IANA name or the "local" fallback).
  assert.match(ctx, /\([^)]+, UTC[+-]\d{2}:\d{2}\)\./);
});

test("getCurrentDateContext reflects the injected date (not a module-level cache)", () => {
  const jan = getCurrentDateContext(new Date(2026, 0, 1, 9, 0));
  const dec = getCurrentDateContext(new Date(2026, 11, 31, 23, 59));
  assert.match(jan, /2026-01-01/);
  assert.match(jan, /09:00/);
  assert.match(dec, /2026-12-31/);
  assert.match(dec, /23:59/);
});

test("getCurrentDateContext defaults to the current time when called with no arg", () => {
  // Smoke: no-arg call must not throw and must still match the shape.
  const ctx = getCurrentDateContext();
  assert.match(ctx, /Today is \w+, \d{4}-\d{2}-\d{2}\./);
  assert.match(ctx, /UTC[+-]\d{2}:\d{2}/);
});
