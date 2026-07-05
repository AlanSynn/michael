/* global AbortController */

// Single-flight control shared by the email flows (summarize / translate /
// translate+summarize / reply) and the calendar-event parse. Only one runs at a
// time: starting a new flow aborts the previous in-flight request so a late
// result can never overwrite the current UI, and the superseded flow's closures
// are released promptly.
//
// Extracted here so flows.js and calendar.js share ONE owner. Earlier the
// calendar parse held its own implicit controller, which let a calendar parse
// and a summarize run concurrently and race on the result-section DOM.

let activeController = null;

/** Abort any in-flight flow and return a fresh signal for the new one. */
export function beginFlow() {
  if (activeController) {
    activeController.abort();
  }
  activeController = new AbortController();
  return activeController.signal;
}

/** True when the error is an AbortError from a superseded (cancelled) flow. */
export function isCancelError(error) {
  return Boolean(error) && error.name === "AbortError";
}

/**
 * True when `signal` still belongs to the active flow. A superseded flow must
 * not touch the loading spinner or results — the newer flow owns the UI now,
 * so callers bail before every DOM mutation.
 */
export function isActiveFlow(signal) {
  return activeController !== null && activeController.signal === signal;
}
