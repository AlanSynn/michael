/* global Office, console */

// Thin async wrappers over Office.context.mailbox.item, plus flat event
// registration helpers. This is the one place the rest of the app reaches
// into the mailbox, so it is easy to stub for tests/mocks.

function mailbox() {
  return (typeof Office !== "undefined" && Office?.context?.mailbox) || null;
}

function item() {
  return mailbox()?.item || null;
}

/** True when a mailbox item with a readable body is selected. */
export function hasSelectedItem() {
  const current = item();
  return Boolean(current && current.body && typeof current.body.getAsync === "function");
}

// Per-itemId in-flight body cache. On ItemChanged the handler dispatches
// autorun + the calendar probe synchronously back-to-back, so both reach
// getEmailContent BEFORE either body.getAsync callback has resolved. Caching
// the Promise (not the value) collapses those concurrent calls into a single
// host round-trip, and also serves sequential re-reads (a manual button click
// on the same item). Read-mode bodies are stable per itemId; the cap bounds
// memory if the user clicks through many items.
const emailContentCache = new Map(); // itemId -> Promise<string>
const EMAIL_CACHE_MAX = 20;

function currentItemId() {
  try {
    return (Office.context.mailbox.item && Office.context.mailbox.item.itemId) || null;
  } catch {
    return null;
  }
}

function noItemError() {
  // Distinguish "no item selected / no body" (a benign state — the user is on
  // the mailbox list) from a real fetch failure, so callers can stay quiet
  // instead of toasting an error on every scroll-through.
  const err = new Error("No email selected.");
  err.code = "NO_ITEM";
  return err;
}

/** Drop every cached body. Call on ItemChanged so a stale body never leaks
 *  forward (and so the next selection re-fetches against the live host). */
export function clearEmailContentCache() {
  emailContentCache.clear();
}

/** @returns {Promise<string>} the current item body as plain text */
export function getEmailContent() {
  const current = item();
  if (!current?.body?.getAsync) {
    return Promise.reject(noItemError());
  }
  const itemId = currentItemId();
  if (itemId) {
    const cached = emailContentCache.get(itemId);
    if (cached) {
      return cached;
    }
  }
  const promise = new Promise((resolve, reject) => {
    current.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email content"));
      }
    });
  });
  if (itemId) {
    if (emailContentCache.size >= EMAIL_CACHE_MAX) {
      emailContentCache.delete(emailContentCache.keys().next().value);
    }
    emailContentCache.set(itemId, promise);
    // A rejected promise must not stay cached (a transient host failure should
    // retry on the next call, not replay the rejection forever). The .catch
    // also absorbs the rejection so an unawaited concurrent caller does not
    // surface an unhandled-rejection warning.
    promise.catch(() => {
      if (emailContentCache.get(itemId) === promise) {
        emailContentCache.delete(itemId);
      }
    });
  }
  return promise;
}

/** @returns {string} the current item subject (empty string if unavailable) */
export function getSubject() {
  const current = item();
  const subject = current?.subject;
  return typeof subject === "string" ? subject : "";
}

/** Replace the current selection with text. Used by the function-file command. */
export function replaceSelectionWithText(textToInsert) {
  return new Promise((resolve, reject) => {
    const current = item();
    if (!current?.body?.setSelectedDataAsync) {
      reject(new Error("Failed to insert/replace text: mailbox not available"));
      return;
    }
    current.body.setSelectedDataAsync(
      textToInsert,
      { coercionType: Office.CoercionType.Text },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Failed to insert/replace text: ${asyncResult.error.message}`));
          return;
        }
        resolve();
      }
    );
  });
}

/**
 * Register a flat ItemChanged handler (no nesting). Fixes the prior broken
 * result.value.addHandlerAsync / register() chains. Guarded because some
 * classic hosts do not expose addHandlerAsync. The async callback logs a
 * failure instead of silently swallowing it (a dead registration otherwise
 * leaves autorun + calendar-button updates quietly inactive).
 */
export function onItemChanged(handler) {
  const mb = mailbox();
  if (!mb?.addHandlerAsync) {
    return;
  }
  mb.addHandlerAsync(Office.EventType.ItemChanged, handler, (result) => {
    if (result && result.status === Office.AsyncResultStatus.Failed) {
      console.error("ItemChanged handler registration failed:", result.error);
    }
  });
}

export function onSettingsChanged(handler) {
  const mb = mailbox();
  if (!mb?.addHandlerAsync) {
    return;
  }
  mb.addHandlerAsync(Office.EventType.SettingsChanged, handler, (result) => {
    if (result && result.status === Office.AsyncResultStatus.Failed) {
      console.error("SettingsChanged handler registration failed:", result.error);
    }
  });
}
