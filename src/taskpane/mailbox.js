// Thin async wrappers over Office.context.mailbox.item, plus flat event
// registration helpers. This is the one place the rest of the app reaches
// into the mailbox, so it is easy to stub for tests/mocks.

function mailbox() {
  return (typeof Office !== "undefined" && Office?.context?.mailbox) || null;
}

function item() {
  return mailbox()?.item || null;
}

/** @returns {Promise<string>} the current item body as plain text */
export function getEmailContent() {
  return new Promise((resolve, reject) => {
    const current = item();
    if (!current?.body?.getAsync) {
      reject(new Error("Failed to get email content"));
      return;
    }
    current.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email content"));
      }
    });
  });
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
 * classic hosts do not expose addHandlerAsync.
 */
export function onItemChanged(handler) {
  const mb = mailbox();
  if (!mb?.addHandlerAsync) {
    return;
  }
  mb.addHandlerAsync(Office.EventType.ItemChanged, handler);
}

export function onSettingsChanged(handler) {
  const mb = mailbox();
  if (!mb?.addHandlerAsync) {
    return;
  }
  mb.addHandlerAsync(Office.EventType.SettingsChanged, handler);
}
