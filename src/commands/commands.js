/* global Office, console */

// Function-file for the "Quick Translate" ribbon command. Imports the same
// shared Z.AI client, language map, prompt builder, storage, and mailbox
// facades as the taskpane, so there is a single implementation of each.
import { generateText, getDefaultZaiModel } from "../shared/zai.js";
import { getLanguageText } from "../taskpane/language.js";
import { fillTemplate } from "../taskpane/prompts.js";
import { DEFAULT_PROMPT_TEMPLATES } from "../taskpane/prompt-templates.js";
import { getSettings } from "../taskpane/storage.js";
import { getEmailContent, replaceSelectionWithText } from "../taskpane/mailbox.js";

Office.onReady(() => {
  // Office is ready.
});

/**
 * Handle the Quick Translate command: translate the whole email body using the
 * saved Outlook settings and insert the translation at the cursor.
 * @param {Office.AddinCommands.Event} event
 */
async function action(event) {
  const settings = getSettings();
  const apiKey = getSavedApiKey(settings);
  if (!apiKey) {
    showErrorNotification("Open Michael Settings and save a Z.AI API key first.", event);
    return;
  }

  const model = getSavedModel(settings);
  const template = getSavedCommandTranslateTemplate(settings);
  const targetLanguage = getLanguageText(settings.defaultLanguage);

  try {
    showProcessingNotification(`Translating email body to ${targetLanguage}...`);

    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;
    const prompt = fillTemplate(template, {
      subject,
      content: emailContent,
      language: targetLanguage,
    });
    const translatedBody = await generateText(prompt, { apiKey, model, temperature: 0.3 });

    await replaceSelectionWithText(translatedBody);

    showSuccessNotification(
      `Email body translated to ${targetLanguage} and inserted at the cursor.`,
      event
    );
  } catch (error) {
    console.error("Error during translation command:", error);
    showErrorNotification(`Translation failed: ${error.message}`, event);
  }
}

Office.actions.associate("action", action);

// --- saved-settings readers -------------------------------------------------

function getSavedApiKey(settings) {
  return typeof settings?.apiKey === "string" ? settings.apiKey.trim() : "";
}

function getSavedModel(settings) {
  const configuredModel = typeof settings?.model === "string" ? settings.model.trim() : "";
  return configuredModel || getDefaultZaiModel();
}

function getSavedCommandTranslateTemplate(settings) {
  const configuredTemplate =
    typeof settings?.templates?.commandTranslate === "string"
      ? settings.templates.commandTranslate.trim()
      : "";

  if (!configuredTemplate) {
    return DEFAULT_PROMPT_TEMPLATES.commandTranslate.trim();
  }

  return configuredTemplate;
}

// --- notifications ----------------------------------------------------------

function showProcessingNotification(message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("ProcessingNotification", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message,
    icon: "Icon.80x80",
    persistent: false,
  });
}

function showSuccessNotification(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionCompleteNotification",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.80x80",
      persistent: true,
    },
    (asyncResult) => handleNotificationResult(asyncResult, event, "success")
  );
}

function showErrorNotification(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionErrorNotification",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message,
      icon: "Icon.80x80",
      persistent: true,
    },
    (asyncResult) => handleNotificationResult(asyncResult, event, "error")
  );
}

function handleNotificationResult(asyncResult, event, type) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.error(`Failed to show ${type} notification: `, asyncResult.error);
  }

  event.completed();
}
