/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console */

// Function-file for the "Quick Translate" ribbon command. Now imports the same
// shared Z.AI client, language map, and prompt builder as the taskpane, so there
// is a single provider implementation (the prior CJS zaiClient/zaiConfig were
// ~40% duplicated).
import { generateText, getDefaultZaiModel } from "../shared/zai.js";
import { getLanguageText } from "../taskpane/language.js";
import { fillTemplate } from "../taskpane/prompts.js";
import { DEFAULT_PROMPT_TEMPLATES } from "../taskpane/prompt-templates.js";

const SETTINGS_KEY = "michael_settings";

Office.onReady(() => {
  // Office is ready.
});

/**
 * Handle the Quick Translate command: translate the whole email body using the
 * saved Outlook settings and insert the translation at the cursor.
 * @param {Office.AddinCommands.Event} event
 */
async function action(event) {
  const settings = getSavedSettings();
  const apiKey = getSavedApiKey(settings);
  if (!apiKey) {
    showErrorNotification("Open Michael Settings and save a Z.AI API key first.", event);
    return;
  }

  const model = getSavedModel(settings);
  const template = getSavedCommandTranslateTemplate(settings);
  const targetLanguage = getLanguageText(settings.defaultLanguage);

  showProcessingNotification(`Translating email body to ${targetLanguage}...`);

  try {
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

// --- mailbox helpers (function-file context) --------------------------------

function replaceSelectionWithText(textToInsert) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert,
      { coercionType: Office.CoercionType.Text },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set selected data:", asyncResult.error);
          reject(new Error(`Failed to insert/replace text: ${asyncResult.error.message}`));
          return;
        }
        resolve();
      }
    );
  });
}

function getEmailContent() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.trim());
        return;
      }
      console.error("Failed to get email body as text:", result.error);
      reject(new Error("Failed to get email content."));
    });
  });
}

// --- saved-settings readers -------------------------------------------------

function getSavedSettings() {
  try {
    const rawValue = Office.context?.roamingSettings?.get(SETTINGS_KEY);
    if (!rawValue || typeof rawValue !== "string") {
      return {};
    }
    const parsed = JSON.parse(rawValue);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (error) {
    console.error("Failed to read saved Outlook settings:", error);
    return {};
  }
}

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
