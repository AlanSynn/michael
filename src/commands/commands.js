/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console, require */

const { ZAI_DEFAULT_MODEL } = require("../shared/zaiConfig");
const { executeZaiChatCompletion } = require("../shared/zaiClient");

const SETTINGS_KEY = "michael_settings";
const COMMAND_TRANSLATE_TEMPLATE_FALLBACK = `
Translate the email below into {language}.

Requirements:
- Return only the translated email body.
- Preserve meaning, tone, names, dates, numbers, and paragraph structure.
- Do not add a summary, bullets, or commentary.

Subject: {subject}
Content:
{content}`;

Office.onReady(() => {
  // Office is ready
});

/**
 * Handles the Add-in Command button click.
 * Translates the entire email body using the saved Outlook settings and inserts it at the cursor.
 * @param {Office.AddinCommands.Event} event The event object.
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

  showProcessingNotification(`Translating email body to ${targetLanguage}...`, event);

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;
    const prompt = template
      .replace("{subject}", subject)
      .replace("{content}", emailContent)
      .replace("{language}", targetLanguage);
    const translatedBody = await generateContent(prompt, model, apiKey);

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

async function replaceSelectionWithText(textToInsert) {
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

async function getEmailContent() {
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

async function generateContent(prompt, modelName, apiKey) {
  const result = await executeZaiChatCompletion({
    apiKey,
    userPrompt: prompt,
    model: modelName,
    temperature: 0.3,
  });

  return result.text;
}

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
  return configuredModel || ZAI_DEFAULT_MODEL;
}

function getSavedCommandTranslateTemplate(settings) {
  const configuredTemplate =
    typeof settings?.templates?.commandTranslate === "string"
      ? settings.templates.commandTranslate.trim()
      : "";

  if (!configuredTemplate) {
    return COMMAND_TRANSLATE_TEMPLATE_FALLBACK.trim();
  }

  return configuredTemplate;
}

function getLanguageText(languageCode) {
  switch (languageCode) {
    case "es":
      return "Spanish";
    case "fr":
      return "French";
    case "de":
      return "German";
    case "it":
      return "Italian";
    case "ja":
      return "Japanese";
    case "ko":
      return "Korean";
    case "zh_cn":
      return "Chinese";
    default:
      return "English";
  }
}

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
    (asyncResult) => {
      handleNotificationResult(asyncResult, event, "success");
    }
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
    (asyncResult) => {
      handleNotificationResult(asyncResult, event, "error");
    }
  );
}

function handleNotificationResult(asyncResult, event, type) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.error(`Failed to show ${type} notification: `, asyncResult.error);
  }

  event.completed();
}
