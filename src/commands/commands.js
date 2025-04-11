/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Safety settings for Gemini API
const safetySettings = [
  {
    category: "HARM_CATEGORY_HARASSMENT",
    threshold: "BLOCK_NONE",
  },
  {
    category: "HARM_CATEGORY_HATE_SPEECH",
    threshold: "BLOCK_NONE",
  },
  {
    category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
    threshold: "BLOCK_NONE",
  },
  {
    category: "HARM_CATEGORY_DANGEROUS_CONTENT",
    threshold: "BLOCK_NONE",
  },
];

// Fixed model for this command
const MODEL_NAME = "gemini-2.0-flash-lite";

Office.onReady(() => {
  // Office is ready
});

/**
 * Handles the Add-in Command button click.
 * Translates the entire email body to Korean and replaces the current selection (or inserts at cursor).
 * @param {Office.AddinCommands.Event} event The event object.
 */
async function action(event) {
  let apiKey = "";

  // 1. Get API Key from Roaming Settings
  try {
    if (Office.context.roamingSettings) {
      apiKey = Office.context.roamingSettings.get("apiKey");
    } else {
      throw new Error("RoamingSettings not available.");
    }
  } catch (error) {
    console.error("Error accessing RoamingSettings:", error);
    showErrorNotification(`Error accessing settings: ${error.message}`, event);
    return;
  }

  if (!apiKey) {
    showErrorNotification("Please set up your Gemini API key in the ReadMeDarling taskpane settings first.", event);
    return;
  }

  // 2. Show Processing Notification
  showProcessingNotification("Translating email body to Korean...", event);

  try {
    // 3. Get Email Content
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject; // Subject might be useful context

    // 4. Prepare the Prompt for Full Translation
    const prompt = `Translate the following email content entirely into Korean. Preserve the original meaning and tone as much as possible.

Email Subject (for context): ${subject}
Email Content:
---
${emailContent}
---

Provide only the translated Korean text. Do not add any introductory or concluding remarks.`;

    // 5. Generate Content using Gemini API
    const translatedBody = await generateContent(prompt, apiKey, MODEL_NAME);

    // 6. Replace the current selection (or insert at cursor) with the translation
    await replaceSelectionWithText(translatedBody, event);

    // 7. Show Success Notification
    showSuccessNotification("Email body translated to Korean and replaced selection/inserted at cursor.", event);

  } catch (error) {
    console.error("Error during translation command:", error);
    showErrorNotification(`Translation failed: ${error.message}`, event);
    // Ensure event.completed is called even after error notification
    event.completed();
  }
}

// --- Helper Functions ---

/**
 * Replaces the currently selected text in the email body with the provided text,
 * or inserts at the cursor if nothing is selected.
 * @param {string} textToInsert The text to insert/replace.
 * @param {Office.AddinCommands.Event} event The event object for completion.
 */
async function replaceSelectionWithText(textToInsert, event) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert,
      { coercionType: Office.CoercionType.Text, asyncContext: { event: event } },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set selected data:", asyncResult.error);
          reject(new Error(`Failed to insert/replace text: ${asyncResult.error.message}`));
        } else {
          console.log("Selected data replaced/inserted successfully.");
          resolve();
        }
      }
    );
  });
}

/**
 * Gets the email body content as plain text.
 * @returns {Promise<string>}
 */
async function getEmailContent() {
  return new Promise((resolve, reject) => {
    // IMPORTANT: Get the body as TEXT for translation, regardless of the original format.
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value.trim()); // Trim whitespace
        } else {
          console.error("Failed to get email body as text:", result.error);
          reject(new Error("Failed to get email content"));
        }
      }
    );
  });
}

/**
 * Generates content using the Gemini API.
 * @param {string} prompt The prompt for the API.
 * @param {string} apiKey The API key.
 * @param {string} modelName The Gemini model name.
 * @returns {Promise<string>} The generated text content.
 */
async function generateContent(prompt, apiKey, modelName) {
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        contents: [
          {
            parts: [
              {
                text: prompt,
              },
            ],
          },
        ],
        safetySettings: safetySettings,
        generationConfig: {
          temperature: 0.3,
          maxOutputTokens: 8192,
        },
      }),
    });

    const data = await response.json();

    if (!response.ok || data.error) {
      throw new Error(data.error?.message || `API Error (${response.status})`);
    }

    if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts[0]) {
      return data.candidates[0].content.parts[0].text.trim();
    } else if (data.candidates && data.candidates[0]?.finishReason === 'SAFETY') {
      throw new Error("Content generation blocked due to safety settings.");
    } else if (data.promptFeedback?.blockReason) {
        throw new Error(`Prompt blocked due to safety settings: ${data.promptFeedback.blockReason}`);
    } else {
      console.warn("Unexpected API response structure:", data);
      throw new Error("No content generated or unexpected format.");
    }
  } catch (error) {
    console.error("Error generating content via command:", error);
    throw error; // Re-throw for handling in the action function
  }
}

// --- Notification Helpers ---

function showProcessingNotification(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ProcessingNotification",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "Icon.80x80", // Make sure this icon is defined in your manifest resources
      persistent: false,
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to show processing notification: ", asyncResult.error);
      }
    }
  );
}

function showSuccessNotification(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionCompleteNotification",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "Icon.80x80",
      persistent: true, // Keep success message visible
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
      message: message,
      icon: "Icon.80x80",
      persistent: true,
    },
    (asyncResult) => {
      handleNotificationResult(asyncResult, event, "error");
    }
  );
}

// Common handler for notification results; crucial for calling event.completed
function handleNotificationResult(asyncResult, event, type) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.error(`Failed to show ${type} notification: `, asyncResult.error);
  }
  // Ensure event.completed() is called AFTER the notification attempt, regardless of success/failure
  // except when the error handler already called it.
  if (type !== 'error') { // Avoid double completion if error handler already called it
      event.completed();
  }
}

// Register the function with Office.
Office.actions.associate("action", action);