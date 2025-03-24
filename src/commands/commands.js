/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Safety settings for Gemini API
const safetySettings = [{
  category: "HARM_CATEGORY_HARASSMENT",
  threshold: "BLOCK_NONE"
},
{
  category: "HARM_CATEGORY_HATE_SPEECH",
  threshold: "BLOCK_NONE"
},
{
  category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
  threshold: "BLOCK_NONE"
},
{
  category: "HARM_CATEGORY_DANGEROUS_CONTENT",
  threshold: "BLOCK_NONE"
}];

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Quickly summarizes the email when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
async function action(event) {
  // Show a notification that we're working on it
  const processingMessage = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Summarizing your email...",
    icon: "Icon.80x80",
    persistent: true,
  };

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ProcessingNotification",
    processingMessage
  );

  try {
    // Get API key from settings
    const apiKey = Office.context.roamingSettings.get("apiKey");

    if (!apiKey) {
      // Show notification to set up API key
      const errorMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Please set up your Gemini API key in the add-in settings first.",
        icon: "Icon.80x80",
        persistent: true,
      };

      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "ErrorNotification",
        errorMessage
      );

      event.completed();
      return;
    }

    // Get email content
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;

    // Create prompt for summarization
    const prompt = `Summarize the following email in a concise way that captures the main points:

Subject: ${subject}

Content:
${emailContent}`;

    // Generate summary
    const summary = await generateContent(prompt, apiKey);

    // Show success notification with summary
    const successMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `Summary: ${summary.substring(0, 150)}${summary.length > 150 ? '...' : ''}`,
      icon: "Icon.80x80",
      persistent: true,
    };

    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "SummaryNotification",
      successMessage
    );
  } catch (error) {
    // Show error notification
    const errorMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `Error: ${error.message}`,
      icon: "Icon.80x80",
      persistent: true,
    };

    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "ErrorNotification",
      errorMessage
    );
  }

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Get email content
async function getEmailContent() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email content"));
      }
    });
  });
}

// Generate content using Gemini API
async function generateContent(prompt, apiKey) {
  const apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent";
  const url = `${apiUrl}?key=${apiKey}`;

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        contents: [{
          parts: [{
            text: prompt
          }]
        }],
        safetySettings: safetySettings
      })
    });

    const data = await response.json();

    if (data.error) {
      throw new Error(data.error.message);
    }

    if (data.candidates && data.candidates[0] && data.candidates[0].content) {
      return data.candidates[0].content.parts[0].text;
    } else {
      throw new Error("No content generated");
    }
  } catch (error) {
    console.error("Error generating content:", error);
    throw error;
  }
}

// Register the function with Office.
Office.actions.associate("action", action);
