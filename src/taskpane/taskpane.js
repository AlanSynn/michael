/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import { marked } from 'marked';

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

const TYPES = {
  SUMMARIZE: 0,
  TRANSLATE: 1,
  TRANSLATE_SUMMARIZE: 2,
  REPLY: 3
};

// Default templates
const DEFAULT_TEMPLATES = {
  summarize: `You are an expert researcher. Your task is to review the provided research document and complete the following tasks with academic depth and varied sentence structures:

1. Summarize the Research Background:
   - Provide a concise yet comprehensive summary of the study's background, including its context and motivation.

2. Extract the Problem Statement:
   - Clearly identify and articulate the central problem or challenge addressed by the research.

3. Identify Strengths:
   - List 3–5 key strengths of the study, focusing on aspects such as methodology, innovation, robustness of results, or any other notable positive attributes.

4. Identify Weaknesses:
   - Enumerate 4–5 significant weaknesses or limitations present in the research, considering issues like methodological gaps, limited scope, or areas lacking clarity.

5. Propose Research Topics:
   - Based on the weaknesses identified, suggest three potential research topics that could address these limitations or explore related areas further.

6. Include a "tl;dr" Section:
   - At the very top of your response, provide a succinct "tl;dr" summary without any syntax highlighting that encapsulates the main points and any necessary takeaways.

Do not include any extraneous messages, introductions, or commentary. Your final output should strictly adhere to these instructions.

Subject: {subject}
Content:
{content}`,
  translate: `You are an expert translator and interpreter with extensive proficiency in various languages, specializing in translating texts into polished, academic Korean. Your task is to translate the provided text from the source language into Korean, ensuring that every nuance, stylistic detail, and analytical aspect is accurately and naturally conveyed. Please follow these guidelines:

Preserve Nuance and Style:
- Accurately reflect the original text's tone, emotional nuance, and stylistic characteristics in Korean.
- Adapt idiomatic expressions, metaphors, and culturally specific references to ensure they resonate with Korean readers.

Maintain Analytical Precision:
- Carefully dissect complex sentences and ideas, ensuring that your translation maintains the original text's logical flow and depth of analysis.
- Where necessary, integrate brief annotations or contextual clarifications to help convey any cultural or conceptual subtleties.

Ensure Accuracy and Consistency:
- Translate specialized vocabulary, technical terms, and academic language with precision and maintain consistency throughout the text.
- Verify that the structure and argumentative progression of the source material are preserved in the Korean version.

Uphold Contextual Integrity:
- Ensure that the overall message and intent of the original text are fully maintained in your translation.
- Make sure that transitions between ideas and sections remain coherent and logically connected in Korean.

Review and Refine:
- Reassess your translation for any potential ambiguities or loss of nuance, refining as necessary to enhance clarity and precision.
- Strive for a balanced outcome that honors the original text while ensuring the translation is engaging and accessible to a Korean audience.

Provide tl;dr at the top:
Write a tl;dr section at the top of your response that summarizes the main points and todos if needed.

Deliver your final translation in refined, academic Korean that faithfully embodies the original text's analytical and stylistic essence.

Subject: {subject}

Content:
{content}`,
  translateSummarize: `You are an expert translator and summarizer with extensive proficiency in various languages, specializing in translating texts into polished, academic Korean. Your task is to translate the provided text from the source language into Korean AND create a concise summary of the main points. Please follow these guidelines:

Translation Aspects:
- Accurately reflect the original text's tone, emotional nuance, and stylistic characteristics in Korean.
- Adapt idiomatic expressions, metaphors, and culturally specific references to ensure they resonate with Korean readers.
- Maintain analytical precision and logical flow in the Korean translation.

Summarization Requirements:
- Create a concise and focused summary of the key points in Korean.
- Prioritize clarity and brevity while maintaining the essential meaning.
- Ensure the summary captures the main ideas, arguments, and conclusions.
- Limit the summary to approximately 30-40% of the original length.

Final Delivery Format:
1. First provide tl;dr at the top:
Write a tl;dr section at the top of your response that summarizes the main points and todos if needed.
2. Then provide a concise summary section (제목: 요약)
3. Then provide the full translation (제목: 전체 번역)

Subject: {subject}

Content:
{content}`,
  reply: `You are an expert assistant that can help me with my email. I will provide you with the email content and you will help me with the following tasks:

  Write a reply to the email.

  Subject: {subject}

  Content:
  {content}
  in {language}`,
  tldrPrompt: `Please provide a very concise summary in {language} of the following content. Focus on the main points and key takeaways. Do not include any extraneous messages, introductions, or commentary. Your final output should strictly adhere to these instructions.

  Subject: {subject}

  Content:
  {content}`
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Initialize the application
    initializeApp();

    // Add event listeners for the application buttons
    document.getElementById("summarize").addEventListener("click", summarizeEmail);
    document.getElementById("translate").addEventListener("click", translateEmail);
    document.getElementById("translate-summarize").addEventListener("click", translateAndSummarizeEmail);
    document.getElementById("settings-toggle").addEventListener("click", toggleSettingsDropdown);
    document.getElementById("dropdown-close-settings").addEventListener("click", toggleSettingsDropdown);
    document.getElementById("dropdown-save-settings").addEventListener("click", saveDropdownSettings);
    document.getElementById("dropdown-reset-templates").addEventListener("click", resetTemplates);
    document.getElementById("dropdown-reset-all").addEventListener("click", resetAllSettings);
    document.getElementById("dropdown-api-key").addEventListener("keypress", (event) => {
      if (event.key === "Enter") {
        saveDropdownSettings();
      }
    });

    // Settings selection change listeners
    document.querySelectorAll(".settings-dropdown-container select").forEach(select => {
      select.addEventListener("change", saveDropdownSettings);
    });

    // Expand button listener
    document.getElementById("expand-content").addEventListener("click", expandContent);

    // Copy buttons listeners
    document.getElementById("copy-result").addEventListener("click", copyResult);
    document.getElementById("generate-reply").addEventListener("click", generateReply);

    // Load saved settings if any
    loadDropdownSettings();

    // Apply current theme
    applyCurrentTheme();

    // Register for theme change events
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.SettingsChanged,
        onSettingsChanged
      );
    }
  }
});

/**
 * Toggle the visibility of the settings dropdown
 */
function toggleSettingsDropdown() {
  const dropdown = document.getElementById("settings-dropdown");
  if (dropdown) {
    const isVisible = dropdown.style.display === "block";
    dropdown.style.display = isVisible ? "none" : "block";
  }
}

// Make the function available globally
window.toggleSettingsDropdown = toggleSettingsDropdown;

/**
 * Apply the current theme based on user preference or Office theme
 */
function applyCurrentTheme() {
  // Default to dark theme instead of system
  const savedTheme = localStorage.getItem('theme') || 'dark';

  if (savedTheme === 'light') {
    document.body.setAttribute('data-theme', 'light');
    document.body.classList.remove('dark-theme');
  } else if (savedTheme === 'dark') {
    document.body.setAttribute('data-theme', 'dark');
    document.body.classList.add('dark-theme');
  } else {
    // Use Office theme
    if (Office.context.officeTheme) {
      const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
      // Only call isDarkTheme if bodyBackgroundColor exists
      if (bodyBackgroundColor) {
        if (isDarkTheme(bodyBackgroundColor)) {
          document.body.setAttribute('data-theme', 'dark');
          document.body.classList.add('dark-theme');
        } else {
          document.body.setAttribute('data-theme', 'light');
          document.body.classList.remove('dark-theme');
        }
      } else {
        // Default to dark theme if no color information
        document.body.setAttribute('data-theme', 'dark');
        document.body.classList.add('dark-theme');
      }
    }
  }
}

/**
 * Determine if a color is dark by converting it to RGB and calculating perceived brightness
 * @param {string} color - Hex color code
 * @returns {boolean} - True if the color is dark
 */
function isDarkTheme(color) {
  // If color is undefined or not a string, default to light theme
  if (!color || typeof color !== 'string') {
    return false;
  }

  try {
    // Convert hex to RGB
    color = color.replace('#', '');
    const r = parseInt(color.substr(0, 2), 16);
    const g = parseInt(color.substr(2, 2), 16);
    const b = parseInt(color.substr(4, 2), 16);

    // Check if we got valid RGB values
    if (isNaN(r) || isNaN(g) || isNaN(b)) {
      return false;
    }

    // Calculate perceived brightness using the formula: (0.299*R + 0.587*G + 0.114*B)
    const brightness = (0.299 * r + 0.587 * g + 0.114 * b);

    // If brightness is less than 128, consider it dark
    return brightness < 128;
  } catch (error) {
    console.error("Error calculating theme brightness:", error);
    return false;
  }
}

/**
 * Handle Office theme change event
 */
function onSettingsChanged(eventArgs) {
  const savedTheme = localStorage.getItem('theme') || 'system';
  if (savedTheme === 'system') {
    applyCurrentTheme();
  }
}

/**
 * Show notification message
 */
function showNotification(message, type = "info") {
  // Remove any existing notification
  const existingNotification = document.getElementById("notification");
  if (existingNotification) {
    existingNotification.remove();
  }

  // Create new notification element
  const notification = document.createElement("div");
  notification.id = "notification";
  notification.className = `notification ${type}`;
  notification.textContent = message;

  // Add to document
  document.body.appendChild(notification);

  // Force reflow to ensure animation plays
  notification.offsetHeight;

  // Add slide-in animation
  notification.style.animation = "slideInFromTop 0.3s ease-out";

  // Set timeout to remove notification
  setTimeout(() => {
    // Add slide-out animation
    notification.style.animation = "slideOutToTop 0.3s ease-out";

    // Remove element after animation
    setTimeout(() => {
      notification.remove();
    }, 300);
  }, 3000);
}

/**
 * Save settings from the dropdown menu
 */
function saveDropdownSettings() {
  // Get existing settings
  let settings = {};
  try {
    const savedSettings = localStorage.getItem('michael_settings');
    if (savedSettings) {
      settings = JSON.parse(savedSettings);
    }
  } catch (error) {
    console.error('Error parsing settings:', error);
  }

  // Update with values from dropdown form
  const apiKey = document.getElementById('dropdown-api-key').value;
  const model = document.getElementById('dropdown-model').value;
  const language = document.getElementById('dropdown-language').value;
  const theme = document.getElementById('dropdown-theme').value;
  const fontSize = document.getElementById('dropdown-font-size').value;
  const tldrMode = document.getElementById('dropdown-tldr-mode').value;
  const showReply = document.getElementById('dropdown-show-reply').value;
  const replyModel = document.getElementById('dropdown-reply-model').value;
  const summarizeTemplate = document.getElementById('dropdown-summarize-template').value;
  const translateTemplate = document.getElementById('dropdown-translate-template').value;
  const translateSummarizeTemplate = document.getElementById('dropdown-translate-summarize-template').value;

  // Update settings object
  settings.apiKey = apiKey;
  settings.model = model;
  settings.defaultLanguage = language;
  settings.theme = theme;
  settings.fontSize = fontSize;
  settings.tldrMode = tldrMode;
  settings.showReply = showReply;
  settings.replyModel = replyModel;

  // Make sure templates object exists
  if (!settings.templates) {
    settings.templates = {};
  }

  settings.templates.summarize = summarizeTemplate;
  settings.templates.translate = translateTemplate;
  settings.templates.translateSummarize = translateSummarizeTemplate;

  // Save settings
  localStorage.setItem('michael_settings', JSON.stringify(settings));

  // Apply theme if changed
  localStorage.setItem('theme', theme);
  applyCurrentTheme();

  // Apply font size
  applyFontSize(fontSize);

  // Update UI for reply button
  updateReplyButtonVisibility(showReply === 'true');

  showNotification('All settings saved successfully');

  // Close the dropdown
  toggleSettingsDropdown();
}

/**
 * Reset template fields to defaults
 */
function resetTemplates() {
  document.getElementById('dropdown-summarize-template').value = DEFAULT_TEMPLATES.summarize;
  document.getElementById('dropdown-translate-template').value = DEFAULT_TEMPLATES.translate;
  // Add the new translateSummarize template to the reset function
  if (document.getElementById('dropdown-translate-summarize-template')) {
    document.getElementById('dropdown-translate-summarize-template').value = DEFAULT_TEMPLATES.translateSummarize;
  }
  showNotification('Templates reset to defaults');
}

/**
 * Load saved settings to the dropdown form fields
 */
function loadDropdownSettings() {
  try {
    const savedSettings = localStorage.getItem('michael_settings');

    if (savedSettings) {
      const settings = JSON.parse(savedSettings);

      // Set form values
      if (settings.apiKey) document.getElementById('dropdown-api-key').value = settings.apiKey;
      if (settings.model) document.getElementById('dropdown-model').value = settings.model;
      if (settings.defaultLanguage) document.getElementById('dropdown-language').value = settings.defaultLanguage;
      if (settings.theme) document.getElementById('dropdown-theme').value = settings.theme;
      if (settings.fontSize) document.getElementById('dropdown-font-size').value = settings.fontSize;
      if (settings.tldrMode) document.getElementById('dropdown-tldr-mode').value = settings.tldrMode;
      if (settings.showReply) document.getElementById('dropdown-show-reply').value = settings.showReply;
      if (settings.replyModel) document.getElementById('dropdown-reply-model').value = settings.replyModel;

      // Apply font size if saved
      if (settings.fontSize) {
        applyFontSize(settings.fontSize);
      }

      // Update reply button visibility
      if (settings.showReply) {
        updateReplyButtonVisibility(settings.showReply === 'true');
      }

      // Set templates
      if (settings.templates) {
        if (settings.templates.summarize) {
          document.getElementById('dropdown-summarize-template').value = settings.templates.summarize;
        }
        if (settings.templates.translate) {
          document.getElementById('dropdown-translate-template').value = settings.templates.translate;
        }
        if (settings.templates.translateSummarize) {
          document.getElementById('dropdown-translate-summarize-template').value = settings.templates.translateSummarize;
        }
      }
    } else {
      // No settings found, load default templates
      resetTemplates();
    }
  } catch (error) {
    console.error('Error loading dropdown settings:', error);
    // If there's an error, reset to defaults
    resetTemplates();
  }
}

/**
 * Apply the selected font size to result content
 * @param {string} size - The font size to apply (small, medium, large)
 */
function applyFontSize(size) {
  document.documentElement.setAttribute('data-font-size', size || 'medium');
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
async function generateContent(prompt, apiKey, modelOverride = null, isTldr = false) {
    // Get model from settings or use default
    let model = "gemini-2.0-flash-light"; // Default model

    if (modelOverride) {
        model = modelOverride;
    } else {
        try {
            const savedSettings = localStorage.getItem('michael_settings');
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                if (settings.model) {
                    model = settings.model;
                }
            }
        } catch (error) {
            console.error("Error getting model from settings:", error);
        }
    }

    // API URL with selected model
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                contents: [
                    {
                        parts: [
                            {
                                text: prompt
                            }
                        ]
                    }
                ],
                generationConfig: {
                    temperature: 0.4,
                    topK: 32,
                    topP: 0.95,
                    maxOutputTokens: isTldr ? 800 : 8192, // Limit tokens for TL;DR
                }
            })
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error?.message || 'Error generating content');
        }

        if (!data.candidates || data.candidates.length === 0) {
            throw new Error('No content generated');
        }

        // Extract the generated text
        const generatedText = data.candidates[0].content.parts[0].text;
        return generatedText;
    } catch (error) {
        console.error('Error generating content:', error);
        throw error;
    } finally {
        // Hide loading spinner only if this is not a TL;DR request
        if (!isTldr) {
            document.getElementById("loading").style.display = "none";
        }
    }
}

// Generate TL;DR content
async function generateTldrContent(prompt, apiKey, language = "Korean",modelOverride = null) {
  const subject = Office.context.mailbox.item.subject;
  const emailContent = await getEmailContent();

  const tldrPrompt = DEFAULT_TEMPLATES.tldrPrompt
      .replace('{subject}', subject)
      .replace('{content}', emailContent)
      .replace('{language}', language);

  const content = await generateContent(tldrPrompt, apiKey, modelOverride, true);
  return content;
}

// Get language display text
function getLanguageText(languageCode) {
  switch (languageCode) {
    case "es": return "Spanish";
    case "fr": return "French";
    case "de": return "German";
    case "it": return "Italian";
    case "ja": return "Japanese";
    case "ko": return "Korean";
    case "zh_cn": return "Chinese";
    default: return "English";
  }
}

// Get API key from settings
function getApiKey() {
  const savedSettings = localStorage.getItem('michael_settings');
  if (savedSettings) {
    const settings = JSON.parse(savedSettings);
    return settings.apiKey;
  }
  return null;
}

// Get language from settings
function getLanguage() {
  try {
    const savedSettings = localStorage.getItem('michael_settings');
    if (savedSettings) {
      const settings = JSON.parse(savedSettings);
      if (settings.defaultLanguage) {
        return settings.defaultLanguage;
      }
    }
    return 'ko'; // Default to Korean
  } catch (error) {
    console.error('Error getting language:', error);
    return 'ko';
  }
}

// Summarize email
async function summarizeEmail() {
    const apiKey = getApiKey();

    if (!apiKey) {
        showNotification("Please add your Gemini API key in the settings", 'error');
        toggleSettingsDropdown(); // Open settings dropdown to prompt for API key
        return;
    }

    // Show loading UI
    showLoading("Summarizing email...");

    try {
        const emailContent = await getEmailContent();
        const subject = Office.context.mailbox.item.subject;

        // Get template from storage or use default
        let template = DEFAULT_TEMPLATES.summarize;
        try {
            const savedSettings = localStorage.getItem('michael_settings');
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                if (settings.templates && settings.templates.summarize) {
                    template = settings.templates.summarize;
                }
            }
        } catch (error) {
            console.error("Error getting template:", error);
        }

        // Replace placeholders in template
        const prompt = template
            .replace('{subject}', subject)
            .replace('{content}', emailContent);

        // Check for TL;DR mode
        let tldrMode = true;
        try {
            const savedSettings = localStorage.getItem('michael_settings');
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                if (settings.tldrMode) {
                    tldrMode = settings.tldrMode === 'true';
                }
            }
        } catch (error) {
            console.error('Error getting TLDR mode setting:', error);
        }

        if (tldrMode) {
            // Generate TL;DR first
            const tldrContent = await generateTldrContent(prompt, apiKey, "English");
            hideLoading();
            showResults(tldrContent, TYPES.SUMMARIZE);

            // Then generate full content in the background
            const fullContent = await generateContent(prompt, apiKey);

            // display notification of full content
            updateResults(fullContent);
            updateExpandButton(true);
        } else {
            // Generate full content only
            const summary = await generateContent(prompt, apiKey);
            showResults(summary, TYPES.SUMMARIZE);
        }
    } catch (error) {
        showNotification(`Error: ${error.message}`, 'error');
    }
}

// Translate email
async function translateEmail() {
    const apiKey = getApiKey();
    const language = getLanguage();

    if (!apiKey) {
        showNotification("Please add your Gemini API key in the settings", 'error');
        toggleSettingsDropdown(); // Open settings dropdown to prompt for API key
        return;
    }

    // Show loading UI
    showLoading("Translating to " + getLanguageText(language) + "...");

    try {
        const emailContent = await getEmailContent();
        const subject = Office.context.mailbox.item.subject;

        // Get template from storage or use default
        let template = DEFAULT_TEMPLATES.translate;
        try {
            const savedSettings = localStorage.getItem('michael_settings');
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                if (settings.templates && settings.templates.translate) {
                    template = settings.templates.translate;
                }
            }
        } catch (error) {
            console.error("Error getting template:", error);
        }

        // Replace placeholders in template
        const prompt = template
            .replace('{subject}', subject)
            .replace('{content}', emailContent);

        // Check for TL;DR mode
        let tldrMode = true;
        try {
            const savedSettings = localStorage.getItem('michael_settings');
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                if (settings.tldrMode) {
                    tldrMode = settings.tldrMode === 'true';
                }
            }
        } catch (error) {
            console.error('Error getting TLDR mode setting:', error);
        }

        if (tldrMode) {
            // Generate TL;DR first
            const tldrContent = await generateTldrContent(prompt, apiKey, language);
            hideLoading();
            showResults(tldrContent, TYPES.TRANSLATE);

            // Then generate full content in the background
            const fullContent = await generateContent(prompt, apiKey);

            // display notification of full content
            updateResults(fullContent);
            updateExpandButton(true);
        } else {
            // Generate full content only
            const translation = await generateContent(prompt, apiKey, language);
            showResults(translation, TYPES.TRANSLATE);
        }
    } catch (error) {
        showNotification(`Error: ${error.message}`, 'error');
    }
}

/**
 * Update the expand button text and style based on the full content display state
 */
function updateExpandButton(isFullContentVisible) {
  const expandButton = document.getElementById('expand-content');

  if (expandButton) {
    expandButton.disabled = !isFullContentVisible;
    expandButton.classList.toggle('ms-Button--disabled', !isFullContentVisible);
    expandButton.innerHTML = isFullContentVisible ? 'Show Full Content' : 'Hide Full Content';
    expandButton.classList.toggle('ms-Button--primary', isFullContentVisible);
  }
}

// Copy result to clipboard
function copyResult() {
  // Check if this is a reply (both TLDR and full content need to be combined)
  let resultContent = "";

  // Get the TLDR content (might contain the subject)
  const tldrContent = document.getElementById("tldr-content").innerText;

  // Get the full content (might contain the body)
  const fullContent = document.getElementById("result-content").innerText;

  // Check if this is a reply format
  const isReply = tldrContent.includes("Subject:") || fullContent.includes("Subject:");

  if (isReply) {
    // If it's a reply, try to extract and format subject and body properly
    let subject = "";
    let body = "";

    // Try to get subject from headings
    const headingMatch = /^(?:Subject:\s*)?(.+?)(?:\n|$)/i.exec(tldrContent);
    if (headingMatch) {
      subject = headingMatch[1].trim();
    }

    // Get the body text (prioritize full content if visible, otherwise use TLDR minus subject)
    if (document.getElementById("full-content-container").style.display !== "none") {
      body = fullContent;
    } else {
      // Attempt to remove subject line from TLDR if present
      if (headingMatch) {
        const lines = tldrContent.split("\n");
        body = lines.slice(1).join("\n").trim();
      } else {
        body = tldrContent;
      }
    }

    // Format as email reply
    resultContent = `Subject: ${subject}\n\n${body}`;
  } else {
    // For normal content, get what's visible (either TLDR or full content)
    if (document.getElementById("full-content-container").style.display !== "none") {
      resultContent = fullContent;
    } else {
      resultContent = tldrContent;
    }
  }

  navigator.clipboard.writeText(resultContent).then(() => {
    const copyStatus = document.getElementById("copy-status");
    copyStatus.textContent = "Copied!";
    setTimeout(() => {
      copyStatus.textContent = "";
    }, 2000);
  }).catch(err => {
    console.error('Could not copy text: ', err);
    showNotification("Failed to copy to clipboard", "error");
  });
}

// Convert markdown to HTML for rendering
function markdownToHtml(markdown) {
  if (!markdown) return '';

  // Get the current font size from settings
  let fontSize = 'medium';
  try {
    const savedSettings = localStorage.getItem('michael_settings');
    if (savedSettings) {
      const settings = JSON.parse(savedSettings);
      if (settings.fontSize) {
        fontSize = settings.fontSize;
      }
    }
  } catch (error) {
    console.error('Error getting font size:', error);
  }

  // Get font size values
  const fontSizeValue = fontSize === 'small' ? '13px' :
                        fontSize === 'large' ? '18px' : '16px';

  const lineHeightValue = fontSize === 'small' ? '1.5' :
                           fontSize === 'large' ? '1.7' : '1.6';

  const codeFontSize = fontSize === 'small' ? '13px' :
                        fontSize === 'large' ? '17px' : '15px';

  // Check if this is a reply format (starts with # Re: or similar)
  const isReply = /^#\s+(?:Re:|Subject:|\[Reply\]|Response:)/i.test(markdown);

  // Simple markdown to HTML conversion
  let html = markdown
    // Handle reply format headings specially
    .replace(/^#\s+(.*$)/gim, function(match, p1) {
      if (isReply) {
        return `<h1 class="reply-heading">${p1}</h1>`;
      } else {
        return `<h1>${p1}</h1>`;
      }
    })
    // Other headings
    .replace(/^###\s+(.*$)/gim, '<h3>$1</h3>')
    .replace(/^##\s+(.*$)/gim, '<h2>$1</h2>')

    // Convert bold and italic with stronger styling
    .replace(/\*\*(.*?)\*\*/gim, '<strong style="font-weight: 700;">$1</strong>')
    .replace(/\*(.*?)\*/gim, '<em style="font-style: italic;">$1</em>')
    .replace(/\_\_([^_]+)\_\_/gim, '<strong style="font-weight: 700;">$1</strong>')
    .replace(/\_([^_]+)\_/gim, '<em style="font-style: italic;">$1</em>')

    // Convert lists - updated to remove indentation
    .replace(/^\s*\n\* (.*)/gim, '<ul class="no-indent">\n<li>$1</li>')
    .replace(/^\* (.*)/gim, '<li>$1</li>')
    .replace(/^\s*\n\d+\. (.*)/gim, '<ol class="no-indent">\n<li>$1</li>')
    .replace(/^\d+\. (.*)/gim, '<li>$1</li>')

    // Convert blockquotes with improved styling
    .replace(/^\> (.*$)/gim, '<blockquote style="border-left: 4px solid var(--accent-color); padding-left: 1em; margin: 1em 0; background-color: rgba(0,0,0,0.03);">$1</blockquote>')

    // Convert code blocks with improved visibility
    .replace(/```([\s\S]*?)```/gim, '<pre style="background-color: var(--code-background); padding: 12px; border-radius: 5px; overflow-x: auto; border: 1px solid var(--border-color);"><code style="font-family: \'Courier New\', Courier, monospace; font-size: ' + codeFontSize + ';">$1</code></pre>')
    .replace(/`([^`]+)`/gim, '<code style="background-color: var(--code-background); padding: 3px 5px; border-radius: 3px; font-family: \'Courier New\', Courier, monospace; font-size: ' + codeFontSize + '; font-weight: 500;">$1</code>')

    // Convert horizontal rules
    .replace(/^\-\-\-$/gim, '<hr style="height: 2px; background-color: var(--border-color); border: 0; margin: 1.5em 0;">')

    // Convert links with better visibility
    .replace(/\[([^\]]+)\]\(([^)]+)\)/gim, '<a href="$2" target="_blank" style="color: var(--accent-color); text-decoration: none; font-weight: bold;">$1</a>')

    // Convert paragraphs - handle newlines
    .replace(/\n\s*\n/gim, '</p><p style="margin: 0.8em 0; font-size: ' + fontSizeValue + '; line-height: ' + lineHeightValue + ';">')
    .replace(/\n/gim, '<br>')

    // Wrap in paragraph if not already wrapped
    .replace(/^(.+)$/gim, '<p style="margin: 0.8em 0; font-size: ' + fontSizeValue + '; line-height: ' + lineHeightValue + ';">$1</p>');

  return html;
}

// Translate and Summarize email
async function translateAndSummarizeEmail() {
    const apiKey = getApiKey();

    if (!apiKey) {
        showNotification("Please add your Gemini API key in the settings", 'error');
        toggleSettingsDropdown(); // Open settings dropdown to prompt for API key
        return;
    }

    // Show loading UI
    showLoading("Translating and summarizing...");

    try {
        const emailContent = await getEmailContent();
        const subject = Office.context.mailbox.item.subject;
        let language = "English";

        // Get template from storage or use default
        let template = DEFAULT_TEMPLATES.translateSummarize;
        try {
            const savedSettings = localStorage.getItem('michael_settings');
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                if (settings.templates && settings.templates.translateSummarize) {
                    template = settings.templates.translateSummarize;
                }
                if (settings.language) {
                    language = settings.language;
                }
            }
        } catch (error) {
            console.error("Error getting template:", error);
        }

        // Replace placeholders in template
        const prompt = template
            .replace('{subject}', subject)
            .replace('{content}', emailContent)
            .replace('{language}', language);

        // Check for TL;DR mode
        let tldrMode = true;
        try {
            const savedSettings = localStorage.getItem('michael_settings');
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);
                if (settings.tldrMode) {
                    tldrMode = settings.tldrMode === 'true';
                }
            }
        } catch (error) {
            console.error('Error getting TLDR mode setting:', error);
        }

        if (tldrMode) {
            // Generate TL;DR first
            const tldrContent = await generateTldrContent(prompt, apiKey);
            hideLoading();
            showResults(tldrContent, TYPES.TRANSLATE_SUMMARIZE);

            // Then generate full content in the background
            const fullContent = await generateContent(prompt, apiKey);


            // display notification of full content
            updateResults(fullContent);
            updateExpandButton(true);
        } else {
            // Generate full content only
            const result = await generateContent(prompt, apiKey);
            showResults(result, TYPES.TRANSLATE_SUMMARIZE);
        }
    } catch (error) {
        showNotification(`Error: ${error.message}`, 'error');
    }
}

/**
 * Show loading indicator with message
 */
function showLoading(message = "Loading...") {
    // Show loading section
    const loadingSection = document.getElementById("loading");
    loadingSection.style.display = "block";

    // Update loading message
    const loadingMessage = document.getElementById("loading-message");
    loadingMessage.textContent = message;

    // Hide other sections
    document.getElementById("landing-screen").style.display = "none";
    document.getElementById("result-section").style.display = "none";
}

/**
 * Hide loading indicator
 */
function hideLoading() {
    const loadingSection = document.getElementById("loading");
    if (loadingSection) {
        loadingSection.style.display = "none";
    }
}

// Function to show the results
function showResults(content, type) {

    // Reset the full result content
    document.getElementById("result-content").innerHTML = "";
    // Reset the tldr content
    document.getElementById("tldr-content").innerHTML = "";

    // Reset the expand button
    const expandButton = document.getElementById("expand-content");
    if (expandButton) {
        expandButton.disabled = true;
        expandButton.classList.add("ms-Button--disabled");
        expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
        expandButton.classList.remove('ms-Button--primary');
    }

    // Hide loading section
    document.getElementById("loading").style.display = "none";

    // Show result section
    document.getElementById("result-section").style.display = "block";

    // Hide landing screen
    document.getElementById("landing-screen").style.display = "none";

    // Show the app body
    document.getElementById("app-body").style.display = "block";

    // Check for TL;DR mode
    let tldrMode = true;
    try {
        const savedSettings = localStorage.getItem('michael_settings');
        if (savedSettings) {
            const settings = JSON.parse(savedSettings);
            if (settings.tldrMode) {
                tldrMode = settings.tldrMode === 'true';
            }
        }
    } catch (error) {
        console.error('Error getting TLDR mode setting:', error);
    }

    // Update content based on TL;DR mode
    if (tldrMode) {
        // For TL;DR mode, show the quick summary first
        document.getElementById("tldr-content").innerHTML = marked.parse(content);

        // // Show loading spinner in full content section
        // const fullContentContainer = document.getElementById("full-content-container");
        // fullContentContainer.style.display = "block";
        // fullContentContainer.innerHTML = `
        //     <div class="loading-container" style="display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 20px;">
        //         <div class="spinner"></div>
        //         <p style="margin-top: 10px; color: var(--text-secondary);">Generating full content...</p>
        //     </div>
        // `;

        // Disable expand button and set to loading state
        const expandButton = document.getElementById("expand-content");
        if (expandButton) {
            expandButton.disabled = true;
            expandButton.classList.add("ms-Button--disabled");
            expandButton.innerHTML = '<span class="ms-Button-label">Loading Full Content...</span>';
        }
    } else {
        // For non-TL;DR mode, show the full content
        document.getElementById("tldr-content").innerHTML = marked.parse(content);

        // Ensure result-content element exists
        let resultContent = document.getElementById("result-content");
        if (!resultContent) {
            resultContent = document.createElement("div");
            resultContent.id = "result-content";
            document.getElementById("full-content-container").appendChild(resultContent);
        }
        resultContent.innerHTML = marked.parse(content);
        document.getElementById("full-content-container").style.display = "block";

        // Enable expand button and set to normal state
        const expandButton = document.getElementById("expand-content");
        if (expandButton) {
            expandButton.disabled = false;
            expandButton.classList.remove("ms-Button--disabled");
            expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
            expandButton.classList.add('ms-Button--primary');
        }
    }

    // Show/hide copy reply button based on type
    const copyReplyButton = document.getElementById("copy-reply");
    if (copyReplyButton) {
        copyReplyButton.style.display = type === TYPES.REPLY ? "inline-block" : "none";
    }

    // Show/hide copy result button based on type
    const copyResultButton = document.getElementById("copy-result");
    if (copyResultButton) {
        copyResultButton.style.display = type === TYPES.REPLY ? "none" : "inline-block";
    }

    // Show/hide generate reply button based on type and settings
    const generateReplyButton = document.getElementById("generate-reply");
    if (generateReplyButton) {
        const showReply = localStorage.getItem('michael_settings') &&
            JSON.parse(localStorage.getItem('michael_settings')).showReply === 'true';
        generateReplyButton.style.display = type === TYPES.REPLY ? "none" : (showReply ? "inline-block" : "none");
    }

    // Apply font size from settings
    try {
        const savedSettings = localStorage.getItem('michael_settings');
        if (savedSettings) {
            const settings = JSON.parse(savedSettings);
            if (settings.fontSize) {
                applyFontSize(settings.fontSize);
            }
            // Update reply button visibility
            if (settings.showReply) {
                updateReplyButtonVisibility(settings.showReply === 'true');
            }
        }
    } catch (error) {
        console.error('Error applying font size:', error);
    }

    // Scroll to top of result content if elements exist
    const resultContent = document.getElementById("result-content");
    const tldrContent = document.getElementById("tldr-content");
    if (resultContent) {
        resultContent.scrollTop = 0;
    }
    if (tldrContent) {
        tldrContent.scrollTop = 0;
    }
}

function updateResults(content) {
  // change the button status to show full content
  const expandButton = document.getElementById("expand-content");
  if (expandButton) {
    expandButton.disabled = false;
    expandButton.classList.remove("ms-Button--disabled");
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
  }

  // Remove the loading container
  const loadingContainer = document.getElementById("loading-container");
  if (loadingContainer) {
    loadingContainer.remove();
  }

  // Redo the result content
  const resultContent = document.getElementById("result-content");
  if (resultContent) {
    resultContent.innerHTML = marked.parse(content);
  }
}

// Function to reset the UI
function resetUI() {
    document.getElementById("loading").style.display = "none";
    document.getElementById("result-section").style.display = "none";
    document.getElementById("landing-screen").style.display = "block";
}

/**
 * Update reply button visibility based on settings
 */
function updateReplyButtonVisibility(show) {
  const replyButton = document.getElementById('generate-reply');
  if (replyButton) {
    replyButton.style.display = show ? 'inline-block' : 'none';
  }
}

/**
 * Expand the full content when the expand button is clicked
 */
function expandContent() {
    const expandButton = document.getElementById('expand-content');
    if (expandButton.disabled) {
        return; // Don't do anything if button is disabled
    }

    const fullContentContainer = document.getElementById('full-content-container');

    if (fullContentContainer.style.display === 'none') {
        fullContentContainer.style.display = 'block';
        expandButton.innerHTML = '<span class="ms-Button-label">Hide Full Content</span>';
        expandButton.classList.remove('ms-Button--primary');
    } else {
        fullContentContainer.style.display = 'none';
        expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
        expandButton.classList.add('ms-Button--primary');
    }
}

// Format a reply with clear subject and body sections
function formatReplyOutput(replyText) {
  // Extract subject and body
  let subject = "";
  let body = "";

  // Check for a SUBJECT: line first
  const subjectMatch = replyText.match(/^(?:SUBJECT:|Subject:)\s*(.+?)(?:\n|$)/m);
  if (subjectMatch) {
    subject = subjectMatch[1].trim();

    // Remove the subject line from the text to get the body
    body = replyText.replace(/^(?:SUBJECT:|Subject:)\s*.+?\n+/m, '').trim();
  } else {
    // If no explicit subject marker, check for first line as subject
    const lines = replyText.trim().split('\n');
    if (lines.length > 0) {
      subject = lines[0].trim();
      body = lines.slice(1).join('\n').trim();
    } else {
      // Fallback if no clear structure
      subject = "Re: Your email";
      body = replyText.trim();
    }
  }

  // Create formatted HTML
  const formattedHtml = `
    <div class="reply-container">
      <div class="reply-subject">
        <span class="reply-label">Subject:</span>${subject}
      </div>
      <div class="reply-body">${body}</div>
    </div>
  `;

  return {
    html: formattedHtml,
    subject: subject,
    body: body,
    raw: `Subject: ${subject}\n\n${body}`
  };
}

// Generate a reply based on the current content
async function generateReply() {
  const apiKey = getApiKey();

  if (!apiKey) {
    showNotification("Please add your Gemini API key in the settings", 'error');
    toggleSettingsDropdown();
    return;
  }

  // Show loading UI
  showLoading("Generating reply...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;

    // Create prompt for reply generation with structured output
    const prompt = `You are a professional email composer. Based on the following email content,
    draft a concise, professional reply. The reply MUST include:

    1. Start with "SUBJECT: " followed by an appropriate subject line, then a blank line
    2. A professional, concise email body

    Make the response clear, helpful, and to the point. Use a professional tone.

    Email content to reply to:
    Subject: ${subject}

    ${emailContent}`;

    const result = await generateContent(prompt, apiKey);
    let formattedReply = formatReplyOutput(result);

    // Display in TLDR and full content sections
    document.getElementById("tldr-content").innerHTML = formattedReply.html;
    document.getElementById("result-content").innerHTML = formattedReply.html;

    // Show the result section
    document.getElementById("result-section").style.display = "block";

    // Show the copy reply button and hide the regular copy button
    document.getElementById("copy-reply").style.display = "inline-block";
    document.getElementById("copy-result").style.display = "none";
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
  } finally {
    hideLoading();
  }
}

// Copy the reply to clipboard
function copyReply() {
  // Get the formatted reply content
  const tldrContent = document.getElementById("tldr-content").innerText;

  // Extract subject and body using regex
  const subjectMatch = tldrContent.match(/Subject:\s*(.+?)(?:\n|$)/i);
  const subject = subjectMatch ? subjectMatch[1].trim() : "";

  // Get the body (everything after the subject)
  const bodyStart = tldrContent.indexOf(subject) + subject.length;
  const body = tldrContent.substring(bodyStart).trim();

  // Format as email reply
  const replyContent = `Subject: ${subject}\n\n${body}`;

  navigator.clipboard.writeText(replyContent).then(() => {
    const copyStatus = document.getElementById("copy-reply-status");
    copyStatus.textContent = "Copied!";
    setTimeout(() => {
      copyStatus.textContent = "";
    }, 2000);
  }).catch(err => {
    console.error('Could not copy reply: ', err);
    showNotification("Failed to copy reply", "error");
  });
}

/**
 * Extract TL;DR from content, or generate a brief summary
 */
function extractTLDR(content) {
    // Check if the content already contains a TL;DR section
    const tldrRegex = /TL;DR:?\s*(.*?)(?:\n\n|$)/is;
    const tldrMatch = content.match(tldrRegex);

    if (tldrMatch && tldrMatch[1]) {
        return tldrMatch[1];
    }

    // Check for "Summary:" section
    const summaryRegex = /Summary:?\s*(.*?)(?:\n\n|$)/is;
    const summaryMatch = content.match(summaryRegex);

    if (summaryMatch && summaryMatch[1]) {
        return summaryMatch[1];
    }

    // If no TL;DR or Summary found, use the first paragraph
    const firstParagraph = content.split('\n\n')[0];
    if (firstParagraph && firstParagraph.length < 300) {
        return firstParagraph;
    } else if (firstParagraph) {
        return firstParagraph.substring(0, 250) + '...';
    }

    // Fallback
    return content.substring(0, 200) + '...';
}

/**
 * Initialize the application
 */
async function initializeApp() {
  // DOM 요소들을 먼저 찾아서 변수에 저장
  const settingsSection = document.getElementById("settings-section");
  const resultSection = document.getElementById("result-section");
  const tldrSection = document.getElementById("tldr-section");

  // 요소가 존재하는지 확인 후 스타일 적용
  if (settingsSection) {
    settingsSection.style.display = "none";
  }

  if (resultSection) {
    resultSection.style.display = "none";
  }

  if (tldrSection) {
    tldrSection.style.display = "none";
  }

  // 설정 로드
  const settings = loadSettings();
  if (settings) {
    // 테마 설정
    if (settings.theme) {
      setTheme(settings.theme);
    }

    // 폰트 크기 설정
    if (settings.fontSize) {
      setFontSize(settings.fontSize);
    }
  }

  // API 키 확인
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification("Please add your Gemini API key in settings", "warning");
  }
}

/**
 * Toggle settings panel visibility
 */
function toggleSettings() {
  const settingsSection = document.getElementById("settings-section");
  if (settingsSection.style.display === "none") {
    settingsSection.style.display = "block";
  } else {
    settingsSection.style.display = "none";
  }
}

/**
 * Save API key to settings
 */
function saveApiKey() {
  const apiKey = document.getElementById("api-key-input").value;
  localStorage.setItem("readmedarling_api_key", apiKey);
  showNotification("API key saved successfully!", "success");
  toggleSettings();
}

/**
 * Save settings to local storage
 */
function saveSettings() {
  const settings = {
    theme: document.getElementById("theme-select").value,
    fontSize: document.getElementById("font-size-select").value,
    apiKey: document.getElementById("api-key-input").value
  };
  localStorage.setItem("settings", JSON.stringify(settings));
  showNotification("Settings saved successfully!", "success");
  loadSettings();
}

/**
 * Load settings from local storage
 */
function loadSettings() {
  try {
    const savedSettings = localStorage.getItem("michael_settings");
    if (savedSettings) {
      const settings = JSON.parse(savedSettings);

      // Apply to dropdown values
      Object.keys(settings).forEach(key => {
        const element = document.getElementById(key);
        if (element && element.tagName === "SELECT") {
          element.value = settings[key];
        }
      });
    }
  } catch (error) {
    console.error("Error loading settings:", error);
  }
}

/**
 * Get a specific setting value
 */
function getSetting(key) {
  try {
    const savedSettings = localStorage.getItem("michael_settings");
    if (savedSettings) {
      const settings = JSON.parse(savedSettings);
      return settings[key];
    }
  } catch (error) {
    console.error(`Error getting setting ${key}:`, error);
  }
  return null;
}

/**
 * Set theme based on selection
 */
function setTheme(theme) {
  const root = document.documentElement;

  if (theme === "system") {
    // Use system preference
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
      root.setAttribute('data-theme', 'dark');
    } else {
      root.setAttribute('data-theme', 'light');
    }
  } else {
    // Use explicit theme
    root.setAttribute('data-theme', theme);
  }
}

/**
 * Set font size for results
 */
function setFontSize(size) {
  const root = document.documentElement;
  root.style.setProperty('--result-font-size', getFontSizeValue(size));
}

/**
 * Get font size value based on size name
 */
function getFontSizeValue(size) {
  switch (size) {
    case "small": return "0.875rem";
    case "medium": return "1rem";
    case "large": return "1.125rem";
    default: return "1rem";
  }
}

/**
 * Reset all settings to defaults
 */
function resetAllSettings() {
  // Reset model selection
  document.getElementById('dropdown-model').value = 'gemini-1.5-flash';

  // Reset language selection
  document.getElementById('dropdown-language').value = 'ko';

  // Reset theme selection
  document.getElementById('dropdown-theme').value = 'system';

  // Reset font size selection
  document.getElementById('dropdown-font-size').value = 'medium';

  // Reset TLDR mode selection
  document.getElementById('dropdown-tldr-mode').value = 'true';

  // Reset show reply selection
  document.getElementById('dropdown-show-reply').value = 'true';

  // Reset reply model selection
  document.getElementById('dropdown-reply-model').value = 'gemini-2.0-flash-lite';

  // Reset templates
  resetTemplates();

  // Clear API key
  document.getElementById('dropdown-api-key').value = '';

  // Clear saved settings from localStorage
  localStorage.removeItem('michael_settings');

  // Apply default theme
  applyCurrentTheme();

  // Apply default font size
  applyFontSize('medium');

  // Update reply button visibility
  updateReplyButtonVisibility(true);

  showNotification('All settings reset to defaults', 'success');
}
