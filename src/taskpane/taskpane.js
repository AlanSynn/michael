/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

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

// Default templates
const DEFAULT_TEMPLATES = {
  summarize: `You are an expert researcher. Your task is to carefully review the provided research document and perform the following tasks:

1. Summarize the Research Background:
Provide a concise yet comprehensive summary of the research background, highlighting the context and motivation behind the study.

2. Extract the Problem Statement:
Identify and articulate the central problem or challenge addressed by the research in clear and precise terms.

3. Identify Strengths:
List between three and five key strengths of the study. Focus on aspects such as methodology, innovation, robustness of results, or any other notable positive attributes.

4. Identify Weaknesses:
Enumerate between four and five significant weaknesses or limitations present in the research. Consider issues like methodological gaps, limited scope, or any areas lacking clarity.

5. Propose Research Topics:
Based on the weaknesses identified, suggest three potential research topics that could address these limitations or explore related areas further.

Ensure your response is thorough and balanced, with academical depth and varied sentence structures that reflect both detailed insight and succinct clarity.

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
1. First provide a concise summary section (제목: 요약)
2. Then provide the full translation (제목: 전체 번역)

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
    document.getElementById("summarize-button").addEventListener("click", summarizeEmail);
    document.getElementById("settings-button").addEventListener("click", toggleSettings);
    document.getElementById("close-settings").addEventListener("click", toggleSettings);
    document.getElementById("save-api-key").addEventListener("click", saveApiKey);
    document.getElementById("api-key-input").addEventListener("keypress", (event) => {
      if (event.key === "Enter") {
        saveApiKey();
      }
    });

    // Settings selection change listeners
    document.querySelectorAll(".settings-dropdown-container select").forEach(select => {
      select.addEventListener("change", saveSettings);
    });

    // Expand button listener
    document.getElementById("expand-button").addEventListener("click", () => {
      const fullContentContainer = document.getElementById("full-content-container");
      const expandButton = document.getElementById("expand-button");

      if (fullContentContainer.style.display === "none") {
        fullContentContainer.style.display = "block";
        expandButton.querySelector(".ms-Button-label").textContent = "Collapse";
      } else {
        fullContentContainer.style.display = "none";
        expandButton.querySelector(".ms-Button-label").textContent = "Expand";
      }
    });

    // Copy buttons listeners
    document.getElementById("copy-result-button").addEventListener("click", copyResult);
    document.getElementById("copy-reply-button").addEventListener("click", copyReply);

    // Generate reply button listener
    document.getElementById("generate-reply-button").addEventListener("click", generateReply);

    // Initialize dropdown settings events
    document.getElementById("dropdown-save-settings").onclick = saveDropdownSettings;
    document.getElementById("dropdown-close-settings").onclick = toggleSettingsDropdown;
    document.getElementById("dropdown-reset-templates").onclick = resetTemplates;

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

    // Add function to window object so it can be called from onclick
    window.toggleSettingsDropdown = toggleSettingsDropdown;

    // TLDR expand button listener
    document.getElementById("expand-content").addEventListener("click", expandContent);
  }
});

/**
 * Toggle the visibility of the settings dropdown
 */
function toggleSettingsDropdown() {
  const dropdown = document.getElementById("settings-dropdown");
  if (dropdown.style.display === "none") {
    dropdown.style.display = "block";
    // Load latest settings when opening
    loadDropdownSettings();
  } else {
    dropdown.style.display = "none";
  }
}

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
  // Create notification element if it doesn't exist
  let notification = document.getElementById("notification");
  if (!notification) {
    notification = document.createElement("div");
    notification.id = "notification";
    document.body.appendChild(notification);
  }

  // Set type class
  notification.className = `notification ${type}`;
  notification.textContent = message;

  // Show notification
  notification.style.display = "block";

  // Hide after delay
  setTimeout(() => {
    notification.style.display = "none";
  }, 3000);
}

/**
 * Save settings from the dropdown menu
 */
function saveDropdownSettings() {
  // Get existing settings
  let settings = {};
  try {
    const savedSettings = localStorage.getItem('readmedarling_settings');
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
  localStorage.setItem('readmedarling_settings', JSON.stringify(settings));

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
    const savedSettings = localStorage.getItem('readmedarling_settings');

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
async function generateContent(prompt, apiKey, modelOverride = null) {
    // Get model from settings or use default
    let model = "gemini-1.5-flash"; // Default model

    if (modelOverride) {
        model = modelOverride;
    } else {
        try {
            const savedSettings = localStorage.getItem('readmedarling_settings');
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
                    maxOutputTokens: 8192,
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
    }
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
  return localStorage.getItem("gemini_api_key");
}

// Get language from settings
function getLanguage() {
  try {
    const savedSettings = localStorage.getItem('readmedarling_settings');
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
      const savedSettings = localStorage.getItem('readmedarling_settings');
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

    const summary = await generateContent(prompt, apiKey);

    // Display result with markdown rendering
    showResults(summary);
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
  }
}

// Translate email
async function translateEmail() {
  const apiKey = getApiKey();

  if (!apiKey) {
    showNotification("Please add your Gemini API key in the settings", 'error');
    toggleSettingsDropdown(); // Open settings dropdown to prompt for API key
    return;
  }

  // Show loading UI
  showLoading("Translating to Korean...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;

    // Get template from storage or use default
    let template = DEFAULT_TEMPLATES.translate;
    try {
      const savedSettings = localStorage.getItem('readmedarling_settings');
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

    const translation = await generateContent(prompt, apiKey);

    // Display result with markdown rendering
    showResults(translation);
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
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
    const savedSettings = localStorage.getItem('readmedarling_settings');
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

    // Get template from storage or use default
    let template = DEFAULT_TEMPLATES.translateSummarize;
    try {
      const savedSettings = localStorage.getItem('readmedarling_settings');
      if (savedSettings) {
        const settings = JSON.parse(savedSettings);
        if (settings.templates && settings.templates.translateSummarize) {
          template = settings.templates.translateSummarize;
        }
      }
    } catch (error) {
      console.error("Error getting template:", error);
    }

    // Replace placeholders in template
    const prompt = template
      .replace('{subject}', subject)
      .replace('{content}', emailContent);

    const result = await generateContent(prompt, apiKey);

    // Display result with markdown rendering
    showResults(result);
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
  }
}

/**
 * Show loading indicator with message
 */
function showLoading(message = "Loading...") {
  // Create or get loading container
  let loadingContainer = document.getElementById("loading-container");
  if (!loadingContainer) {
    loadingContainer = document.createElement("div");
    loadingContainer.id = "loading-container";
    loadingContainer.className = "loading-container";

    const spinner = document.createElement("div");
    spinner.className = "loading-spinner";

    const messageElem = document.createElement("div");
    messageElem.id = "loading-message";
    messageElem.className = "loading-message";

    loadingContainer.appendChild(spinner);
    loadingContainer.appendChild(messageElem);
    document.body.appendChild(loadingContainer);
  }

  // Set message
  document.getElementById("loading-message").textContent = message;

  // Show loading
  loadingContainer.style.display = "flex";
}

/**
 * Hide loading indicator
 */
function hideLoading() {
  const loadingContainer = document.getElementById("loading-container");
  if (loadingContainer) {
    loadingContainer.style.display = "none";
  }
}

// Function to show the results
function showResults(content) {
    document.getElementById("loading").style.display = "none";
    document.getElementById("landing-screen").style.display = "none";
    document.getElementById("result-section").style.display = "block";

    // Check for TLDR mode
    let tldrMode = true; // Default to TLDR mode on
    try {
      const savedSettings = localStorage.getItem('readmedarling_settings');
      if (savedSettings) {
        const settings = JSON.parse(savedSettings);
        if (settings.tldrMode) {
          tldrMode = settings.tldrMode === 'true';
        }
      }
    } catch (error) {
      console.error('Error getting TLDR mode setting:', error);
    }

    // Generate a TL;DR if none exists in the content
    let tldrContent = extractTLDR(content);

    // Use the existing markdownToHtml function
    document.getElementById("result-content").innerHTML = markdownToHtml(content);
    document.getElementById("tldr-content").innerHTML = markdownToHtml(tldrContent);

    // Show/hide based on TLDR mode
    if (tldrMode) {
      document.getElementById("full-content-container").style.display = "none";
      document.getElementById("expand-content").innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    } else {
      document.getElementById("full-content-container").style.display = "block";
      document.getElementById("expand-content").innerHTML = '<span class="ms-Button-label">Hide Full Content</span>';
    }

    // Apply font size from settings
    try {
      const savedSettings = localStorage.getItem('readmedarling_settings');
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

    // Scroll to top of result content
    document.getElementById("result-content").scrollTop = 0;
    document.getElementById("tldr-content").scrollTop = 0;
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
  const replyButton = document.getElementById('generate-reply-button');
  if (replyButton) {
    replyButton.style.display = show ? 'inline-block' : 'none';
  }
}

/**
 * Expand the full content when the expand button is clicked
 */
function expandContent() {
  const fullContentContainer = document.getElementById('full-content-container');
  const expandButton = document.getElementById('expand-content');

  if (fullContentContainer.style.display === 'none') {
    fullContentContainer.style.display = 'block';
    expandButton.innerHTML = '<span class="ms-Button-label">Hide Full Content</span>';
  } else {
    fullContentContainer.style.display = 'none';
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
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
    toggleSettingsDropdown(); // Open settings dropdown to prompt for API key
    return;
  }

  // Show loading UI
  showLoading("Generating reply...");

  try {
    // Get settings
    let replyModel = "gemini-2.0-flash-light"; // Default model

    try {
      const savedSettings = localStorage.getItem('readmedarling_settings');
      if (savedSettings) {
        const settings = JSON.parse(savedSettings);
        if (settings.replyModel) {
          replyModel = settings.replyModel;
        }
      }
    } catch (error) {
      console.error("Error getting reply model setting:", error);
    }

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

    const result = await generateContent(prompt, apiKey, replyModel);

    // Format the result to display subject and body separately
    let formattedReply = formatReplyOutput(result);

    // Display in TLDR and full content sections
    document.getElementById("tldr-content").innerHTML = formattedReply.html;
    document.getElementById("result-content").innerHTML = formattedReply.html;

    // Show the result section
    document.getElementById("result-section").style.display = "block";

    // Show the copy reply button and hide the regular copy button
    document.getElementById("copy-reply-button").style.display = "inline-block";
    document.getElementById("copy-result-button").style.display = "none";

    // Hide loading
    hideLoading();
  } catch (error) {
    hideLoading();
    showNotification(`Error: ${error.message}`, 'error');
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
function initializeApp() {
  // Load settings from local storage
  loadSettings();

  // Check if API key is set
  const apiKey = getApiKey();
  if (!apiKey) {
    // Show settings panel if no API key
    document.getElementById("settings-section").style.display = "block";
    showNotification("Please set your Gemini API key in the settings", "info");
  }

  // Set theme based on saved settings
  const savedTheme = getSetting("theme") || "system";
  setTheme(savedTheme);

  // Set font size based on saved settings
  const savedFontSize = getSetting("resultFontSize") || "medium";
  setFontSize(savedFontSize);
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
 * Save API key to local storage
 */
function saveApiKey() {
  const apiKeyInput = document.getElementById("api-key-input");
  const apiKey = apiKeyInput.value.trim();

  if (apiKey) {
    localStorage.setItem("gemini_api_key", apiKey);
    showNotification("API key saved successfully", "success");
    document.getElementById("settings-section").style.display = "none";
  } else {
    showNotification("Please enter a valid API key", "error");
  }
}

/**
 * Save settings to local storage
 */
function saveSettings() {
  const settings = {};

  // Get all dropdown settings
  document.querySelectorAll(".settings-dropdown-container select").forEach(select => {
    settings[select.id] = select.value;
  });

  // Save to local storage
  localStorage.setItem("readmedarling_settings", JSON.stringify(settings));

  // Apply settings
  if (settings.theme) {
    setTheme(settings.theme);
  }

  if (settings.resultFontSize) {
    setFontSize(settings.resultFontSize);
  }

  showNotification("Settings saved", "success");
}

/**
 * Load settings from local storage
 */
function loadSettings() {
  try {
    const savedSettings = localStorage.getItem("readmedarling_settings");
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
    const savedSettings = localStorage.getItem("readmedarling_settings");
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
