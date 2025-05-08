<div align="center">
  <img src="assets/meet-michael-black.png" alt="Meet Michael Logo" width="300"/>
</div>

Meet Michael, your AI sidekick for Outlook! Michael enhances your email productivity by leveraging the power of Google's Gemini AI. It allows you to quickly summarize, translate, and understand your emails directly within Outlook.

## Features

*   **Summarize:** Get concise summaries of long emails.
*   **Translate:** Translate email content into various languages (default: Korean).
*   **Translate & Summarize:** Get both a translation and a summary in one go.
*   **Calendar Event Creation:** Automatically extract event details from email content and create calendar entries.
*   **Customizable Settings:**
    *   Configure your Google Gemini API Key.
    *   Select the Gemini model to use (e.g., gemini-2.0-flash-lite).
    *   Set default translation language.
    *   Choose UI theme (Light, Dark, System Default).
    *   Adjust font size for readability.
    *   Enable/Disable TL;DR mode for quick previews.
    *   Configure Autorun action on email open.
    *   Export/Import custom prompt templates.
*   🌓 Dark/Light theme support, adapting to your Outlook theme or manual selection.

## Setup for Development

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/AlanSynn/michael.git # Or your repository URL
    cd michael # Or the directory name you cloned into
    ```
2.  **Install dependencies:**
    ```bash
    npm install
    ```
3.  **Start the development server:**
    ```bash
    npm start
    ```
    This will start a local web server (usually `https://localhost:3000`) and watch for file changes.

4.  **Sideload the Add-in in Outlook:**
    *   Follow the instructions for sideloading Outlook add-ins based on your platform (Windows, Mac, Web). You will typically need the `manifest.xml` file located in the project root.
    *   Microsoft Docs: [Sideload Outlook add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)

## Configuration

1.  **Obtain a Gemini API Key:** Get an API key from [Google AI Studio](https://aistudio.google.com/app/apikey).
2.  **Open the Add-in in Outlook:** Select an email and click the "Summarize & Translate" button in the Outlook ribbon (or the button name you configured in `manifest.xml`).
3.  **Open Settings:** Click the settings cog icon within the add-in taskpane.
4.  **Enter API Key:** Paste your Gemini API key into the designated field and save.
5.  **Configure other settings** like model, language, theme, etc., as needed.

## Usage

1.  Select an email in Outlook.
2.  Click the "Summarize & Translate" button in the ribbon to open the My Sidekick, Michael taskpane.
3.  Use the buttons within the taskpane (Summarize, Translate, etc.) to interact with the email content using Gemini AI.
4.  If the email content is detected as a potential calendar event, the "Create Calendar Event" button will be enabled.
5.  Adjust settings via the settings cog icon.

## Technology Stack

*   Office Add-ins Platform
*   Google Gemini API
*   JavaScript (ES6+)
*   HTML5
*   CSS3
*   Webpack
*   Node.js / npm

## License

MIT License