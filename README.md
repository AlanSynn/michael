# ReadMe Darling - Outlook Add-in

A powerful Outlook add-in that helps you summarize and translate emails using Google's Gemini AI.

## Features

- 📝 Summarize emails with AI-powered analysis
- 🌐 Translate emails to multiple languages
- 🔄 Translate and summarize in one click
- ⚡ Fast and efficient processing
- 🎨 Beautiful and intuitive interface
- 🌓 Dark/Light theme support

## Installation

1. Download the manifest file from the [releases page](https://github.com/AlanSynn/michael/releases)
2. Open Outlook
3. Go to File > Options > Trust Center
4. Click "Trust Center Settings"
5. Click "Trusted Add-in Catalogs"
6. Add the following URL to the catalog:
   ```
   https://alansynn.github.io/michael
   ```
7. Check "Show in Menu"
8. Click "OK" and restart Outlook
9. Go to Home > Get Add-ins
10. Click "My Add-ins"
11. Click "Add from a File"
12. Select the downloaded manifest file
13. Click "Install"

## Usage

1. Open any email in Outlook
2. Click the "ReadMe Darling" button in the ribbon
3. Choose your desired action:
   - Summarize: Get a concise summary of the email
   - Translate: Translate the email to your preferred language
   - Translate & Summarize: Get both translation and summary

## Settings

Click the settings icon to configure:
- API Key: Your Gemini API key
- Default Language: Preferred translation language
- Theme: Light/Dark mode
- Font Size: Adjust text size
- Display Options: TL;DR mode and reply generation settings

## Development

To run the add-in locally:

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the development server:
   ```bash
   npm run dev-server
   ```
4. Follow the [Office Add-in development guide](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/) to sideload the add-in

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.