# Email Validation using AI

This code helps you compose emails with AI-powered validation and refinement capabilities using Google's Gemini API.

## Features

- **GUI**: Clean, responsive interface built with PyQt5
- **Rich Text Editing**: Format your emails with various styling options
- **AI-Powered Validation**: Validate your emails for common issues using Google's Gemini AI
- **Content Refinement**: Get AI suggestions to improve your email content
- **File Attachments**: Easily attach and manage files
- **Gmail Integration**: Send emails directly through Gmail's SMTP server

## Requirements

- Python 3.6+
- PyQt5
- Requests
- Google Gemini API key

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/email-composer-ai.git
   cd email-composer-ai
   ```

2. Install required packages:
   ```
   pip install PyQt5 requests
   ```

3. Add GEMINI API key and your email address and app password.
   
## Getting a Google Gemini API Key

1. Visit the [Google AI Studio](https://aistudio.google.com/apikey)
2. Sign in with your Google account
3. Click on "Get API key" or "Create API key"
4. Copy the generated API key
5. Add it to the code at `self.api_key`

## Getting a Gmail App Password

For security reasons, Gmail requires an app password for SMTP access:

1. Go to your [Google Account](https://myaccount.google.com/)
2. Select "Security"
3. Under "Signing in to Google," select "2-Step Verification" (enable if not already)
4. At the bottom of the page, select "App passwords"
5. Copy the 16-character password
8. Add email address and password at `self.email` and `self.password` 

## Usage

1. Run the application
   
2. Compose your email:
   - Enter recipient email address
   - Add a subject
   - Compose your message using the rich text editor
   - Add attachments if needed

3. Validate your email:
   - Click "Validate with Gemini" to check for issues
   - Review any validation feedback

4. Refine your email:
   - Click "Refine Email" to get AI suggestions
   - Review and insert the refined content if desired

6. Send your email:
   - Click "Send Email" when ready

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
