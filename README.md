# Mistral AI Google Sheets Integration

This Google Apps Script allows you to interact with the Mistral AI API directly from Google Sheets, enabling AI-powered text processing in your spreadsheets.

![Mistral AI Google Sheets Demo](demo.mp4)

## Features

- Custom function `=MISTRAL(cell)` for single-cell processing
- Bulk processing of selected cells
- Secure API key management

## Prerequisites

1. A Google account with access to Google Sheets
2. A Mistral AI API key ([Get one here](https://console.mistral.ai/))

## Installation

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Copy the contents of `Code.gs` into the script editor
4. Save the project (Ctrl/Cmd + S)
5. Click "Run" and authorize the script
6. Refresh your Google Sheet

## Configuration

1. In your spreadsheet, you'll see a new menu item "Mistral AI"
2. Click Mistral AI > Configure API Key
3. Enter your Mistral API key when prompted
4. Click OK to save

## Usage

### Method 1: Custom Function
Use the MISTRAL function directly in cells: 
```
=MISTRAL(, A1)    // Process content from cell A1 with default system message
=MISTRAL("You are a French language expert", A1)    // With custom system message
=MISTRAL("You are a helpful assistant", "What is the capital of France?")    // Direct prompt
```

The function accepts two parameters:
- `systemMessage`: Instructions for the AI (optional, first parameter)
- `input`: The text to process (required, second parameter)

Example system messages:
```
=MISTRAL("You are a professional translator. Translate to French.", A1)
=MISTRAL("You are a coding expert. Answer with code examples.", A1)
=MISTRAL("You are a concise writer. Keep responses under 50 words.", A1)
```

### Method 2: Bulk Processing
1. Select multiple cells containing text
2. Click Mistral AI > Process Selected Cells
3. Enter a system message when prompted (or click Cancel to use default)
4. Wait for processing to complete
5. Results will replace the original cell contents

Example use cases:
- Select a column of English text and use system message "Translate to French"
- Select product descriptions and use "Summarize this in 2 sentences"
- Select customer feedback and use "Analyze the sentiment and key points"

## Limitations

- Custom functions have a 30-second execution timeout
- Google Apps Script has daily quotas for API calls
- Mistral API has its own rate limits
- Maximum response length is set to 500 tokens

## Troubleshooting

### Common Issues:
1. "Please configure your Mistral API key first"
   - Solution: Use the Configure API Key menu option

2. Authorization Required
   - Solution: Grant necessary permissions when prompted

3. API Errors
   - Check if your API key is valid
   - Verify your Mistral API subscription status
   - Check API response in Apps Script logs

### Viewing Logs
1. Open Apps Script editor
2. Click View > Execution log
3. Run your function to see detailed logs

## Customization

You can modify the following parameters in the code:
- `max_tokens`: Maximum length of responses (default: 500)
- `temperature`: Creativity of responses (default: 0.7)
- `model`: Mistral model to use (default: "mistral-large-latest")
- `DEFAULT_SYSTEM_MESSAGE`: Default system message for all calls (default: "You are a helpful AI assistant.")

## Security Notes

- API keys are stored securely in Google's UserProperties
- Never share your API key or include it directly in the code
- Each user needs to configure their own API key