// Mistral API configuration
const MISTRAL_API_KEY = 'YOUR_API_KEY_HERE';
const MISTRAL_API_URL = 'https://api.mistral.ai/v1/chat/completions';

/**
 * Creates a menu item for the spreadsheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mistral AI')
    .addItem('Process Selected Cells', 'processSelectedCells')
    .addItem('Configure API Key', 'showApiKeyDialog')
    .addToUi();
}

/**
 * Shows a dialog to configure the API key
 */
function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Configure Mistral API Key',
    'Please enter your Mistral API key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const apiKey = result.getResponseText();
    PropertiesService.getUserProperties().setProperty('MISTRAL_API_KEY', apiKey);
    ui.alert('API key saved successfully!');
  }
}

/**
 * Gets the stored API key or prompts user to enter one
 */
function getMistralApiKey() {
  const apiKey = PropertiesService.getUserProperties().getProperty('MISTRAL_API_KEY');
  if (!apiKey) {
    throw new Error('Please configure your Mistral API key first');
  }
  return apiKey;
}

/**
 * Processes the selected cells using Mistral AI
 */
function processSelectedCells() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    const values = range.getValues();
    
    // Show processing indicator
    const statusRange = sheet.getRange(1, range.getLastColumn() + 2);
    statusRange.setValue('Processing...');
    
    // Process each cell in the selection
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const cellValue = values[i][j];
        if (cellValue) {
          const response = callMistralAPI(cellValue);
          range.getCell(i + 1, j + 1).setValue(response);
          // Add a small delay to avoid rate limits
          Utilities.sleep(100);
        }
      }
    }
    
    statusRange.setValue('Done!');
    Utilities.sleep(2000);
    statusRange.clearContent();
    
  } catch (error) {
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Calls the Mistral API with the given prompt
 */
function callMistralAPI(prompt) {
  const apiKey = getMistralApiKey();
  const payload = {
    model: "mistral-large-latest",  // Using the latest large model
    messages: [
      {
        role: "user",
        content: prompt
      }
    ],
    temperature: 0.7,
    max_tokens: 500  // Add token limit for safety
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(MISTRAL_API_URL, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`API returned status ${responseCode}: ${response.getContentText()}`);
    }
    
    const jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse.choices[0].message.content;
  } catch (error) {
    Logger.log('Error calling Mistral API:', error);
    throw new Error(`Failed to process request: ${error.message}`);
  }
}

/**
 * Custom function to call Mistral AI from a cell
 * @param {string} input The input text to process
 * @customfunction
 */
function MISTRAL(input) {
  if (!input) return '';
  
  try {
    // Convert input to string if it's not already
    const prompt = input.toString();
    return callMistralAPI(prompt);
  } catch (error) {
    return `Error: ${error.message}`;
  }
} 