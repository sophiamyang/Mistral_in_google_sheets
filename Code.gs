// Mistral API configuration
const MISTRAL_API_KEY = 'YOUR_API_KEY_HERE';
const MISTRAL_API_URL = 'https://api.mistral.ai/v1/chat/completions';
const DEFAULT_SYSTEM_MESSAGE = "You are a helpful AI assistant.";

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
    // Prompt for system message
    const systemResult = ui.prompt(
      'System Message (Optional)',
      'Enter a system message to guide the AI, or click Cancel to use default:',
      ui.ButtonSet.OK_CANCEL
    );
    
    const systemMessage = systemResult.getSelectedButton() == ui.Button.OK ? 
      systemResult.getResponseText() : 
      DEFAULT_SYSTEM_MESSAGE;
    
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
          const response = callMistralAPI(cellValue, systemMessage);
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
function callMistralAPI(prompt, systemMessage = DEFAULT_SYSTEM_MESSAGE) {
  const apiKey = getMistralApiKey();
  const messages = [];
  
  // Add system message if provided
  if (systemMessage) {
    messages.push({
      role: "system",
      content: systemMessage
    });
  }
  
  // Add user message
  messages.push({
    role: "user",
    content: prompt
  });

  const payload = {
    model: "mistral-large-latest",
    messages: messages,
    temperature: 0.7,
    max_tokens: 500
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
 * @param {string} systemMessage Optional system message to guide the AI
 * @param {string} input The input text to process
 * @customfunction
 */
function MISTRAL(systemMessage, input) {
  if (!input) return '';
  
  try {
    const prompt = input.toString();
    return callMistralAPI(prompt, systemMessage || DEFAULT_SYSTEM_MESSAGE);
  } catch (error) {
    return `Error: ${error.message}`;
  }
} 