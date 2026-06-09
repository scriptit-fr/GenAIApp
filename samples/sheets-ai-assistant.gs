/*
 * Purpose: Demonstrates a complete Google Sheets AI-assistant integration.
 * Use case: Read active-sheet data, ask AI for a summary, and write the result back to the sheet.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY; run from a bound Google Sheet.
 * Expected output: Writes an AI summary into cell A1 of a new sheet named AI Summary.
 */
function sheetsAiAssistantSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getActiveSheet();
  const data = sourceSheet.getDataRange().getDisplayValues();
  const previewRows = data.slice(0, 20).map(function (row) {
    return row.join(' | ');
  }).join('\n');

  const prompt = 'Analyze this sheet data and return three bullets with key observations:\n' + previewRows;
  const response = GenAIApp.newChat()
    .addMessage(prompt)
    .run({ model: 'gpt-5.4', max_tokens: 800 });

  const outputSheet = spreadsheet.getSheetByName('AI Summary') || spreadsheet.insertSheet('AI Summary');
  outputSheet.clear();
  outputSheet.getRange('A1').setValue(response);
  outputSheet.getRange('A1').setWrap(true);
}
