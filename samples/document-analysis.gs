/*
 * Purpose: Demonstrates document analysis with addFile() using Drive IDs and Blobs.
 * Use case: Summarize PDFs or exported Google Workspace files from Apps Script.
 * Required config: Store OPENAI_API_KEY and SAMPLE_PDF_FILE_ID in Script Properties.
 * Expected output: Logs three concise bullets summarizing the supplied documents.
 */
function documentAnalysisSample() {
  const scriptProperties = PropertiesService.getScriptProperties();
  GenAIApp.setOpenAIAPIKey(scriptProperties.getProperty('OPENAI_API_KEY'));

  const driveFileId = scriptProperties.getProperty('SAMPLE_PDF_FILE_ID');
  const textBlob = Utilities.newBlob('Quarterly goals: improve support speed and reduce manual reporting.', 'text/plain', 'goals.txt');

  const chat = GenAIApp.newChat()
    .addMessage('Summarize the attached files in three bullets.')
    .addFile(driveFileId)
    .addFile(textBlob);

  const response = chat.run({ model: 'gpt-5.4' });
  Logger.log(response);
}
