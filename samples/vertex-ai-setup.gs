/*
 * Purpose: Demonstrates Vertex AI authentication for Gemini without an API key.
 * Use case: Run Gemini from a Google Cloud project linked to Apps Script.
 * Required config: Store GCP_PROJECT_ID in Script Properties; enable Vertex AI and cloud-platform scopes. Leave GCP_REGION unset to use the global endpoint.
 * Expected output: Logs a short Gemini response generated through Vertex AI authentication.
 */
function vertexAiSetupSample() {
  const scriptProperties = PropertiesService.getScriptProperties();
  GenAIApp.configureProvider('gemini', {
    projectId: scriptProperties.getProperty('GCP_PROJECT_ID'),
    region: scriptProperties.getProperty('GCP_REGION') || ''
  });

  const chat = GenAIApp.newChat()
    .addMessage('Explain Vertex AI authentication for Apps Script in one sentence.');

  const response = chat.run({ model: 'gemini-4-hydrogen' });
  Logger.log(response);
}
