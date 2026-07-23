/*
 * Purpose: Demonstrates Vertex AI authentication for Gemini without an API key.
 * Use case: Run Gemini from a Google Cloud project linked to Apps Script.
 * Required config: Store GCP_PROJECT_ID and GCP_REGION in Script Properties; enable Vertex AI and cloud-platform scopes.
 * Expected output: Logs a short Gemini response generated through Vertex AI authentication.
 */
function vertexAiSetupSample() {
  const scriptProperties = PropertiesService.getScriptProperties();
  GenAIApp.setGeminiAuth(
    scriptProperties.getProperty('GCP_PROJECT_ID'),
    scriptProperties.getProperty('GCP_REGION') || 'us-central1'
  );

  const chat = GenAIApp.newChat()
    .addMessage('Explain Vertex AI authentication for Apps Script in one sentence.');

  const response = chat.run({ model: 'gemini-model' });
  Logger.log(response);
}
