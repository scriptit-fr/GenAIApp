/*
 * Purpose: Demonstrates reusing the same chat setup with different models.
 * Use case: Compare OpenAI and Gemini responses without changing prompts.
 * Required config: Store OPENAI_API_KEY and GEMINI_API_KEY in Script Properties.
 * Expected output: Logs one response per model for the same haiku-generation prompt.
 */
function multiModelUsageSample() {
  const scriptProperties = PropertiesService.getScriptProperties();
  GenAIApp.setOpenAIAPIKey(scriptProperties.getProperty('OPENAI_API_KEY'));
  GenAIApp.setGeminiAPIKey(scriptProperties.getProperty('GEMINI_API_KEY'));

  const models = ['openAIModel', 'GeminiModel'];
  models.forEach(function (model) {
    const chat = GenAIApp.newChat()
      .addMessage('Write a haiku about Apps Script automation.');

    const response = chat.run({ model: model });
    Logger.log(model + ': ' + response);
  });
}
