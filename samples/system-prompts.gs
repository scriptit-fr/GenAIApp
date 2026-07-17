/*
 * Purpose: Shows how to set assistant personality and context with a system message.
 * Use case: Keep responses in a specific tone, role, or format for a user-facing app.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a concise response written in the configured librarian style.
 */
function systemPromptsSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const chat = GenAIApp.newChat()
    .addMessage('You are a patient librarian. Answer in two calm bullet points.', true)
    .addMessage('How should I choose my next book?');

  const response = chat.run({ model: 'gpt-5.6-terra' });
  Logger.log(response);
}
