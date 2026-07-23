/*
 * Purpose: Demonstrates the smallest GenAIApp chat request.
 * Use case: Use this as a hello-world smoke test after installing the library.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a short greeting or one-sentence introduction from the model.
 */
function simpleChatSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const chat = GenAIApp.newChat();
  chat.addMessage('Say hello in one friendly sentence.');

  const response = chat.run();
  Logger.log(response);
}
