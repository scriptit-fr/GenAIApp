/*
 * Purpose: Demonstrates real-time web browsing with an optional domain restriction.
 * Use case: Ask for current information while limiting browsing to a trusted site.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a brief answer grounded in content found under developers.google.com.
 */
function webBrowsingSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const chat = GenAIApp.newChat()
    .enableBrowsing(true, 'https://developers.google.com')
    .addMessage('Find one current Apps Script documentation page about triggers and summarize it in two sentences.');

  const response = chat.run();
  Logger.log(response);
}
