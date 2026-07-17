/*
 * Purpose: Demonstrates injecting a web page as direct context with addKnowledgeLink().
 * Use case: Answer from a known page without allowing broad web search.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a concise answer based on the Apps Script libraries guide.
 */
function knowledgeLinksSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const chat = GenAIApp.newChat()
    .addKnowledgeLink('https://developers.google.com/apps-script/guides/libraries')
    .addMessage('Based only on the provided knowledge link, what is one reason to use an Apps Script library?');

  const response = chat.run({ model: 'gpt-5.6-terra' });
  Logger.log(response);
}
