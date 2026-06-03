/*
 * Purpose: Demonstrates multi-turn OpenAI conversations with previous response IDs.
 * Use case: Continue a Responses API conversation without resending the whole transcript.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs the first response ID and a second answer that remembers the chosen color.
 */
function conversationContinuationSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const firstChat = GenAIApp.newChat()
    .addMessage('Remember this preference: my dashboard accent color is teal.');
  Logger.log(firstChat.run({ model: 'gpt-5.4' }));

  const previousResponseId = firstChat.retrieveLastResponseId();
  Logger.log('Previous response ID: ' + previousResponseId);

  const secondChat = GenAIApp.newChat()
    .setPreviousResponseId(previousResponseId)
    .addMessage('What accent color did I choose?');

  const response = secondChat.run({ model: 'gpt-5.4' });
  Logger.log(response);
}
