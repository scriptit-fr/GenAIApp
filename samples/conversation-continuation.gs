/*
 * Purpose: Demonstrates multi-turn OpenAI Responses API and Gemini Interactions API conversations.
 * Use case: Continue a conversation without resending the whole transcript.
 * Required config: Store OPENAI_API_KEY and GEMINI_API_KEY in Script Properties.
 * Expected output: Logs continuation IDs and follow-up answers that remember the chosen colors.
 */
function conversationContinuationSample() {
  const scriptProperties = PropertiesService.getScriptProperties();
  GenAIApp.setOpenAIAPIKey(scriptProperties.getProperty('OPENAI_API_KEY'));
  GenAIApp.setGeminiAPIKey(scriptProperties.getProperty('GEMINI_API_KEY'));

  const firstOpenAiChat = GenAIApp.newChat()
    .addMessage('Remember this preference: my dashboard accent color is teal.');
  Logger.log(firstOpenAiChat.run());

  const previousResponseId = firstOpenAiChat.retrieveLastResponseId();
  Logger.log('Previous OpenAI response ID: ' + previousResponseId);

  const secondOpenAiChat = GenAIApp.newChat()
    .setPreviousResponseId(previousResponseId)
    .addMessage('What accent color did I choose?');
  Logger.log(secondOpenAiChat.run());

  const firstGeminiChat = GenAIApp.newChat()
    .addMessage('Remember this preference: my report accent color is amber.');
  Logger.log(firstGeminiChat.run({ model: 'gemini-model' }));

  const previousInteractionId = firstGeminiChat.retrieveLastInteractionId();
  Logger.log('Previous Gemini interaction ID: ' + previousInteractionId);

  const secondGeminiChat = GenAIApp.newChat()
    .setPreviousInteractionId(previousInteractionId)
    .addMessage('What report accent color did I choose?');
  Logger.log(secondGeminiChat.run({ model: 'gemini-model' }));
}
