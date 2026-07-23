/*
 * Purpose: Minimal OpenAI vector-store retrieval example.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 */
function openAiVectorStoreQuickstartSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const blob = Utilities.newBlob(
    'Support note: paid plans include priority support during business hours.',
    'text/plain',
    'support-note.txt'
  );

  const store = GenAIApp.newVectorStore('openai')
    .setName('OpenAI Quickstart Store')
    .createVectorStore();

  store.uploadAndAttachFile(blob, { source: 'openai quickstart' });

  const answer = GenAIApp.newChat()
    .addVectorStores(store.getId())
    .addMessage('What support is included with paid plans?')
    .run();

  Logger.log(answer);
}
