/*
 * Purpose: Minimal Google Gemini File Search Store retrieval example.
 * Required config: Store a Gemini API key in Script Properties as GEMINI_API_KEY.
 */
function googleFileSearchStoreQuickstartSample() {
  GenAIApp.configureProvider('gemini', { apiKey: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') });

  const blob = Utilities.newBlob(
    'Handbook note: team demos happen every Friday at 10 AM.',
    'text/plain',
    'handbook-note.txt'
  );

  const store = GenAIApp.newVectorStore('gemini')
    .setName('Google Quickstart Store')
    .createVectorStore();

  store.uploadFile(blob, { source: 'google quickstart' });

  const answer = GenAIApp.newChat()
    .addVectorStores(store.getId())
    .addMessage('When are team demos?')
    .run({ model: 'gemini-model' });

  Logger.log(answer);
}
