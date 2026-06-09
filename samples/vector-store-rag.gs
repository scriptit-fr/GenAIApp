/*
 * Purpose: Demonstrates a full OpenAI vector-store RAG workflow.
 * Use case: Create a store, upload source files with attributes, attach it to a chat, and query it.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs the answer from file search, then logs raw chunks from onlyReturnChunks(true).
 */
function vectorStoreRagSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const policyBlob = Utilities.newBlob(
    'Refund policy: refunds are available within 30 days with a receipt.',
    'text/plain',
    'refund-policy.txt'
  );

  const vectorStore = GenAIApp.newVectorStore()
    .setName('Sample Support Knowledge')
    .setDescription('Tiny sample knowledge base for support answers')
    .setChunkingStrategy(800, 200)
    .createVectorStore();

  vectorStore.uploadAndAttachFile(policyBlob, { topic: 'refunds', source: 'sample' });

  const answer = GenAIApp.newChat()
    .addVectorStores(vectorStore.getId())
    .addMessage('What is the refund window?')
    .run({ model: 'gpt-5.4' });
  Logger.log(answer);

  const chunks = GenAIApp.newChat()
    .addVectorStores(vectorStore.getId())
    .onlyReturnChunks(true)
    .setMaxChunks(3)
    .addMessage('refund window')
    .run({ model: 'gpt-5.4' });
  Logger.log(chunks);
}
