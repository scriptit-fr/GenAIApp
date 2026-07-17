/*
 * Purpose: Demonstrates common GenAIApp configuration and guardrail options.
 * Use case: Control budget, monitor token use, reduce logs, and enable long-context compaction.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a compact project-summary response while enforcing the configured limits.
 */
function configurationOptionsSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const chat = GenAIApp.newChat()
    .setMaximumAPICalls(3)
    .warnIfResponseTokenUsageAbove(500)
    .disableLogs(true)
    .enableCompaction(true)
    .setCompactionThreshold(10000)
    .addMessage('Summarize three practical ways to keep AI usage predictable in Apps Script.');

  const response = chat.run({ model: 'gpt-5.6-terra', max_tokens: 800 });
  Logger.log(response);
}
