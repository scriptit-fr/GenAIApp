/*
 * Purpose: Demonstrates Google Workspace MCP connector setup for Gmail, Calendar, and Drive.
 * Use case: Let an OpenAI Responses API model inspect Workspace data through authorized connectors.
 * Required config: Store OPENAI_API_KEY in Script Properties; link Apps Script to a standard GCP project with MCP APIs enabled.
 * Expected output: Logs a concise Workspace summary after the model uses approved connectors.
 */
function mcpConnectorsSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const gmail = GenAIApp.newConnector().setConnectorId('gmail').setAuthorization(ScriptApp.getOAuthToken()).setRequireApproval('never');
  const calendar = GenAIApp.newConnector().setConnectorId('calendar').setAuthorization(ScriptApp.getOAuthToken()).setRequireApproval('never');
  const drive = GenAIApp.newConnector().setConnectorId('drive').setAuthorization(ScriptApp.getOAuthToken()).setRequireApproval('never');

  const chat = GenAIApp.newChat()
    .addMessage('Summarize my latest unread Gmail message, next calendar event, and one recently modified Drive file.')
    .addMCP(gmail)
    .addMCP(calendar)
    .addMCP(drive);

  const response = chat.run({ model: 'gpt-5.4', max_tokens: 20000 });
  Logger.log(response);
}
