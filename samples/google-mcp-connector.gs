/*
 * Purpose: Demonstrates direct Google Native Gmail MCP connector setup.
 * Use case: Let an OpenAI Responses API model summarize Gmail data through Google's Native MCP infrastructure.
 * Required config: Store OPENAI_API_KEY in Script Properties; link Apps Script to a standard GCP project with both Gmail API and Gmail MCP API enabled.
 * Expected output: Logs a concise summary after the model uses the authorized Gmail MCP connector.
 */
function googleMcpConnectorSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const chat = GenAIApp.newChat()
    .addMessage('Summarize my latest unread Gmail message in three bullets.');

  const gmailConnector = GenAIApp.newConnector()
    .setServerUrl('https://gmailmcp.googleapis.com/mcp/v1')
    .setLabel('Google_Native_Gmail')
    .setDescription('Official Google Workspace MCP server for Gmail')
    .setAuthorization(ScriptApp.getOAuthToken())
    .setRequireApproval('never');

  chat.addMCP(gmailConnector);

  const summary = chat.run({ model: 'gpt-5.6-terra', max_tokens: 10000 });
  Logger.log(summary);
}
