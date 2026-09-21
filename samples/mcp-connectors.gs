/*
 * Purpose: Demonstrates Google Workspace MCP connector setup for Gmail, Calendar, and Drive.
 * Use case: Let an OpenAI Responses API model inspect Workspace data through authorized connectors.
 * Required config: Store OPENAI_API_KEY in Script Properties; link Apps Script to a standard GCP project with MCP APIs enabled.
 * Expected output: Logs a concise Workspace summary after the model uses approved connectors.
 */
function mcpConnectorsSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const gmail = GenAIApp.newConnector()
    .setLabel('gmail')
    .setServerUrl('https://gmailmcp.googleapis.com/mcp/v1')
    .setAuthorization(ScriptApp.getOAuthToken())
    .setRequireApproval('never');
  const calendar = GenAIApp.newConnector()
    .setLabel('google_calendar')
    .setServerUrl('https://calendarmcp.googleapis.com/mcp/v1')
    .setAuthorization(ScriptApp.getOAuthToken())
    .setRequireApproval('never');
  const drive = GenAIApp.newConnector()
    .setLabel('google_drive')
    .setServerUrl('https://drivemcp.googleapis.com/mcp/v1')
    .setAuthorization(ScriptApp.getOAuthToken())
    .setRequireApproval('never');

  const chat = GenAIApp.newChat()
    .addMessage('Summarize my latest unread Gmail message, next calendar event, and one recently modified Drive file.')
    .addMCP(gmail)
    .addMCP(calendar)
    .addMCP(drive);

  const response = chat.run();
  Logger.log(response);
}
