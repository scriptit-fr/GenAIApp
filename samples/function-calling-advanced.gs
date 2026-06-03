/*
 * Purpose: Shows advanced function-calling controls in one extraction/routing flow.
 * Use case: Extract arguments without execution, or terminate early after a tool result.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a JSON object containing ticket fields because onlyReturnArguments(true) ends before execution.
 */
function functionCallingAdvancedSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const extractTicket = GenAIApp.newFunction()
    .setName('extractSupportTicket')
    .setDescription('Extracts support-ticket fields from a user message.')
    .addParameter('email', 'string', 'Customer email address')
    .addParameter('category', 'string', 'Short issue category')
    .addParameter('priority', 'string', 'low, normal, or urgent')
    .onlyReturnArguments(true);

  const lookupCustomer = GenAIApp.newFunction()
    .setName('lookupCustomerPlan')
    .setDescription('Looks up a customer plan by email.')
    .addParameter('email', 'string', 'Customer email address')
    .endWithResult(true);

  const chat = GenAIApp.newChat()
    .addMessage('Extract this ticket: urgent billing problem for ana@example.com.')
    .addFunction(extractTicket)
    .addFunction(lookupCustomer);

  const response = chat.run({ model: 'gpt-5.4', function_call: 'extractSupportTicket' });
  Logger.log(response);
}

function extractSupportTicket(email, category, priority) {
  return { email: email, category: category, priority: priority };
}

function lookupCustomerPlan(email) {
  return { email: email, plan: 'Business', status: 'active' };
}
