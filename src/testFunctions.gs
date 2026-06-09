const GPT_MODEL = "gpt-5.4";
const REASONING_MODEL = "o4-mini";
const GEMINI_MODEL = "gemini-3.5-flash";
const TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID = "";
const TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID = "";

const AUTH_TEST_CONFIG_KEYS = {
  ENABLE_API_KEY_AUTH_TESTS: "ENABLE_API_KEY_AUTH_TESTS",
  ENABLE_VERTEX_AI_AUTH_TESTS: "ENABLE_VERTEX_AI_AUTH_TESTS",
  GEMINI_API_KEY: "GEMINI_API_KEY",
  VERTEX_AI_GCP_PROJECT_ID: "VERTEX_AI_GCP_PROJECT_ID",
  VERTEX_AI_GCP_REGION: "VERTEX_AI_GCP_REGION",
  OPEN_AI_API_KEY: "OPEN_AI_API_KEY"
};

/**
 * Reads test configuration from environment variables, Apps Script properties,
 * or legacy global constants. This helper never logs credential values.
 */
function getAuthTestConfigValue(name, defaultValue) {
  if (typeof process !== "undefined" && process.env && process.env[name] !== undefined) {
    return process.env[name];
  }

  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const propertyValue = scriptProperties.getProperty(name);
    if (propertyValue !== null && propertyValue !== undefined) {
      return propertyValue;
    }
  }
  catch (err) {
    // PropertiesService is unavailable outside Apps Script. Fall through to
    // constants/defaults so local syntax checks can still load this file.
  }

  if (name === AUTH_TEST_CONFIG_KEYS.ENABLE_API_KEY_AUTH_TESTS && typeof ENABLE_API_KEY_AUTH_TESTS !== "undefined") {
    return ENABLE_API_KEY_AUTH_TESTS;
  }
  if (name === AUTH_TEST_CONFIG_KEYS.ENABLE_VERTEX_AI_AUTH_TESTS && typeof ENABLE_VERTEX_AI_AUTH_TESTS !== "undefined") {
    return ENABLE_VERTEX_AI_AUTH_TESTS;
  }
  if (name === AUTH_TEST_CONFIG_KEYS.GEMINI_API_KEY && typeof GEMINI_API_KEY !== "undefined") {
    return GEMINI_API_KEY;
  }
  if (name === AUTH_TEST_CONFIG_KEYS.VERTEX_AI_GCP_PROJECT_ID && typeof VERTEX_AI_GCP_PROJECT_ID !== "undefined") {
    return VERTEX_AI_GCP_PROJECT_ID;
  }
  if (name === AUTH_TEST_CONFIG_KEYS.VERTEX_AI_GCP_REGION && typeof VERTEX_AI_GCP_REGION !== "undefined") {
    return VERTEX_AI_GCP_REGION;
  }
  if (name === AUTH_TEST_CONFIG_KEYS.OPEN_AI_API_KEY && typeof OPEN_AI_API_KEY !== "undefined") {
    return OPEN_AI_API_KEY;
  }

  return defaultValue;
}

function getAuthTestBoolean(name, defaultValue) {
  const rawValue = getAuthTestConfigValue(name, defaultValue ? "true" : "false");
  if (typeof rawValue === "boolean") {
    return rawValue;
  }
  return String(rawValue).toLowerCase() === "true";
}

function requireAuthTestCredential(name, modeName) {
  const value = getAuthTestConfigValue(name, "").trim();
  if (!value) {
    throw new Error(`[GenAIApp tests] ${modeName} tests are enabled, but required credential/configuration ${name} is missing.`);
  }
  return value;
}

function skipAuthTest(modeName, reason) {
  console.log(`[GenAIApp tests] Skipping ${modeName} authentication tests: ${reason}`);
}

/**
 * Stores auth-test switches in script properties before running the Apps Script
 * test entrypoint. Do not pass raw secrets through this function.
 */
function configureAuthTestSwitches(enableApiKeyAuthTests, enableVertexAiAuthTests) {
  PropertiesService.getScriptProperties().setProperties({
    ENABLE_API_KEY_AUTH_TESTS: String(enableApiKeyAuthTests),
    ENABLE_VERTEX_AI_AUTH_TESTS: String(enableVertexAiAuthTests)
  });
}

/**
 * Entrypoint for authentication coverage. Set ENABLE_API_KEY_AUTH_TESTS and
 * ENABLE_VERTEX_AI_AUTH_TESTS independently to run API-key auth, Vertex AI
 * auth, both, or neither. Disabled modes log a clear skip. Enabled modes fail
 * before calling the API when required credentials/configuration are absent.
 */
function testConfiguredAuthenticationModes() {
  const runApiKeyAuthTests = getAuthTestBoolean(AUTH_TEST_CONFIG_KEYS.ENABLE_API_KEY_AUTH_TESTS, true);
  const runVertexAiAuthTests = getAuthTestBoolean(AUTH_TEST_CONFIG_KEYS.ENABLE_VERTEX_AI_AUTH_TESTS, false);

  if (!runApiKeyAuthTests && !runVertexAiAuthTests) {
    console.log("[GenAIApp tests] All Gemini authentication-mode tests are disabled.");
    return;
  }

  if (runApiKeyAuthTests) {
    testApiKeyAuthentication();
  }
  else {
    skipAuthTest("API key", "ENABLE_API_KEY_AUTH_TESTS is not true");
  }

  if (runVertexAiAuthTests) {
    testVertexAiAuthentication();
  }
  else {
    skipAuthTest("Vertex AI", "ENABLE_VERTEX_AI_AUTH_TESTS is not true");
  }
}

function testApiKeyAuthentication() {
  const geminiApiKey = requireAuthTestCredential(AUTH_TEST_CONFIG_KEYS.GEMINI_API_KEY, "API key authentication");

  // Force the Gemini public API-key path and clear Vertex settings so this
  // test cannot accidentally pass with Vertex AI credentials.
  GenAIApp.setGeminiAuth(null, null);
  GenAIApp.setGeminiAPIKey(geminiApiKey);

  const chat = GenAIApp.newChat();
  chat.addMessage("Reply with exactly: API key authentication ok");
  const response = chat.run({ model: GEMINI_MODEL, max_tokens: 512 });
  console.log(`[GenAIApp tests] API key authentication completed with response length ${String(response).length}.`);
}

function testVertexAiAuthentication() {
  const projectId = requireAuthTestCredential(AUTH_TEST_CONFIG_KEYS.VERTEX_AI_GCP_PROJECT_ID, "Vertex AI authentication");
  const vertexRegion = getAuthTestConfigValue(AUTH_TEST_CONFIG_KEYS.VERTEX_AI_GCP_REGION, "");

  // Force the Vertex AI path and clear any Gemini API key so this test cannot
  // accidentally pass through the Generative Language API.
  GenAIApp.setGeminiAPIKey(null);
  GenAIApp.setGeminiAuth(projectId, vertexRegion);

  const chat = GenAIApp.newChat();
  chat.addMessage("Reply with exactly: Vertex AI authentication ok");
  const response = chat.run({ model: GEMINI_MODEL, max_tokens: 512 });
  console.log(`[GenAIApp tests] Vertex AI authentication completed with response length ${String(response).length}.`);
}

// Run all tests
function testAll() {
  testConfiguredAuthenticationModes();
  GenAIApp.resetGeminiAuthState();
  testSimpleChatInstance();
  testFunctionCalling();
  testFunctionCallingEndWithResult();
  testFunctionCallingOnlyReturnArguments();
  testBrowsing();
  testKnowledgeLink();
  //testVision();
  testMaximumAPICalls();
  testInputTokenWarning();
  // OpenAI-only tests - require valid Drive file IDs.
  if (TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID) {
    testCodeInterpreterExcel(TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID);
  }
  if (TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID) {
    testCodeInterpreterPDF(TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID);
  }
}


// Helper to set API keys and run tests across models
function runTestAcrossModels(testName, setupFunction, runOptions = {}) {
  const geminiApiKey = getAuthTestConfigValue("GEMINI_API_KEY", "");
  const openAiApiKey = getAuthTestConfigValue("OPEN_AI_API_KEY", "");

  if (geminiApiKey) {
    GenAIApp.setGeminiAPIKey(geminiApiKey);
  }

  if (openAiApiKey) {
    GenAIApp.setOpenAIAPIKey(openAiApiKey);
  }

  const models = [
    { name: GPT_MODEL, label: "GPT" },
    { name: REASONING_MODEL, label: "reasoning" },
    { name: GEMINI_MODEL, label: "gemini" }
  ];

  models.forEach(model => {
    const chat = GenAIApp.newChat();
    setupFunction(chat);
    const options = { model: model.name, ...runOptions };
    const response = chat.run(options);
    console.log(`${testName} ${model.label}:\n${response}`);
  });
}

// Test functions using the helper
function testSimpleChatInstance() {
  runTestAcrossModels("Simple chat", chat => {
    chat
      .addMessage("You're name is Tom, you're a Google Developper Expert and always willing to give useful tips. Always answer in a friendly manner, and include one joke at the end of your messages.", true)
      .addMessage("What are the best pratices to document a project?");
  }, { max_tokens: 10000 });
}

function testFunctionCalling() {
  const weatherFunction = GenAIApp.newFunction()
    .setName("getWeather")
    .setDescription("To retrieve the weather in a city in °C")
    .addParameter("cityName", "string", "The name of the city.");

  runTestAcrossModels("Function calling", chat => {
    chat
      .addMessage("What's the weather in Lyon and Paris today?")
      .addFunction(weatherFunction);
  }, { max_tokens: 10000 });
}

function testFunctionCallingEndWithResult() {
  const weatherFunction = GenAIApp.newFunction()
    .setName("getWeather")
    .setDescription("To retrieve the weather in a city in °C")
    .addParameter("cityName", "string", "The name of the city.")
    .endWithResult(true);

  runTestAcrossModels("End-with-result", chat => {
    chat
      .addMessage("Tell me the weather in Paris")
      .addFunction(weatherFunction);
  });
}

function testFunctionCallingOnlyReturnArguments() {
  const emailExtractor = GenAIApp.newFunction()
    .setName("getEmailAddress")
    .setDescription("Extract an email address from text")
    .addParameter("emailAddress", "string", "the email address")
    .onlyReturnArguments(true);

  runTestAcrossModels("Only-return-args", chat => {
    chat
      .addMessage("Here is a support ticket : 'Please contact me at user@example.com'")
      .addMessage("What's the customer email address ? Use getEmailAddress")
      .addFunction(emailExtractor);
  });
}

function testBrowsing() {
  runTestAcrossModels("Browsing", chat => {
    chat
      .addMessage("Find the latest news about Google Apps Script")
      .enableBrowsing(true);
  });
}

function testKnowledgeLink() {
  runTestAcrossModels("Knowledge link", chat => {
    chat
      .addMessage("Summarize the content of the referenced page.")
      .addKnowledgeLink("https://developers.google.com/apps-script");
  });
}

function testVision() {
  runTestAcrossModels("Vision", chat => {
    chat
      .enableVision(true)
      .addMessage("Describe the following image.")
      .addImage(
        "https://good-nature-blog-uploads.s3.amazonaws.com/uploads/2014/02/slide_336579_3401508_free-1200x640.jpg",
        "high"
      );
  });
}

function testMaximumAPICalls() {
  runTestAcrossModels("Max API calls", chat => {
    chat
      .setMaximumAPICalls(2)
      .addMessage("Give me a step by step plan to become an Apps Script expert.");
  });
}


function testInputTokenWarning() {
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  // Case 1: low threshold should log warning (manual log inspection).
  const lowThresholdChat = GenAIApp.newChat();
  lowThresholdChat
    .warnIfResponseTokenUsageAbove(1)
    .addMessage("In one sentence, explain what token usage means for an API call.");
  const lowThresholdResponse = lowThresholdChat.run({ model: GPT_MODEL, max_tokens: 200 });
  console.log(`Input token warning test (low threshold) response:
${lowThresholdResponse}`);
  console.log("Input token warning test (low threshold): verify that a warning log was emitted.");

  // Case 2: high threshold should not log warning (manual log inspection).
  const highThresholdChat = GenAIApp.newChat();
  highThresholdChat
    .warnIfResponseTokenUsageAbove(180)
    .addMessage("In one sentence, explain what token usage means for an API call.");
  const highThresholdResponse = highThresholdChat.run({ model: GPT_MODEL, max_tokens: 200 });
  console.log(`Input token warning test (high threshold) response:
${highThresholdResponse}`);
  console.log("Input token warning test (high threshold): verify that no warning log was emitted.");
}

function testCodeInterpreterExcel(driveFileId) {
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
  const inputBlob = DriveApp.getFileById(driveFileId).getBlob();
  const chat = GenAIApp.newChat();
  chat
    .addFile(inputBlob)
    .enableCodeInterpreter()
    .addMessage("Add a new column at the end that calculates row totals for all numeric columns. Then generate and attach the updated Excel file as output.");
  const response = chat.run({ model: GPT_MODEL, max_tokens: 4000 });
  console.log(`Generated Excel file url: ${response}`);
  console.log(`Generated files:\n${JSON.stringify(chat.getGeneratedFiles())}`);
}

function testCodeInterpreterPDF(driveFileId) {
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
  const inputBlob = DriveApp.getFileById(driveFileId).getBlob();
  const chat = GenAIApp.newChat();
  chat
    .addFile(inputBlob)
    .enableCodeInterpreter()
    .addMessage("Add a summary paragraph at the top of the document describing its main contents. Then generate and attach the updated PDF file as output.");
  const response = chat.run({ model: GPT_MODEL, max_tokens: 4000 });
  console.log(`Generated PDF file url: ${response}`);
  console.log(`Generated files:\n${JSON.stringify(chat.getGeneratedFiles())}`);
}

// Weather function implementation
function getWeather(cityName) {
  return `The weather in ${cityName} is 19°C today.`;
}
