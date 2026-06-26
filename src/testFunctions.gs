const GPT_MODEL = "gpt-5.4";
const REASONING_MODEL = "o4-mini";
const GEMINI_MODEL = "gemini-3.5-flash";
const TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID = "";
const TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID = "";
const TEST_MAX_TOKENS = 20000;
let TEST_MODEL_TARGETS = ["gpt", "thinking", "gemini"];

/**
 * Restrict cross-model tests to one or more model families: "gpt", "thinking", or "gemini".
 * @param {string|string[]} targets - A model family label or list of labels.
 */
function setTestModelTargets(targets) {
  TEST_MODEL_TARGETS = (Array.isArray(targets) ? targets : [targets])
    .map(target => String(target).toLowerCase());
}

function testAllGpt() {
  setTestModelTargets("gpt");
  testAll();
}

function testAllThinking() {
  setTestModelTargets("thinking");
  testAll();
}

function testAllGemini() {
  setTestModelTargets("gemini");
  testAll();
}

function testAllModels() {
  setTestModelTargets(["gpt", "thinking", "gemini"]);
  testAll();
}

function _shouldRunModelLabel(label) {
  return TEST_MODEL_TARGETS.indexOf(String(label).toLowerCase()) !== -1;
}


function _isNonEmptyResponse(response) {
  if (typeof response === "string") return response.trim().length > 0;
  return response !== null && response !== undefined;
}

function _logTestResult(testName, modelLabel, passed, details = "") {
  const suffix = details ? ` - ${details}` : "";
  console.log(`${passed ? "PASS" : "FAIL"}: ${testName} [${modelLabel}]${suffix}`);
}

function _runSingleTest(testName, modelLabel, testFunction) {
  try {
    const details = testFunction();
    _logTestResult(testName, modelLabel, true, details);
  }
  catch (err) {
    _logTestResult(testName, modelLabel, false, err && err.message ? err.message : String(err));
  }
}

// Run all tests
function testAll() {
  testSimpleChatInstance();
  testFunctionCalling();
  testFunctionCallingEndWithResult();
  testFunctionCallingOnlyReturnArguments();
  testBrowsing();
  testKnowledgeLink();
  testMaximumAPICalls();
  testInputTokenWarning();
  if (_shouldRunModelLabel("gemini")) {
    testGeminiInteractionThreading();
    testGeminiRetrieveLastInteractionId();
    testGeminiFunctionCallingInteractionContinuation();
  }
  // OpenAI-only tests - require valid Drive file IDs.
  if (_shouldRunModelLabel("gpt") && TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID) {
    testCodeInterpreterExcel(TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID);
  }
  if (_shouldRunModelLabel("gpt") && TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID) {
    testCodeInterpreterPDF(TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID);
  }
}


// Helper to set API keys and run tests across models
function runTestAcrossModels(testName, setupFunction, runOptions = {}, validateResponse = _isNonEmptyResponse) {
  // Set API keys once per batch
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  const models = [
    { name: GPT_MODEL, label: "gpt" },
    { name: REASONING_MODEL, label: "thinking" },
    { name: GEMINI_MODEL, label: "gemini" }
  ].filter(model => _shouldRunModelLabel(model.label));

  models.forEach(model => {
    _runSingleTest(testName, model.label, () => {
      const chat = GenAIApp.newChat().disableLogs(true);
      setupFunction(chat);
      const response = chat.run({ model: model.name, ...runOptions, max_tokens: runOptions.max_tokens ?? TEST_MAX_TOKENS });
      if (!validateResponse(response, chat, model)) {
        throw new Error("Unexpected response");
      }
      return "OK";
    });
  });
}

// Test functions using the helper
function testSimpleChatInstance() {
  runTestAcrossModels("Simple chat", chat => {
    chat
      .addMessage("You're name is Tom, you're a Google Developper Expert and always willing to give useful tips. Always answer in a friendly manner, and include one joke at the end of your messages.", true)
      .addMessage("What are the best pratices to document a project?");
  }, { max_tokens: TEST_MAX_TOKENS });
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
  }, { max_tokens: TEST_MAX_TOKENS });
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
  }, {}, response => response === "OK");
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
  }, {}, response => JSON.stringify(response).indexOf("user@example.com") !== -1);
}

function testBrowsing() {
  runTestAcrossModels("Browsing", chat => {
    chat
      .addMessage("Find the latest news about Google Apps Script")
      .enableBrowsing(true);
  }, { max_tokens: TEST_MAX_TOKENS });
}

function testKnowledgeLink() {
  runTestAcrossModels("Knowledge link", chat => {
    chat
      .addMessage("Summarize the content of the referenced page.")
      .addKnowledgeLink("https://developers.google.com/apps-script");
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
  if (!_shouldRunModelLabel("gpt")) {
    _logTestResult("Input token warning", "gpt", true, "skipped");
    return;
  }
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  _runSingleTest("Input token warning", "gpt", () => {
    const chat = GenAIApp.newChat().disableLogs(true);
    chat
      .warnIfResponseTokenUsageAbove(1000000)
      .addMessage("In one sentence, explain what token usage means for an API call.");
    const response = chat.run({ model: GPT_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(response) || !chat._lastUsage) {
      throw new Error("Expected a response and usage information");
    }
    return "OK";
  });
}

function testCodeInterpreterExcel(driveFileId) {
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
  const inputBlob = DriveApp.getFileById(driveFileId).getBlob();
  const chat = GenAIApp.newChat().disableLogs(true);
  chat
    .addFile(inputBlob)
    .enableCodeInterpreter()
    .addMessage("Add a new column at the end that calculates row totals for all numeric columns. Then generate and attach the updated Excel file as output.");
  _runSingleTest("Code interpreter Excel", "gpt", () => {
    const response = chat.run({ model: GPT_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(response) || chat.getGeneratedFiles().length === 0) {
      throw new Error("Expected a generated file");
    }
    return "OK";
  });
}

function testCodeInterpreterPDF(driveFileId) {
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
  const inputBlob = DriveApp.getFileById(driveFileId).getBlob();
  const chat = GenAIApp.newChat().disableLogs(true);
  chat
    .addFile(inputBlob)
    .enableCodeInterpreter()
    .addMessage("Add a summary paragraph at the top of the document describing its main contents. Then generate and attach the updated PDF file as output.");
  _runSingleTest("Code interpreter PDF", "gpt", () => {
    const response = chat.run({ model: GPT_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(response) || chat.getGeneratedFiles().length === 0) {
      throw new Error("Expected a generated file");
    }
    return "OK";
  });
}

// Weather function implementation
function getWeather(cityName) {
  return `The weather in ${cityName} is 19°C today.`;
}

function testGeminiInteractionThreading() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  _runSingleTest("Gemini interaction threading", "gemini", () => {
    const chat = GenAIApp.newChat().disableLogs(true);
    chat.addMessage("Remember this keyword for the next turn: papaya.");
    const firstResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const interactionId = chat.retrieveLastInteractionId();
    if (!_isNonEmptyResponse(firstResponse) || !interactionId) {
      throw new Error("Expected first response and interaction ID");
    }
    chat.addMessage("What keyword did I ask you to remember?");
    const secondResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(secondResponse)) {
      throw new Error("Expected threaded response");
    }
    return "OK";
  });
}

function testGeminiRetrieveLastInteractionId() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  _runSingleTest("Gemini retrieve last interaction ID", "gemini", () => {
    const chat = GenAIApp.newChat().disableLogs(true);
    chat.addMessage("Reply with one short sentence about Apps Script.");
    const response = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const interactionId = chat.retrieveLastInteractionId();
    if (!_isNonEmptyResponse(response) || typeof interactionId !== "string" || interactionId.length === 0) {
      throw new Error("Expected response and valid interaction ID");
    }
    return "OK";
  });
}

function testGeminiFunctionCallingInteractionContinuation() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  _runSingleTest("Gemini function continuation", "gemini", () => {
    const weatherFunction = GenAIApp.newFunction()
      .setName("getWeather")
      .setDescription("To retrieve the weather in a city in °C")
      .addParameter("cityName", "string", "The name of the city.");

    const chat = GenAIApp.newChat().disableLogs(true);
    chat
      .addMessage("What's the weather in Paris? Use the available function, then answer normally.")
      .addFunction(weatherFunction);
    const firstResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const firstInteractionId = chat.retrieveLastInteractionId();
    if (!_isNonEmptyResponse(firstResponse) || !firstInteractionId) {
      throw new Error("Expected function-call response and interaction ID");
    }

    chat.addMessage("Continue from the previous interaction: which city did we just discuss?");
    const secondResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const secondInteractionId = chat.retrieveLastInteractionId();
    if (!_isNonEmptyResponse(secondResponse) || !secondInteractionId) {
      throw new Error("Expected continuation response and interaction ID");
    }
    return "OK";
  });
}
