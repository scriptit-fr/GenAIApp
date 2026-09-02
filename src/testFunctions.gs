const GPT_MODEL = "gpt-5.6-terra";
const REASONING_MODEL = "gpt-5.6-sol";
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
    testGeminiInteractionRequestPayloads();
    testGeminiFailedInteractionState();
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

function _mockGeminiApiCaller(responses, requests) {
  let responseIndex = 0;
  return (endpoint, payload) => {
    requests.push({ endpoint, payload: JSON.parse(JSON.stringify(payload)) });
    if (responseIndex >= responses.length) {
      throw new Error("Unexpected mocked Gemini request");
    }
    return responses[responseIndex++];
  };
}

function _geminiTextResponse(id, text, status = "completed") {
  return {
    id,
    status,
    steps: text ? [{ type: "model_output", content: [{ type: "text", text }] }] : []
  };
}

function testGeminiInteractionRequestPayloads() {
  GenAIApp.setGeminiAPIKey("mock-gemini-key");
  _runSingleTest("Gemini stateful interaction payloads", "gemini", () => {
    const requests = [];
    const chat = GenAIApp.newChat().disableLogs(true);
    chat._apiCaller = _mockGeminiApiCaller([
      _geminiTextResponse("interaction-1", "Remembered papaya."),
      _geminiTextResponse("interaction-2", "The keyword was papaya.")
    ], requests);

    chat.addMessage("Remember this keyword: papaya.");
    const firstResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const firstInteractionId = chat.retrieveLastInteractionId();
    if (!_isNonEmptyResponse(firstResponse) || firstInteractionId !== "interaction-1") {
      throw new Error("Expected first response and interaction ID");
    }

    chat.addMessage("What keyword did I ask you to remember?");
    const secondResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(secondResponse) || chat.retrieveLastInteractionId() !== "interaction-2") {
      throw new Error("Expected threaded response and interaction ID");
    }
    if (requests[0].payload.store !== true || requests[1].payload.store !== true) {
      throw new Error("Gemini interaction requests must set store to true");
    }
    if (requests[0].payload.previous_interaction_id !== undefined
      || requests[1].payload.previous_interaction_id !== "interaction-1") {
      throw new Error("Expected previous_interaction_id only on the continuation request");
    }
    if (requests[1].payload.input.length !== 1
      || requests[1].payload.input[0]?.content?.[0]?.text !== "What keyword did I ask you to remember?") {
      throw new Error("Continuation input must contain only content added after the prior interaction");
    }

    const functionRequests = [];
    const functionChat = GenAIApp.newChat().disableLogs(true);
    const weatherFunction = GenAIApp.newFunction()
      .setName("getWeather")
      .setDescription("Get weather")
      .addParameter("cityName", "string", "City name");
    functionChat._apiCaller = _mockGeminiApiCaller([
      {
        id: "function-interaction-1",
        status: "completed",
        steps: [
          { type: "thought", signature: "opaque-weather-signature" },
          {
            type: "function_call",
            id: "weather-call-1",
            name: "getWeather",
            args: { cityName: "Paris" }
          }
        ]
      },
      _geminiTextResponse("function-interaction-2", "It is 19°C in Paris.")
    ], functionRequests);
    functionChat.addMessage("What's the weather in Paris?").addFunction(weatherFunction);
    const functionResponse = functionChat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(functionResponse) || functionChat.retrieveLastInteractionId() !== "function-interaction-2") {
      throw new Error("Expected function continuation response and interaction ID");
    }
    const functionContinuation = functionRequests[1].payload;
    if (functionContinuation.previous_interaction_id !== "function-interaction-1"
      || functionContinuation.input.length !== 1
      || functionContinuation.input[0].type !== "function_result"
      || functionContinuation.input[0].call_id !== "weather-call-1"
      || functionContinuation.input[0].result?.[0]?.text !== "The weather in Paris is 19°C today.") {
      throw new Error("Function-result continuation did not preserve the expected delta input");
    }
    if (functionContinuation.input[0].thought_signature !== undefined) {
      throw new Error("Stored interaction continuations must not copy the thought signature onto function results");
    }
    if (functionChat.retrieveLastThoughtSignature() !== "opaque-weather-signature") {
      throw new Error("Expected the Gemini thought signature to remain retrievable");
    }
    return "OK";
  });
}

function testGeminiFailedInteractionState() {
  GenAIApp.setGeminiAPIKey("mock-gemini-key");
  _runSingleTest("Gemini failed interaction state", "gemini", () => {
    const requests = [];
    const chat = GenAIApp.newChat().disableLogs(true);
    chat._apiCaller = _mockGeminiApiCaller([
      _geminiTextResponse("valid-interaction", "First response."),
      _geminiTextResponse("failed-interaction", "", "failed"),
      _geminiTextResponse("recovered-interaction", "Recovered response.")
    ], requests);

    chat.addMessage("First turn.");
    const firstResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(firstResponse) || chat.retrieveLastInteractionId() !== "valid-interaction") {
      throw new Error("Expected first response and interaction ID");
    }
    chat.addMessage("This turn receives a mocked HTTP 200 failed interaction.");
    chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (chat.retrieveLastInteractionId() !== "valid-interaction") {
      throw new Error("Failed interaction replaced the last valid interaction ID");
    }
    chat.addMessage("Retry after failure.");
    const recoveredResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(recoveredResponse) || chat.retrieveLastInteractionId() !== "recovered-interaction") {
      throw new Error("Expected recovered response and interaction ID");
    }
    if (requests[2].payload.previous_interaction_id !== "valid-interaction"
      || requests[2].payload.previous_interaction_id === "failed-interaction") {
      throw new Error("A failed interaction ID was used as a continuation handle");
    }
    if (requests[2].payload.input.length !== 2) {
      throw new Error("Failed interaction advanced the Gemini content boundary");
    }
    return "OK";
  });
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
    if (!_isNonEmptyResponse(secondResponse) || !secondInteractionId || secondInteractionId === firstInteractionId) {
      throw new Error("Expected continuation response and interaction ID");
    }
    if (!/paris/i.test(secondResponse)) {
      throw new Error("Gemini function-call continuation did not preserve context.");
    }
    return "OK";
  });
}
