const OPENAI_MODEL = "gpt-5.6-terra";
const GEMINI_MODEL = "gemini-3.5-flash";
const TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID = "";
const TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID = "";
const TEST_MAX_TOKENS = 20000;
let TEST_PROVIDER_TARGETS = ["openai", "gemini"];

/**
 * Restrict cross-provider tests to "openai", "gemini", or both.
 * @param {string|string[]} targets - A provider label or list of labels.
 */
function setTestProviderTargets(targets) {
  TEST_PROVIDER_TARGETS = (Array.isArray(targets) ? targets : [targets])
    .map(target => String(target).toLowerCase());
}

function testAllOpenAI() {
  setTestProviderTargets("openai");
  testAll();
}

function testAllGemini() {
  setTestProviderTargets("gemini");
  testAll();
}

function testAllProviders() {
  setTestProviderTargets(["openai", "gemini"]);
  testAll();
}

function _shouldRunProvider(providerName) {
  return TEST_PROVIDER_TARGETS.indexOf(String(providerName).toLowerCase()) !== -1;
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
  testMCPConnectorPayloads();
  testReasoningLevelPayloads();
  testVectorStoreStateIsolation();
  testOpenAIToolContinuationState();
  testSimpleChatInstance();
  testFunctionCalling();
  testFunctionCallingEndWithResult();
  testFunctionCallingOnlyReturnArguments();
  testBrowsing();
  testKnowledgeLink();
  testMaximumAPICalls();
  testInputTokenWarning();
  if (_shouldRunProvider("gemini")) {
    testGeminiInteractionRequestPayloads();
    testGeminiBuiltInToolCallsAreNotDispatchedLocally();
    testGeminiDeferredBuiltInToolCallsWithoutLocalFunctions();
    testGeminiGlobalFunctionCallsRemainEligible();
    testGeminiFailedInteractionState();
    testGeminiInteractionThreading();
    testGeminiRetrieveLastInteractionId();
    testGeminiFunctionCallingInteractionContinuation();
  }
  // OpenAI-only tests - require valid Drive file IDs.
  if (_shouldRunProvider("openai") && TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID) {
    testCodeInterpreterExcel(TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID);
  }
  if (_shouldRunProvider("openai") && TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID) {
    testCodeInterpreterPDF(TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID);
  }
}

function testMCPConnectorPayloads() {
  _runSingleTest("MCP connector payloads", "local", () => {
    const remote = GenAIApp.newConnector()
      .setLabel("remote")
      .setServerUrl("https://mcp.example.com")
      ._toJson();
    if (remote.server_url !== "https://mcp.example.com" || "tunnel_id" in remote || "connector_id" in remote) {
      throw new Error("Expected remote MCP payload to use server_url only");
    }

    const tunnel = GenAIApp.newConnector()
      .setLabel("local")
      .setTunnelId("tunnel_test")
      ._toJson();
    if (tunnel.tunnel_id !== "tunnel_test" || "server_url" in tunnel || "connector_id" in tunnel) {
      throw new Error("Expected local MCP payload to use tunnel_id only");
    }

    let tunnelRejectedByGemini = false;
    try {
      GenAIApp.newConnector().setTunnelId("tunnel_test")._toGeminiJson();
    }
    catch (err) {
      tunnelRejectedByGemini = /only supported for OpenAI/.test(err.message);
    }
    if (!tunnelRejectedByGemini) {
      throw new Error("Expected Gemini payload builder to reject tunnel_id");
    }
    return "OK";
  });
}

function testReasoningLevelPayloads() {
  _runSingleTest("Reasoning level payloads", "local", () => {
    const openAIPayload = GenAIApp.newChat()
      .setReasoningLevel("high")
      ._buildOpenAIPayload();
    if (openAIPayload.reasoning?.effort !== "high") {
      throw new Error("Expected OpenAI reasoning.effort to use the generic reasoning level");
    }

    const geminiPayload = GenAIApp.newChat()
      .setReasoningLevel("low")
      ._buildGeminiPayload({});
    if (geminiPayload.generation_config?.thinking_level !== "low") {
      throw new Error("Expected Gemini thinking_level to use the generic reasoning level");
    }
    return "OK";
  });
}

function testVectorStoreStateIsolation() {
  _runSingleTest("Vector store state isolation", "local", () => {
    const storeChatPayload = GenAIApp.newChat()
      .addVectorStores("fileSearchStores/test-store")
      ._buildGeminiPayload({});
    if (storeChatPayload.tools?.[0]?.file_search_store_names?.[0] !== "fileSearchStores/test-store") {
      throw new Error("Expected vector store on configured chat");
    }

    const cleanChatPayload = GenAIApp.newChat()._buildGeminiPayload({});
    if (cleanChatPayload.tools.some(tool => tool.type === "file_search")) {
      throw new Error("Vector store leaked into another chat");
    }
    return "OK";
  });
}

function testOpenAIToolContinuationState() {
  _runSingleTest("OpenAI tool continuation state", "local", () => {
    GenAIApp.configureProvider("openai", { apiKey: "mock-openai-key" });
    const requests = [];
    const responses = [
      {
        id: "response-tool-call",
        output: [{
          type: "function_call",
          name: "getWeather",
          arguments: JSON.stringify({ cityName: "Paris" }),
          call_id: "weather-call"
        }]
      },
      {
        id: "response-final",
        output: [{
          type: "message",
          status: "final_answer",
          content: [{ type: "output_text", text: "It is 19°C in Paris." }]
        }]
      },
      {
        id: "response-follow-up",
        output: [{
          type: "message",
          status: "final_answer",
          content: [{ type: "output_text", text: "Paris." }]
        }]
      }
    ];
    const chat = GenAIApp.newChat().disableLogs(true);
    chat._apiCaller = (endpoint, payload) => {
      requests.push({ endpoint, payload: JSON.parse(JSON.stringify(payload)) });
      return responses.shift();
    };
    chat
      .addMessage("What's the weather in Paris?")
      .addFunction(GenAIApp.newFunction()
        .setName("getWeather")
        .setDescription("Get weather")
        .addParameter("cityName", "string", "City name"));

    chat.run({ model: OPENAI_MODEL, max_tokens: TEST_MAX_TOKENS });
    chat.addMessage("Which city did we discuss?");
    chat.run({ model: OPENAI_MODEL, max_tokens: TEST_MAX_TOKENS });

    const followUpPayload = requests[2].payload;
    if (followUpPayload.previous_response_id !== "response-final") {
      throw new Error("Expected follow-up to continue from the final response");
    }
    if (followUpPayload.input.some(item => item.type === "function_call_output")) {
      throw new Error("Consumed function output leaked into follow-up input");
    }
    return "OK";
  });
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

let geminiUnregisteredFunctionCallCount = 0;
function unregisteredGeminiFunction() {
  geminiUnregisteredFunctionCallCount++;
}

function testGeminiBuiltInToolCallsAreNotDispatchedLocally() {
  GenAIApp.configureProvider("gemini", { apiKey: "mock-gemini-key" });
  _runSingleTest("Gemini built-in tool dispatch", "gemini", () => {
    geminiUnregisteredFunctionCallCount = 0;
    const requests = [];
    const chat = GenAIApp.newChat().disableLogs(true);
    const localFunction = GenAIApp.newFunction()
      .setName("getWeather")
      .setDescription("Get weather")
      .addParameter("cityName", "string", "City name");
    chat._apiCaller = _mockGeminiApiCaller([{
      id: "file-search-interaction",
      status: "completed",
      steps: [
        {
          type: "function_call",
          id: "built-in-file-search-call",
          name: "google:file_search",
          args: { queries: ["team demos"] }
        },
        {
          type: "function_call",
          id: "unregistered-global-call",
          name: "unregisteredGeminiFunction",
          args: {}
        },
        {
          type: "model_output",
          content: [{ type: "text", text: "Team demos happen every Friday at 10 AM." }]
        }
      ]
    }], requests);

    chat.addMessage("When are team demos?").addFunction(localFunction);
    const response = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (response !== "Team demos happen every Friday at 10 AM.") {
      throw new Error("Gemini built-in file search was treated as a local function call");
    }
    if (requests.length !== 1) {
      throw new Error("Built-in tool activity unexpectedly triggered a continuation request");
    }
    if (geminiUnregisteredFunctionCallCount !== 0) {
      throw new Error("An unregistered global function was dispatched locally");
    }
    return "OK";
  });
}

function testGeminiDeferredBuiltInToolCallsWithoutLocalFunctions() {
  GenAIApp.configureProvider("gemini", { apiKey: "mock-gemini-key" });
  _runSingleTest("Gemini deferred built-in tool without local functions", "gemini", () => {
    const requests = [];
    const chat = GenAIApp.newChat().disableLogs(true);
    chat._apiCaller = _mockGeminiApiCaller([
      {
        id: "deferred-file-search-interaction",
        status: "requires_action",
        steps: [{
          type: "function_call",
          id: "deferred-file-search-call",
          name: "google:file_search",
          args: { queries: ["team demos"] }
        }]
      },
      _geminiTextResponse("completed-file-search-interaction", "Team demos happen every Friday at 10 AM.")
    ], requests);

    chat
      .addVectorStores("fileSearchStores/test-store")
      .addMessage("When are team demos?");
    const response = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (response !== "Team demos happen every Friday at 10 AM.") {
      throw new Error("Deferred built-in File Search did not continue to the final response");
    }
    if (requests.length !== 2
      || requests[1].payload.previous_interaction_id !== "deferred-file-search-interaction"
      || requests[1].payload.input?.[0]?.type !== "function_result"
      || requests[1].payload.input?.[0]?.call_id !== "deferred-file-search-call"
      || requests[1].payload.tools?.[0]?.type !== "file_search") {
      throw new Error("Deferred built-in File Search continuation payload was invalid");
    }
    return "OK";
  });
}

function testGeminiGlobalFunctionCallsRemainEligible() {
  GenAIApp.configureProvider("gemini", { apiKey: "mock-gemini-key" });
  _runSingleTest("Gemini global function dispatch", "gemini", () => {
    const requests = [];
    const chat = GenAIApp.newChat().disableLogs(true);
    const functionDeclaration = GenAIApp.newFunction()
      .setName("getWeather")
      .setDescription("Get weather")
      .addParameter("cityName", "string", "City name");
    chat._apiCaller = _mockGeminiApiCaller([
      {
        id: "function-interaction-1",
        status: "completed",
        steps: [{
          type: "function_call",
          id: "weather-call-1",
          name: "getWeather",
          args: { cityName: "Paris" }
        }]
      },
      _geminiTextResponse("function-interaction-2", "It is 19°C in Paris.")
    ], requests);

    chat.addMessage("What's the weather in Paris?").addFunction(functionDeclaration);
    const response = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (response !== "It is 19°C in Paris.") {
      throw new Error("Gemini local function was not called");
    }
    if (requests.length !== 2
      || requests[1].payload.input?.[0]?.name !== "getWeather"
      || requests[1].payload.input?.[0]?.result?.[0]?.text !== "The weather in Paris is 19°C today.") {
      throw new Error("Gemini local function result was not sent back to the model");
    }
    return "OK";
  });
}

function testGeminiInteractionRequestPayloads() {
  GenAIApp.configureProvider("gemini", { apiKey: "mock-gemini-key" });
  _runSingleTest("Gemini stateful interaction payloads", "gemini", () => {
    const requests = [];
    const chat = GenAIApp.newChat().disableLogs(true);
    chat._apiCaller = _mockGeminiApiCaller([
      _geminiTextResponse("interaction-1", "Remembered papaya."),
      _geminiTextResponse("interaction-2", "The keyword was papaya.")
    ], requests);

    chat.addMessage("Remember this keyword: papaya.");
    const firstResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const firstInteractionId = chat.getLastConversationId();
    if (!_isNonEmptyResponse(firstResponse) || firstInteractionId !== "interaction-1") {
      throw new Error("Expected first response and interaction ID");
    }

    chat.addMessage("What keyword did I ask you to remember?");
    const secondResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(secondResponse) || chat.getLastConversationId() !== "interaction-2") {
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
    if (!_isNonEmptyResponse(functionResponse) || functionChat.getLastConversationId() !== "function-interaction-2") {
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
    return "OK";
  });
}

function testGeminiFailedInteractionState() {
  GenAIApp.configureProvider("gemini", { apiKey: "mock-gemini-key" });
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
    if (!_isNonEmptyResponse(firstResponse) || chat.getLastConversationId() !== "valid-interaction") {
      throw new Error("Expected first response and interaction ID");
    }
    chat.addMessage("This turn receives a mocked HTTP 200 failed interaction.");
    chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (chat.getLastConversationId() !== "valid-interaction") {
      throw new Error("Failed interaction replaced the last valid interaction ID");
    }
    chat.addMessage("Retry after failure.");
    const recoveredResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(recoveredResponse) || chat.getLastConversationId() !== "recovered-interaction") {
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


// Helper to configure providers and run shared tests across them
function runTestAcrossProviders(testName, setupFunction, runOptions = {}, validateResponse = _isNonEmptyResponse) {
  // Set API keys once per batch
  GenAIApp.configureProvider("gemini", { apiKey: GEMINI_API_KEY });
  GenAIApp.configureProvider("openai", { apiKey: OPEN_AI_API_KEY });

  const providers = [
    { name: OPENAI_MODEL, label: "openai" },
    { name: GEMINI_MODEL, label: "gemini" }
  ].filter(provider => _shouldRunProvider(provider.label));

  providers.forEach(provider => {
    _runSingleTest(testName, provider.label, () => {
      const chat = GenAIApp.newChat().disableLogs(true);
      setupFunction(chat);
      const response = chat.run({ model: provider.name, ...runOptions, max_tokens: runOptions.max_tokens ?? TEST_MAX_TOKENS });
      if (!validateResponse(response, chat, provider)) {
        throw new Error("Unexpected response");
      }
      return "OK";
    });
  });
}

// Test functions using the helper
function testSimpleChatInstance() {
  runTestAcrossProviders("Simple chat", chat => {
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

  runTestAcrossProviders("Function calling", chat => {
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

  runTestAcrossProviders("End-with-result", chat => {
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

  runTestAcrossProviders("Only-return-args", chat => {
    chat
      .addMessage("Here is a support ticket : 'Please contact me at user@example.com'")
      .addMessage("What's the customer email address ? Use getEmailAddress")
      .addFunction(emailExtractor);
  }, {}, response => JSON.stringify(response).indexOf("user@example.com") !== -1);
}

function testBrowsing() {
  runTestAcrossProviders("Browsing", chat => {
    chat
      .addMessage("Find the latest news about Google Apps Script")
      .enableBrowsing(true);
  }, { max_tokens: TEST_MAX_TOKENS });
}

function testKnowledgeLink() {
  runTestAcrossProviders("Knowledge link", chat => {
    chat
      .addMessage("Summarize the content of the referenced page.")
      .addKnowledgeLink("https://developers.google.com/apps-script");
  });
}

function testMaximumAPICalls() {
  runTestAcrossProviders("Max API calls", chat => {
    chat
      .setMaximumAPICalls(2)
      .addMessage("Give me a step by step plan to become an Apps Script expert.");
  });
}


function testInputTokenWarning() {
  if (!_shouldRunProvider("openai")) {
    _logTestResult("Input token warning", "openai", true, "skipped");
    return;
  }
  GenAIApp.configureProvider("openai", { apiKey: OPEN_AI_API_KEY });

  _runSingleTest("Input token warning", "openai", () => {
    const chat = GenAIApp.newChat().disableLogs(true);
    chat
      .warnIfResponseTokenUsageAbove(1000000)
      .addMessage("In one sentence, explain what token usage means for an API call.");
    const response = chat.run({ model: OPENAI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(response) || !chat._lastUsage) {
      throw new Error("Expected a response and usage information");
    }
    return "OK";
  });
}

function testCodeInterpreterExcel(driveFileId) {
  GenAIApp.configureProvider("openai", { apiKey: OPEN_AI_API_KEY });
  const inputBlob = DriveApp.getFileById(driveFileId).getBlob();
  const chat = GenAIApp.newChat().disableLogs(true);
  chat
    .addFile(inputBlob)
    .enableCodeInterpreter()
    .addMessage("Add a new column at the end that calculates row totals for all numeric columns. Then generate and attach the updated Excel file as output.");
  _runSingleTest("Code interpreter Excel", "openai", () => {
    const response = chat.run({ model: OPENAI_MODEL, max_tokens: TEST_MAX_TOKENS });
    if (!_isNonEmptyResponse(response) || chat.getGeneratedFiles().length === 0) {
      throw new Error("Expected a generated file");
    }
    return "OK";
  });
}

function testCodeInterpreterPDF(driveFileId) {
  GenAIApp.configureProvider("openai", { apiKey: OPEN_AI_API_KEY });
  const inputBlob = DriveApp.getFileById(driveFileId).getBlob();
  const chat = GenAIApp.newChat().disableLogs(true);
  chat
    .addFile(inputBlob)
    .enableCodeInterpreter()
    .addMessage("Add a summary paragraph at the top of the document describing its main contents. Then generate and attach the updated PDF file as output.");
  _runSingleTest("Code interpreter PDF", "openai", () => {
    const response = chat.run({ model: OPENAI_MODEL, max_tokens: TEST_MAX_TOKENS });
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
  GenAIApp.configureProvider("gemini", { apiKey: GEMINI_API_KEY });
  _runSingleTest("Gemini interaction threading", "gemini", () => {
    const chat = GenAIApp.newChat().disableLogs(true);
    chat.addMessage("Remember this keyword for the next turn: papaya.");
    const firstResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const interactionId = chat.getLastConversationId();
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
  GenAIApp.configureProvider("gemini", { apiKey: GEMINI_API_KEY });
  _runSingleTest("Gemini retrieve last interaction ID", "gemini", () => {
    const chat = GenAIApp.newChat().disableLogs(true);
    chat.addMessage("Reply with one short sentence about Apps Script.");
    const response = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const interactionId = chat.getLastConversationId();
    if (!_isNonEmptyResponse(response) || typeof interactionId !== "string" || interactionId.length === 0) {
      throw new Error("Expected response and valid interaction ID");
    }
    return "OK";
  });
}

function testGeminiFunctionCallingInteractionContinuation() {
  GenAIApp.configureProvider("gemini", { apiKey: GEMINI_API_KEY });
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
    const firstInteractionId = chat.getLastConversationId();
    if (!_isNonEmptyResponse(firstResponse) || !firstInteractionId) {
      throw new Error("Expected function-call response and interaction ID");
    }

    chat.addMessage("Continue from the previous interaction: which city did we just discuss?");
    const secondResponse = chat.run({ model: GEMINI_MODEL, max_tokens: TEST_MAX_TOKENS });
    const secondInteractionId = chat.getLastConversationId();
    if (!_isNonEmptyResponse(secondResponse) || !secondInteractionId || secondInteractionId === firstInteractionId) {
      throw new Error("Expected continuation response and interaction ID");
    }
    if (!/paris/i.test(secondResponse)) {
      throw new Error("Gemini function-call continuation did not preserve context.");
    }
    return "OK";
  });
}
