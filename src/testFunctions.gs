const GPT_MODEL = "gpt-4.1";
const REASONING_MODEL = "o4-mini";
const GEMINI_MODEL = "gemini-2.5-pro";

// Run all tests
function testAll() {
  testSimpleChatInstance();
  testFunctionCalling();
  testFunctionCallingEndWithResult();
  testFunctionCallingOnlyReturnArguments();
  testBrowsing();
  testKnowledgeLink();
  testVision();
  testMaximumAPICalls();
  testAddFile();
  testMetadataAndLogs();
  testResponseIdTracking();
  testVectorStoreLifecycle();
}


// Helper to set API keys and run tests across models
function runTestAcrossModels(testName, setupFunction, runOptions = {}, afterRun) {
  // Set API keys once per batch
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

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
    if (afterRun) {
      afterRun(chat, response, model);
    }
  });
}

// Test functions using the helper
function testSimpleChatInstance() {
  runTestAcrossModels("Simple chat", chat => {
    chat
      .addMessage("You're name is Tom, you're a Google Developper Expert and always willing to give useful tips. Always answer in a friendly manner, and include one joke at the end of your messages.", true)
      .addMessage("What are the best pratices to document a project?");
  }, { max_tokens: 1000 });
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
  }, { max_tokens: 1000 });
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

// Weather function implementation
function getWeather(cityName) {
  return `The weather in ${cityName} is 19°C today.`;
}

function testAddFile() {
  const blob = Utilities.newBlob("Hello from a file", "text/plain", "hello.txt");
  runTestAcrossModels("Add file", chat => {
    chat
      .addMessage("Read the file and summarize its content.")
      .addFile(blob);
  });
}

function testMetadataAndLogs() {
  GenAIApp.setGlobalMetadata("suite", "tests");
  const dummyFunction = GenAIApp.newFunction()
    .setName("dummy")
    .setDescription("A dummy function");

  runTestAcrossModels(
    "Metadata and logs",
    chat => {
      chat
        .disableLogs(true)
        .addMetadata("requestId", "12345")
        .addMessage("Say hello")
        .addFunction(dummyFunction);
    },
    {},
    chat => {
      console.log(`Messages: ${chat.getMessages()}`);
      console.log(`Functions: ${chat.getFunctions()}`);
    }
  );
}

function testResponseIdTracking() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  const models = [
    { name: GPT_MODEL, label: "GPT" },
    { name: REASONING_MODEL, label: "reasoning" },
    { name: GEMINI_MODEL, label: "gemini" }
  ];

  models.forEach(model => {
    const chat = GenAIApp.newChat();
    chat.addMessage("Hello");
    chat.run({ model: model.name });
    const lastId = chat.retrieveLastResponseId();
    chat.addMessage("Continue this conversation.");
    if (lastId) {
      chat.setPreviousResponseId(lastId);
    }
    const response = chat.run({ model: model.name });
    console.log(`Response ID test ${model.label}:\n${response}`);
  });
}

function testVectorStoreLifecycle() {
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
  const blob = Utilities.newBlob("Vector store content", "text/plain", "vs.txt");

  const store = GenAIApp.newVectorStore()
    .setName("test-store-" + Date.now())
    .setDescription("Temporary store for tests")
    .createVectorStore();

  const storeId = store.getId();
  const fileId = store.uploadAndAttachFile(blob, { source: "test" });
  const files = store.listFiles();
  console.log(`Vector store files: ${JSON.stringify(files)}`);

  store.deleteFile(fileId);
  store.deleteVectorStore();
  console.log(`Vector store ${storeId} cleaned up.`);
}

