const GPT_MODEL = "gpt-5.4";
const REASONING_MODEL = "o4-mini";
const GEMINI_MODEL = "gemini-2.5-pro";
const TEST_CODE_INTERPRETER_XLSX_DRIVE_FILE_ID = "";
const TEST_CODE_INTERPRETER_PDF_DRIVE_FILE_ID = "";

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
  testInputTokenWarning();
  testGeminiCodeExecution();
  testGeminiCodeExecutionWithArtifact();
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

function testGeminiCodeExecution() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);

  const chat = GenAIApp.newChat();
  chat
    .enableCodeInterpreter()
    .addMessage("Use code execution to compute the mean, median, and sum of this list: [4, 8, 15, 16, 23, 42]. Return the numeric results in your final answer.");

  const response = chat.run({ model: GEMINI_MODEL, max_tokens: 4000 });
  if (!response) {
    throw new Error("Gemini code execution test failed: expected a response.");
  }
  if (chat.getContainerId() !== null) {
    throw new Error("Gemini code execution test failed: Gemini should not return a container ID.");
  }

  const generatedFiles = chat.getGeneratedFiles();
  console.log(`Gemini code execution response:
${response}`);
  console.log(`Gemini code execution generated files:
${JSON.stringify(generatedFiles)}`);

  generatedFiles.forEach((artifact, index) => {
    assertGeminiArtifactMetadata(artifact, `Gemini code execution artifact ${index}`);
    const blob = chat.downloadGeneratedFile(index);
    assertBlobLike(blob, `Gemini code execution artifact ${index}`);
  });
}

function testGeminiCodeExecutionWithArtifact() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);

  const chat = GenAIApp.newChat();
  chat
    .enableCodeInterpreter()
    .addMessage("Use code execution to create a PNG bar chart file showing quarterly revenue values Q1=12, Q2=18, Q3=9, Q4=24. Return the chart as an output artifact.");

  const response = chat.run({ model: GEMINI_MODEL, max_tokens: 4000 });
  if (!response) {
    throw new Error("Gemini artifact test failed: expected a response.");
  }

  const generatedFiles = chat.getGeneratedFiles();
  console.log(`Gemini code execution artifact response:
${response}`);
  console.log(`Gemini code execution artifact files:
${JSON.stringify(generatedFiles)}`);

  if (generatedFiles.length === 0) {
    throw new Error("Gemini artifact test failed: expected at least one generated artifact.");
  }

  generatedFiles.forEach((artifact, index) => {
    assertGeminiArtifactMetadata(artifact, `Gemini artifact ${index}`);
    const blob = chat.downloadGeneratedFile(artifact.filename);
    assertBlobLike(blob, `Gemini artifact ${index}`);
    if (blob.getContentType() !== artifact.mimeType) {
      throw new Error(`Gemini artifact ${index} failed: blob mime type ${blob.getContentType()} did not match ${artifact.mimeType}.`);
    }
  });

  const savedFile = DriveApp.createFile(chat.downloadGeneratedFile(0));
  console.log(`Gemini generated artifact file url: ${savedFile.getUrl()}`);
}

function assertGeminiArtifactMetadata(artifact, label) {
  if (!artifact || !artifact.mimeType || !artifact.data || !artifact.filename) {
    throw new Error(`${label} failed: expected mimeType, data, and filename fields.`);
  }
}

function assertBlobLike(blob, label) {
  if (!blob || typeof blob.getBytes !== "function" || blob.getBytes().length === 0) {
    throw new Error(`${label} failed: expected a non-empty Blob.`);
  }
}

// Weather function implementation
function getWeather(cityName) {
  return `The weather in ${cityName} is 19°C today.`;
}
