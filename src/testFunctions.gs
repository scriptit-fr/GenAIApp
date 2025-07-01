// a bunch of functions to test the different features of the library

const GPT_MODEL = "gpt-4.1";
const REASONING_MODEL = "o4-mini";
const GEMINI_MODEL = "gemini-2.5-pro";

function testAll() {
  testSimpleChatInstance();
  testFunctionCalling();
  testFunctionCallingEndWithResult();
  testFunctionCallingOnlyReturnArguments();
  testBrowsing();
  testKnowledgeLink();
  testVision();
  testMaximumAPICalls();
}

function testSimpleChatInstance() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let chatGPT = GenAIApp.newChat();
  chatGPT.addMessage("You're name is Tom, you're a Google Developper Expert and always willing to give useful tips. Always answer in a friendly manner, and include one joke at the end of your messages.", true)
    .addMessage("What are the best pratices to document a project?");

  let reasonningChat = GenAIApp.newChat();
  reasonningChat.addMessage("You're name is Tom, you're a Google Developper Expert and always willing to give useful tips. Always answer in a friendly manner, and include one joke at the end of your messages.", true)
    .addMessage("What are the best pratices to document a project?");

  let geminiChat = GenAIApp.newChat();
  geminiChat.addMessage("You're name is Tom, you're a Google Developper Expert and always willing to give useful tips. Always answer in a friendly manner, and include one joke at the end of your messages.", true)
    .addMessage("What are the best pratices to document a project?");

  let gptResponse = chatGPT.run({ model: GPT_MODEL, max_tokens: 1000 });
  console.log(`Retrieved GPT answer :\n${gptResponse}`);

  let reasonningResponse = reasonningChat.run({ model: REASONING_MODEL, max_tokens: 1000 });
  console.log(`Retrieved reasonning model answer :\n${reasonningResponse}`);

  let geminiResponse = geminiChat.run({ model: GEMINI_MODEL, max_tokens: 1000 });
  console.log(`Retrieved gemini answer :\n${geminiResponse}`);

}

function testFunctionCalling() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let weatherFunction = GenAIApp.newFunction()
    .setName("getWeather")
    .setDescription("To retrieve the weather in a city in °C")
    .addParameter("cityName", "string", "The name of the city.");

  let chatGPT = GenAIApp.newChat();
  chatGPT.addMessage("What's the weather in Lyon and Paris today?")
    .addFunction(weatherFunction);

  let reasonningChat = GenAIApp.newChat();
  reasonningChat.addMessage("What's the weather in Lyon and Paris today?")
    .addFunction(weatherFunction);

  let geminiChat = GenAIApp.newChat();
  geminiChat.addMessage("What's the weather in Lyon and Paris today?")
    .addFunction(weatherFunction);

  let gptResponse = chatGPT.run({ model: GPT_MODEL, max_tokens: 1000 });
  console.log(`Retrieved GPT answer :\n${gptResponse}`);

  let reasonningResponse = reasonningChat.run({ model: REASONING_MODEL, max_tokens: 1000 });
  console.log(`Retrieved reasonning model answer :\n${reasonningResponse}`);

  let geminiResponse = geminiChat.run({ model: GEMINI_MODEL, max_tokens: 1000 });
  console.log(`Retrieved gemini answer :\n${geminiResponse}`);
}

function testFunctionCallingEndWithResult() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let weatherFunction = GenAIApp.newFunction()
    .setName("getWeather")
    .setDescription("To retrieve the weather in a city in °C")
    .addParameter("cityName", "string", "The name of the city.")
    .endWithResult(true);

  let gptChat = GenAIApp.newChat();
  gptChat.addMessage("Tell me the weather in Paris")
    .addFunction(weatherFunction);
  let gptResponse = gptChat.run({ model: GPT_MODEL });
  console.log(`endWithResult GPT:\n${gptResponse}`);

  let reasonChat = GenAIApp.newChat();
  reasonChat.addMessage("Tell me the weather in Paris")
    .addFunction(weatherFunction);
  let reasonResponse = reasonChat.run({ model: REASONING_MODEL });
  console.log(`endWithResult reasonning:\n${reasonResponse}`);

  let geminiChat = GenAIApp.newChat();
  geminiChat.addMessage("Tell me the weather in Paris")
    .addFunction(weatherFunction);
  let geminiResponse = geminiChat.run({ model: GEMINI_MODEL });
  console.log(`endWithResult gemini:\n${geminiResponse}`);
}

function testFunctionCallingOnlyReturnArguments() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let emailExtractor = GenAIApp.newFunction()
    .setName("getEmailAddress")
    .setDescription("Extract an email address from text")
    .addParameter("emailAddress", "string", "the email address")
    .onlyReturnArguments(true);

  let gptChat = GenAIApp.newChat();
  gptChat.addMessage("Here is a support ticket : 'Please contact me at user@example.com'")
    .addMessage("What's the customer email address ? Use getEmailAddress")
    .addFunction(emailExtractor);
  let gptArgs = gptChat.run({ model: GPT_MODEL });
  console.log(`onlyReturnArguments GPT:\n${gptArgs}`);

  let reasonChat = GenAIApp.newChat();
  reasonChat.addMessage("Here is a support ticket : 'Please contact me at user@example.com'")
    .addMessage("What's the customer email address ? Use getEmailAddress")
    .addFunction(emailExtractor);
  let reasonArgs = reasonChat.run({ model: REASONING_MODEL });
  console.log(`onlyReturnArguments reasonning:\n${reasonArgs}`);

  let geminiChat = GenAIApp.newChat();
  geminiChat.addMessage("Here is a support ticket : 'Please contact me at user@example.com'")
    .addMessage("What's the customer email address ? Use getEmailAddress")
    .addFunction(emailExtractor);
  let geminiArgs = geminiChat.run({ model: GEMINI_MODEL });
  console.log(`onlyReturnArguments gemini:\n${geminiArgs}`);
}

function testBrowsing() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let gptChat = GenAIApp.newChat();
  gptChat.addMessage("Find the latest news about Google Apps Script")
    .enableBrowsing(true);
  let gptAnswer = gptChat.run({ model: GPT_MODEL });
  console.log(`Browsing GPT:\n${gptAnswer}`);

  let reasonChat = GenAIApp.newChat();
  reasonChat.addMessage("Find the latest news about Google Apps Script")
    .enableBrowsing(true);
  let reasonAnswer = reasonChat.run({ model: REASONING_MODEL });
  console.log(`Browsing reasonning:\n${reasonAnswer}`);

  let geminiChat = GenAIApp.newChat();
  geminiChat.addMessage("Find the latest news about Google Apps Script")
    .enableBrowsing(true);
  let geminiAnswer = geminiChat.run({ model: GEMINI_MODEL });
  console.log(`Browsing gemini:\n${geminiAnswer}`);
}

function testKnowledgeLink() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let gptChat = GenAIApp.newChat();
  gptChat.addMessage("Summarize the content of the referenced page.")
    .addKnowledgeLink("https://developers.google.com/apps-script");
  let gptAnswer = gptChat.run({ model: GPT_MODEL });
  console.log(`Knowledge link GPT:\n${gptAnswer}`);

  let reasonChat = GenAIApp.newChat();
  reasonChat.addMessage("Summarize the content of the referenced page.")
    .addKnowledgeLink("https://developers.google.com/apps-script");
  let reasonAnswer = reasonChat.run({ model: REASONING_MODEL });
  console.log(`Knowledge link reasonning:\n${reasonAnswer}`);

  let geminiChat = GenAIApp.newChat();
  geminiChat.addMessage("Summarize the content of the referenced page.")
    .addKnowledgeLink("https://developers.google.com/apps-script");
  let geminiAnswer = geminiChat.run({ model: GEMINI_MODEL });
  console.log(`Knowledge link gemini:\n${geminiAnswer}`);
}

function testVision() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let gptChat = GenAIApp.newChat();
  gptChat.enableVision(true);
  gptChat.addMessage("Describe the following image.")
    .addImage("https://good-nature-blog-uploads.s3.amazonaws.com/uploads/2014/02/slide_336579_3401508_free-1200x640.jpg", "high");
  let gptAnswer = gptChat.run({ model: GPT_MODEL });
  console.log(`Vision GPT:\n${gptAnswer}`);

  let reasonChat = GenAIApp.newChat();
  reasonChat.enableVision(true);
  reasonChat.addMessage("Describe the following image.")
    .addImage("https://good-nature-blog-uploads.s3.amazonaws.com/uploads/2014/02/slide_336579_3401508_free-1200x640.jpg", "high");
  let reasonAnswer = reasonChat.run({ model: REASONING_MODEL });
  console.log(`Vision reasonning:\n${reasonAnswer}`);

  let geminiChat = GenAIApp.newChat();
  geminiChat.enableVision(true);
  geminiChat.addMessage("Describe the following image.")
    .addImage("https://good-nature-blog-uploads.s3.amazonaws.com/uploads/2014/02/slide_336579_3401508_free-1200x640.jpg", "high");
  let geminiAnswer = geminiChat.run({ model: GEMINI_MODEL });
  console.log(`Vision gemini:\n${geminiAnswer}`);
}

function testMaximumAPICalls() {
  GenAIApp.setGeminiAPIKey(GEMINI_API_KEY);
  GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

  let gptChat = GenAIApp.newChat();
  gptChat.setMaximumAPICalls(2);
  gptChat.addMessage("Give me a step by step plan to become an Apps Script expert.");
  let gptAnswer = gptChat.run({ model: GPT_MODEL });
  console.log(`Max API calls GPT:\n${gptAnswer}`);

  let reasonChat = GenAIApp.newChat();
  reasonChat.setMaximumAPICalls(2);
  reasonChat.addMessage("Give me a step by step plan to become an Apps Script expert.");
  let reasonAnswer = reasonChat.run({ model: REASONING_MODEL });
  console.log(`Max API calls reasonning:\n${reasonAnswer}`);

  let geminiChat = GenAIApp.newChat();
  geminiChat.setMaximumAPICalls(2);
  geminiChat.addMessage("Give me a step by step plan to become an Apps Script expert.");
  let geminiAnswer = geminiChat.run({ model: GEMINI_MODEL });
  console.log(`Max API calls gemini:\n${geminiAnswer}`);
}

function getWeather(cityName) {
  return `The weather in ${cityName} is 19°C today.`;
}
