# GenAIApp

The **GenAIApp** library is a Google Apps Script library designed for creating, managing, and interacting with LLMs using Gemini and OpenAI's API. The library provides features like text-based conversation, browsing the web, image analysis, and more, allowing you to build versatile AI chat applications that can integrate with various functionalities and external data sources.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Setting API Keys](#setting-api-keys)
  - [Creating a New Chat](#creating-a-new-chat)
  - [Adding Messages](#adding-messages)
  - [Adding Callable Functions to the Chat](#adding-callable-functions-to-the-chat)
  - [Enable Web Browsing (Optional)](#enable-web-browsing-optional)
  - [Give a Web Page as a Knowledge Base (Optional)](#give-a-web-page-as-a-knowledge-base-optional)
  - [Enable Vision (Optional)](#enable-vision-optional)
  - [Running the Chat](#running-the-chat)
- [FunctionObject Class](#functionobject-class)
  - [Creating a Function](#creating-a-function)
  - [Configuring Parameters](#configuring-parameters)
- [Advanced Options](#advanced-options)
  - [Retrieving Knowledge from an OpenAI Assistant](#retrieving-knowledge-from-an-openai-assistant)
  - [Analyzing Documents with an OpenAI Assistant](#analyzing-documents-with-an-openai-assistant)
- [Examples](#examples)
  - [Example 1: Send a Prompt and Get Completion](#example-1--send-a-prompt-and-get-completion)
  - [Example 2: Ask Open AI to Create a Draft Reply for the Last Email in Gmail Inbox](#example-2--ask-open-ai-to-create-a-draft-reply-for-the-last-email-in-gmail-inbox)
  - [Example 3: Retrieve Structured Data Instead of Raw Text with onlyReturnArguments](#example-3--retrieve-structured-data-instead-of-raw-text-with-onlyreturnargument)
  - [Example 4: Use Web Browsing](#example-4--use-web-browsing)
  - [Example 5: Describe an Image](#example-5--describe-an-image)
  - [Example 6: Access Google Sheet Content](#example-6--access-google-sheet-content)
- [Debugging and Logging](#debugging-and-logging)
- [Contributing](#contributing)
- [License](#license)
- [Reference](#reference)
  - [Function Object](#function-object)


## Features

- **Chat Creation:** Create interactive chats that can send and receive messages using Gemini or OpenAI's API.
- **Web Search Integration:** Perform web searches using the Google Custom Search API to enhance chatbot responses.
- **Image Analysis:** Retrieve image descriptions using Gemini and OpenAI's vision models.
- **Function Calling:** Enable the chat to call predefined functions and utilize their results in conversations.
- **Assistant Knowledge Retrieval:** Retrieve knowledge from OpenAI vector search assistants for a better contextual response.
- **Document Analysis:** Analyze documents from Google Drive with support for various formats.

## Prerequisites

The setup for **GenAIApp** varies depending on which models you plan to use: 
1. If you want to use **OpenAI models**: You'll need an **OpenAI API key**
2. If you want to use **Google Gemini models**: you’ll need a **Google Cloud Platform (GCP) project** with **Vertex AI** enabled for access to Gemini models.
Ensure to link your Google Apps Script project to a GCP project with Vertex AI enabled, and to include the following scopes in your manifest file:
```js
"oauthScopes": [
    "https://www.googleapis.com/auth/cloud-platform",
    "https://www.googleapis.com/auth/script.external_request"
  ]
```

1. An **OpenAI API key** for accessing OpenAI models.
2. A **Gemini API key** OR a **Google Cloud Platform (GCP) project** for using Gemini models.
3. (Optionnal) A **Google Custom Search API key** for utilizing the Google Custom Search API (if web browsing is enabled).

## Installation

To start using the library, include the **GenAIApp** code in your Google Apps Script project environment. 

## Usage

### Setting API Keys

You need to set your API keys before starting any chat:

```js
// Set Gemini API Key
GenAIApp.setGeminiAPIKey('your-gemini-api-key');

// Set Gemini Auth if using Google Cloud
GenAIApp.setGeminiAuth('your-gcp-project-id','your-region');

// Set OpenAI API Key if using OpenAI
GenAIApp.setOpenAIAPIKey('your-openai-api-key');

// Set Google Search API Key (optional, for web browsing)
GenAIApp.setGoogleSearchAPIKey('your-google-search-api-key');
```

### Creating a New Chat

To start a new chat, call the `newChat()` method. This creates a new Chat instance.

```js
let chat = GenAIApp.newChat();
```

### Adding Messages

You can add messages to your chat using the `addMessage()` method. Messages can be from the user or the system.

```js
// Add a user message
chat.addMessage("Hello, how are you?");

// Add a system message (optional)
chat.addMessage("Answer to the user in a professional way.", true);
```

### Adding callable Functions to the Chat

You can create and add functions to the chat that the AI can call during the conversation:
The `newFunction()` method allows you to create a new Function instance. You can then add this function to your chat using the `addFunction()` method.

```js
// Create a new function
let myFunction = GenAIApp.newFunction()
  .setName("getWeather")
  .setDescription("Retrieve the current weather for a given city.")
  .addParameter("city", "string", "The name of the city.");

// Add the function to the chat
chat.addFunction(myFunction);
```

From the moment that you add a function to chat, we will use function calling features.
For more information : 
- [https://ai.google.dev/gemini-api/docs/function-calling](https://ai.google.dev/gemini-api/docs/function-calling)
- [https://platform.openai.com/docs/guides/gpt/function-calling](https://platform.openai.com/docs/guides/gpt/function-calling)

### Enable web browsing (optional)

If you want to allow the chat to perform web searches and fetch web pages, you can either:
- run the chat with a Gemini model
- if you're using OpenAI's model, you'll need to enable browsing (and set your Google Custom Search API key, as desbribed above)

```javascript
chat.enableBrowsing(true);
```
If want to restrict your browsing to a specific web page, you can add as a second argument the url of this web page as bellow.

```javascript
  chat.enableBrowsing(true, "https://support.google.com");
```
### Give a web page as a knowledge base (optional)

If you don't need the perform a web search and want to directly give a link for a web page you want the chat to read before performing any action, you can use the addKnowledgeLink(url) function.

```javascript
  chat.addKnowledgeLink("https://developers.google.com/apps-script/guides/libraries");
```
### Enable Vision (optional)

To enable the chat model to describe images, use the `enableVision()` method

```javascript
chat.enableVision(true);
```

   - [Enable Vision (Optional)](#enable-vision-optional)
   - [Add File to Gemini (optional)](#add-file-to-gemini-optional)

```javascript
// Add a Google Drive file to the Gemini chat context using its Drive file ID
chat.addFile('your-google-drive-file-id');
```

### Running the Chat

Once you've set up the chat and added the necessary components, you can start the conversation by calling the `run()` method.

```js
let response = chat.run({
  model: "gemini-1.5-pro-002", // Optional: set the model to use
  temperature: 0.5 // Optional: set response creativity
  function_call: "getWeather" // Optional: force the first API response to call a function
});

console.log(response);
```
The library supports the following models: 
1. Gemini: "gemini-1.5-pro-002" | "gemini-1.5-pro" | "gemini-1.5-flash-002" | "gemini-1.5-flash"
2. OpenAI: "gpt-3.5-turbo" | "gpt-4" | "gpt-4-turbo" | "gpt-4o" | "gpt-4o-mini"

⚠️ **Warning:** the "function_call" advanced parameter is only supported by OpenAI models, gemini-1.5-pro and gemini-1.5-flash

## FunctionObject Class

### Creating a Function

The **FunctionObject** class represents a callable function by the chatbot. It is highly customizable with various options:

```js
let functionObject = GenAIApp.newFunction()
  .setName("searchMovies")
  .setDescription("Search for movies based on a genre.")
  .addParameter("genre", "string", "The genre of movies to search for.");
```

### Configuring Parameters

The function parameters can be configured to be required or optional:

```js
// Adding required parameter
functionObject.addParameter("year", "number", "The year of the movie release.");

// Adding optional parameter
functionObject.addParameter("rating", "number", "The minimum rating of movies to return.", true);
```

## Advanced Options

### Retrieving Knowledge from an OpenAI Assistant

Retrieve contextual information from a specific OpenAI vector search assistant:

```js
chat.retrieveKnowledgeFromAssistant("assistant-id", "A description of available knowledge.");
```
To find out more : [https://platform.openai.com/docs/assistants/overview](https://platform.openai.com/docs/assistants/overview)

### Analyzing Documents with an OpenAI Assistant

Analyze a document from Google Drive using an assistant:

```js
chat.analyzeDocumentWithAssistant("assistant-id", "drive-file-id");
```

## Examples

### Example 1 : Send a prompt and get completion

```javascript
 ChatGPTApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

 const chat = ChatGPTApp.newChat();
 chat.addMessage("What are the steps to add an external library to my Google Apps Script project?");

 const chatAnswer = chat.run();
 Logger.log(chatAnswer);
```

### Example 2 : Ask Open AI to create a draft reply for the last email in Gmail inbox

```javascript
 ChatGPTApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
 const chat = ChatGPTApp.newChat();

 var getLatestThreadFunction = ChatGPTApp.newFunction()
    .setName("getLatestThread")
    .setDescription("Retrieve information from the last message received.");

 var createDraftResponseFunction = ChatGPTApp.newFunction()
    .setName("createDraftResponse")
    .setDescription("Create a draft response.")
    .addParameter("threadId", "string", "the ID of the thread to retrieve")
    .addParameter("body", "string", "the body of the email in plain text");

  var resp = ChatGPTApp.newChat()
    .addMessage("You are an assistant managing my Gmail inbox.", true)
    .addMessage("Retrieve the latest message I received and draft a response.")
    .addFunction(getLatestThreadFunction)
    .addFunction(createDraftResponseFunction)
    .run();

  console.log(resp);
```

### Example 3 : Retrieve structured data instead of raw text with onlyReturnArgument()

```javascript
const ticket = "Hello, could you check the status of my subscription under customer@example.com";

  chat.addMessage("You just received this ticket : " + ticket);
  chat.addMessage("What's the customer email address ? You will give it to me using the function getEmailAddress.");

  const myFunction = ChatGPTApp.newFunction() // in this example, getEmailAddress is not actually a real function in your script
    .setName("getEmailAddress")
    .setDescription("To give the user an email address")
    .addParameter("emailAddress", "string", "the email address")
    .onlyReturnArguments(true) // you will get your parameters in a json object

  chat.addFunction(myFunction);

  const chatAnswer = chat.run();
  Logger.log(chatAnswer["emailAddress"]); // the name of the parameter of your "fake" function

  // output : 	"customer@example.com"
```

### Example 4 : Use web browsing

```javascript
 const message = "You're a google support agent, a customer is asking you how to install a library he found on github in a google appscript project."

 const chat = ChatGPTApp.newChat();
 chat.addMessage(message);
 chat.addMessage("Browse this website to answer : https://developers.google.com/apps-script", true)
 chat.enableBrowsing(true);

 const chatAnswer = chat.run();
 Logger.log(chatAnswer);
```

### Example 5 : Describe an Image

To have the chat model describe an image: 

```javascript
const chat = ChatGPTApp.newChat();
chat.enableVision(true);
chat.addMessage("Describe the following image.");
chat.addImage("https://example.com/image.jpg", "high");
const response = chat.run();
Logger.log(response);
```
This will enable the vision capability and use the OpenAI model to provide a description of the image at the specified URL. The fidelity parameter can be "low" or "high", affecting the detail level of the description.

### Example 6 : Access Google Sheet Content

To retrieve data from a Google Sheet:

```javascript
const chat = ChatGPTApp.newChat();
chat.enableGoogleSheetsAccess(true);
chat.addMessage("What data is stored in the following spreadsheet?");
const spreadsheetId = "your_spreadsheet_id_here";
chat.run({
  function_call: "getDataFromGoogleSheets",
  arguments: { spreadsheetId: spreadsheetId }
});
const response = chat.run();
Logger.log(response);
```
This example demonstrates how to enable access to Google Sheets and retrieve data from a specified spreadsheet.

## Debugging and Logging

To debug the chat and view the search queries and pages opened:

```js
let debugInfo = GenAIApp.debug(chat);

// Get web search queries
let webSearchQueries = debugInfo.getWebSearchQueries();

// Get web pages opened
let webPagesOpened = debugInfo.getWebPagesOpened();

console.log(webSearchQueries, webPagesOpened);
```

## Contributing

Contributions are welcome! If you find any bugs, have feature requests, or want to contribute code, please submit an issue or pull request on this [GitHub repository](https://github.com/scriptit-fr/GenAIApp).

## License

The **GenAIApp** library is licensed under the Apache License, Version 2.0. You may not use this file except in compliance with the License. For more details, please see the [LICENSE](http://www.apache.org/licenses/LICENSE-2.0).

## Reference

### Function Object

A `FunctionObject` represents a function that can be called by the chat.

Creating a function object and setting its name to the name of an actual function you have in your script will permit the library to call your real function.

#### `setName(name)`

Sets the name of the function.

#### `setDescription(description)`

Sets the description of the function.

#### `addParameter(name, type, description, [isOptional])`

Adds a parameter to the function. Parameters are required by default. Set 'isOptional' to true to make a parameter optional.

#### `endWithResult(bool)`

If enabled, the conversation with the chat will automatically end after this function is executed.

#### `onlyReturnArguments(bool)`

If enabled, the conversation will automatically end when this function is called and the chat will return the arguments in a stringified JSON object.

#### `toJSON()`

Returns a JSON representation of the function object.

### Chat

A `Chat` represents a conversation with the chat.

#### `addMessage(messageContent, [system])`

Add a message to the chat. If `system` is true, the message is from the system, else it's from the user.

#### `addFunction(functionObject)`

Add a function to the chat.

#### `enableBrowsing(bool)`

Enable the chat to use a Google search engine to browse the web.

#### `run([advancedParametersObject])`

Start the chat conversation. It sends all your messages and any added function to the chat GPT. It will return the last chat answer.

Supported attributes for the advanced parameters :

```javascript
advancedParametersObject = {
	temperature: temperature, 
	model: model,
	function_call: function_call
}
```

**Temperature** : Lower values for temperature result in more consistent outputs, while higher values generate more diverse and creative results. Select a temperature value based on the desired trade-off between coherence and creativity for your specific application.

**Model** : The Gemini and OpenAI API are powered by a diverse set of models with different capabilities and price points. 
- [Find out more about Gemini models](https://ai.google.dev/gemini-api/docs/models/gemini)
- [Find out more about OpenAI models](https://platform.openai.com/docs/models/overview)

**Function_call** : If you want to force the model to call a specific function you can do so by setting `function_call: "<insert-function-name>"`.

### Note

If you wish to disable the library logs and keep only your own, call `disableLogs()`:

```javascript
ChatGPTApp.disableLogs();
```

This can be useful for keeping your logs clean and specific to your application.

---

Happy coding and enjoy building with the **GenAIApp** library!
