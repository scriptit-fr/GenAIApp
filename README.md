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
  - [Add Image (Optional)](#add-image-optional)
  - [Add File to Chat (optional)](#add-file-to-chat-optional)
  - [Add MCP Connector (optional)](#add-mcp-connector-optional)
  - [Running the Chat](#running-the-chat)
  - [FunctionObject Class](#functionobject-class)
    - [Creating a Function](#creating-a-function)
    - [Configuring Parameters](#configuring-parameters)
  - [VectorStoreObject Class](#vectorstoreobject-class)
  - [Retrieving Knowledge from an OpenAI Vector Store](#retrieving-knowledge-from-an-openai-vector-store)
- [Examples](#examples)
  - [Example 1: Send a Prompt and Get Completion](#example-1--send-a-prompt-and-get-completion)
  - [Example 2: Ask Open AI to Create a Draft Reply for the Last Email in Gmail Inbox](#example-2--ask-open-ai-to-create-a-draft-reply-for-the-last-email-in-gmail-inbox)
  - [Example 3: Retrieve Structured Data Instead of Raw Text with onlyReturnArguments](#example-3--retrieve-structured-data-instead-of-raw-text-with-onlyreturnargument)
  - [Example 4: Use Web Browsing](#example-4--use-web-browsing)
  - [Example 5: Describe an Image](#example-5--describe-an-image)
  - [Example 6: Extend a Chat with an MCP Connector](#example-6--extend-a-chat-with-an-mcp-connector)
- [Contributing](#contributing)
- [License](#license)
- [Reference](#reference)
  - [GenAIApp](#genaiapp)
  - [Chat](#chat)
  - [Function Object](#function-object)
  - [Vector Store Object](#vector-store-object)
  - [Connector Object](#connector-object)


## Features

- **Chat Creation:** Create interactive chats that can send and receive messages using Gemini or OpenAI's API.
- **Web Search Integration:** Perform web searches to enhance chatbot responses.
- **Image Analysis:** Retrieve image descriptions using Gemini and OpenAI's vision models.
- **Function Calling:** Enable the chat to call predefined functions and utilize their results in conversations.
- **Vector Store Search:** Retrieve knowledge from OpenAI vector stores for a better contextual response.
- **Document Analysis:** Analyze documents from Google Drive with support for various formats.
- **MCP Connectors:** Attach Google Workspace or custom Model Context Protocol connectors to securely retrieve additional context
  during a conversation.

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

// Set global metadata passed with each request (optional)
GenAIApp.setGlobalMetadata('app', 'demo');

// Use a custom OpenAI-compatible endpoint (optional)
GenAIApp.setPrivateInstanceBaseUrl('https://your-endpoint.example.com');
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

If you want to allow the chat to perform web searches and fetch web pages, enable browsing on your chat instance:

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
### Add Image (optional)

To include an image in the conversation, use the `addImage()` method with a URL or a Blob.

```javascript
chat.addImage('https://example.com/image.jpg');
```

### Add File to Chat (optional)

You can include the contents of a Google Drive file or a Blob in your conversation using the `addFile()` method. This works with both Gemini and OpenAI multimodal models.

```javascript
// Add a Google Drive file to the chat context using its Drive file ID
chat.addFile('your-google-drive-file-id');
```

### Add MCP Connector (optional)

Use Model Context Protocol (MCP) connectors to let OpenAI Responses API models reach structured data sources such as Gmail,
Calendar, Drive, or your own custom MCP server.

```javascript
const chat = GenAIApp.newChat();

const gmailConnector = GenAIApp.newConnector()
  .setConnectorId('gmail')
  .setRequireApproval('domain');

chat.addMCP(gmailConnector);

const customConnector = GenAIApp.newConnector()
  .setLabel('Salesforce CRM')
  .setDescription('Query opportunity data from Salesforce via MCP proxy')
  .setServerUrl('https://mcp.example.com/salesforce')
  .setAuthorization('Bearer ' + SALESFORCE_MCP_TOKEN)
  .setRequireApproval('always');

chat.addMCP(customConnector);
```

- **Google Workspace connectors:** Call `.setConnectorId("gmail" | "calendar" | "drive")` to use Google-managed connectors
  authenticated with your script's OAuth token by default.
- **Custom MCP servers:** Configure a connector with `.setLabel()`, `.setDescription()`, and `.setServerUrl("https://...")`,
  and optionally `.setAuthorization()` if the server expects a bearer token or API key.
- **Approval workflows:** `.setRequireApproval('never' | 'domain' | 'always')` lets you enforce end-user approval before the
  model calls the connector.

> ⚠️ MCP connectors are currently available only when you run the chat with OpenAI Responses API models (for example, `gpt-4.1`,
> `o4-mini`, `o3`, or `gpt-5`).

### Running the Chat

Once you've set up the chat and added the necessary components, you can start the conversation by calling the `run()` method.

```js
let response = chat.run({
  model: "gemini-2.5-flash", // Optional: set the model to use
  temperature: 0.5, // Optional: set response creativity
  function_call: "getWeather" // Optional: force the first API response to call a function
});

console.log(response);
```
The library supports the following models: 
1. Gemini: "gemini-2.5-pro" | "gemini-2.5-flash"
2. OpenAI: "gpt-4.1" | "o4-mini" | "o3" | "gpt-5"

⚠️ **Warning:** the "function_call" advanced parameter is supported by:
  - OpenAI models (including GPT-5)  
  - Gemini 2.5 variants (gemini-2.5-pro, gemini-2.5-flash, gemini-2.5-flash-lite, gemini-2.5-flash-native-audio)

  The "reasoning_effort" parameter is supported only by reasoning-capable OpenAI models and ignored by all others.
⚠️ **Warning:** The "reasoning_effort" parameter is supported only by reasoning-capable OpenAI models and ignored by all others.

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

## VectorStoreObject Class

### Retrieving Knowledge from an OpenAI Vector Store

Retrieve contextual information from a specific OpenAI vector search :

```js
const vectorStoreObject = GenAIApp.newVectorStore()
  .initializeFromId("your-vector-store-id");
chat.addVectorStore(vectorStoreObject);
```
To find out more : [https://platform.openai.com/docs/api-reference/vector_stores/search](https://platform.openai.com/docs/api-reference/vector_stores/search)

## Examples

### Example 1 : Send a prompt and get completion

```javascript
 GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

 const chat = GenAIApp.newChat();
 chat.addMessage("What are the steps to add an external library to my Google Apps Script project?");

 const chatAnswer = chat.run();
 Logger.log(chatAnswer);
```

### Example 2 : Ask Open AI to create a draft reply for the last email in Gmail inbox

```javascript
 GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
 const chat = GenAIApp.newChat();

 var getLatestThreadFunction = GenAIApp.newFunction()
    .setName("getLatestThread")
    .setDescription("Retrieve information from the last message received.");

 var createDraftResponseFunction = GenAIApp.newFunction()
    .setName("createDraftResponse")
    .setDescription("Create a draft response.")
    .addParameter("threadId", "string", "the ID of the thread to retrieve")
    .addParameter("body", "string", "the body of the email in plain text");

  var resp = GenAIApp.newChat()
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

  const myFunction = GenAIApp.newFunction() // in this example, getEmailAddress is not actually a real function in your script
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

 const chat = GenAIApp.newChat();
 chat.addMessage(message);
 chat.addMessage("Browse this website to answer : https://developers.google.com/apps-script", true)
 chat.enableBrowsing(true);

 const chatAnswer = chat.run();
 Logger.log(chatAnswer);
```

### Example 5 : Describe an Image

To have the chat model describe an image:

```javascript
const chat = GenAIApp.newChat();
chat.addMessage("Describe the following image.");
chat.addImage("https://example.com/image.jpg");
const response = chat.run();
Logger.log(response);
```
This will use the selected model to provide a description of the image at the specified URL.

### Example 6 : Extend a Chat with an MCP Connector

```javascript
GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

const chat = GenAIApp.newChat();
chat.addMessage('Search my latest unread Gmail message and summarize it.');

const gmailConnector = GenAIApp.newConnector()
  .setConnectorId('gmail')
  .setRequireApproval('domain');

chat.addMCP(gmailConnector);

const summary = chat.run({ model: 'gpt-4.1' });
Logger.log(summary);
```

In this example the Gmail connector gives the model controlled access to your inbox. The `requireApproval('domain')` setting
ensures end users in your domain must approve access before the connector is used.

### Example 7 : Connect to a Custom MCP Server with setServerUrl()

```javascript
GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

const chat = GenAIApp.newChat();
chat.addMessage('Check the latest closed-won opportunities and report total revenue.');

const salesforceConnector = GenAIApp.newConnector()
  .setLabel('Salesforce CRM')
  .setDescription('Internal MCP service that proxies Salesforce data')
  .setServerUrl('https://mcp.example.com/salesforce')
  .setAuthorization('Bearer ' + SALESFORCE_MCP_TOKEN)
  .setRequireApproval('always');

chat.addMCP(salesforceConnector);

const report = chat.run({ model: 'gpt-4.1' });
Logger.log(report);
```

The `setServerUrl()` method points the connector to your MCP gateway, while `setAuthorization()` injects a bearer token or API
key that the proxy expects. Combine these settings with `.setRequireApproval('always')` if you want end users to explicitly
authorize every connector invocation.

## Contributing

Contributions are welcome! If you find any bugs, have feature requests, or want to contribute code, please submit an issue or pull request on this [GitHub repository](https://github.com/scriptit-fr/GenAIApp).

## License

The **GenAIApp** library is licensed under the Apache License, Version 2.0. You may not use this file except in compliance with the License. For more details, please see the [LICENSE](http://www.apache.org/licenses/LICENSE-2.0).

## Reference

### GenAIApp

- `newChat()`: Create a new `Chat` instance.
- `newFunction()`: Create a new `FunctionObject`.
- `newConnector()`: Create a new `ConnectorObject` for MCP integrations.
- `newVectorStore()`: Create a new `VectorStoreObject`.
- `setOpenAIAPIKey(apiKey)`: Set the OpenAI API key.
- `setGeminiAPIKey(apiKey)`: Set the Gemini API key.
- `setGeminiAuth(projectId, region)`: Use Vertex AI authentication.
- `setGlobalMetadata(key, value)`: Attach a key/value pair to every request.
- `setPrivateInstanceBaseUrl(baseUrl)`: Use a custom OpenAI‑compatible endpoint.

### Chat

A `Chat` represents a conversation with the model.

- `addMessage(messageContent, [system])`: Add a user or system message.
- `addFunction(functionObject)`: Attach a `FunctionObject` for function calling.
- `addImage(imageInput)`: Include an image URL or Blob in the conversation.
- `addFile(fileInput)`: Include the content of a Google Drive file or Blob.
- `addMetadata(key, value)`: Add metadata sent with the next request.
- `getAttributes()`: Retrieve attributes from vector store search results.
- `onlyReturnChunks(bool)`: Return raw chunks from vector store searches.
- `setMaxChunks(maxChunks)`: Limit the number of chunks returned by vector stores.
- `getMessages()`: Get the messages as a JSON string.
- `getFunctions()`: Get the functions as a JSON string.
- `disableLogs(bool)`: Disable library logs.
- `enableBrowsing(bool, [url])`: Allow the model to browse the web, optionally restricted to a URL.
- `addKnowledgeLink(url)`: Inject the content of a web page into the conversation.
- `addMCP(connectorObject)`: Attach one or more MCP connectors to the chat request.
- `setMaximumAPICalls(maxAPICalls)`: Limit the number of API calls in a run.
- `retrieveLastResponseId()`: Get the last response ID.
- `setPreviousResponseId(id)`: Provide the previous response ID to continue a conversation.
- `addVectorStores(vectorStoreIds)`: Attach vector store IDs for retrieval.
- `run([advancedParametersObject])`: Execute the chat and return the response. Supports `model`, `temperature`, `reasoning_effort`, `max_tokens`, and `function_call` parameters.

### Function Object

A `FunctionObject` represents a function that can be called by the chat.

- `setName(name)`: Set the function name.
- `setDescription(description)`: Set the function description.
- `addParameter(name, type, description, [isOptional])`: Add a parameter to the function. Parameters are required by default; set `isOptional` to `true` to make a parameter optional.
- `endWithResult(bool)`: End the conversation after the function is executed.
- `onlyReturnArguments(bool)`: End the conversation and return only the arguments.

### Vector Store Object

A `VectorStoreObject` represents an OpenAI vector store.

- `setName(newName)`: Set the vector store name.
- `setDescription(newDesc)`: Set the description.
- `setChunkingStrategy(maxChunkSize, chunkOverlap)`: Configure chunking before uploads.
- `createVectorStore()`: Create the vector store.
- `initializeFromId(vectorStoreId)`: Initialize from an existing vector store ID.
- `getId()`: Get the vector store ID.
- `uploadAndAttachFile(blob, attributes)`: Upload a file and attach it to the store.
- `listFiles()`: List files attached to the store.
- `deleteFile(fileId)`: Delete a file from the store.
- `deleteVectorStore()`: Delete the vector store.

### Connector Object

A `ConnectorObject` represents a Google Workspace or custom MCP connector that can be attached to an OpenAI chat request.

- `setLabel(label)`: Set the identifier used in the chat payload (required for custom servers).
- `setDescription(description)`: Provide an optional description visible to the model.
- `setServerUrl(url)`: Use a custom MCP server hosted at the provided HTTPS URL.
- `setConnectorId('gmail'|'calendar'|'drive')`: Reference a Google Workspace MCP connector by its predefined ID.
- `setAuthorization(token)`: Override the default OAuth token (for example, supply `Bearer ...`).
- `setRequireApproval('never'|'domain'|'always')`: Control whether the connector requires user approval before execution.

---

Happy coding and enjoy building with the **GenAIApp** library!
