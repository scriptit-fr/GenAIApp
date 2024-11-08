# GenAIApp

The **GenAIApp** library is a JavaScript library designed for creating, managing, and interacting with AI-powered chatbots using Gemini and OpenAI's API. The library provides features like text-based conversation, browsing the web, image analysis, and more, allowing you to build versatile AI chat applications that can integrate with various functionalities and external data sources.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Setting API Keys](#setting-api-keys)
  - [Creating a New Chat](#creating-a-new-chat)
  - [Adding Messages](#adding-messages)
  - [Adding Functions to the Chat](#adding-functions-to-the-chat)
  - [Running the Chat](#running-the-chat)
- [FunctionObject Class](#functionobject-class)
  - [Creating a Function](#creating-a-function)
  - [Configuring Parameters](#configuring-parameters)
- [Advanced Options](#advanced-options)
  - [Enabling Web Browsing](#enabling-web-browsing)
  - [Enabling Vision Analysis](#enabling-vision-analysis)
  - [Retrieving Knowledge from an OpenAI Assistant](#retrieving-knowledge-from-an-openai-assistant)
  - [Analyzing Documents with an OpenAI Assistant](#analyzing-documents-with-an-openai-assistant)
- [Debugging and Logging](#debugging-and-logging)
- [Contributing](#contributing)
- [License](#license)

## Features

- **Chatbot Creation:** Create interactive chatbots that can send and receive messages using Gemini or OpenAI's API.
- **Web Search Integration:** Perform web searches using the Google Custom Search API to enhance chatbot responses.
- **Image Analysis:** Retrieve image descriptions using Gemini and OpenAI's vision models.
- **Function Execution:** Enable the chat to call predefined functions and utilize their results in conversations.
- **Assistant Knowledge Retrieval:** Retrieve knowledge from specific assistants for a better contextual response.
- **Document Analysis:** Analyze documents from Google Drive with support for various formats.

## Prerequisites

To use the **GenAIApp** library, you will need:

1. An **OpenAI API key** for accessing OpenAI models.
2. A **Google Custom Search API key** for utilizing the Google Custom Search API (if web browsing is enabled).
3. A **Google Cloud Platform (GCP) project** for using Gemini models and document analysis features.

## Installation

To start using the library, include the **GenAIApp** code in your Google Apps Script project environment. 

## Usage

### Setting API Keys

You need to set your API keys before starting any chat:

```js
// Set OpenAI API Key
GenAIApp.setOpenAIAPIKey('your-openai-api-key');

// Set Gemini API Key
GenAIApp.setGeminiAPIKey('your-gemini-api-key');

// Set Gemini Auth if using Google Cloud
GenAIApp.setGeminiAuth({
  project_id: 'your-gcp-project-id',
  region: 'your-region'
});

// Set Google Search API Key (optional, for web browsing)
GenAIApp.setGoogleSearchAPIKey('your-google-search-api-key');
```

### Creating a New Chat

To create a new chat instance:

```js
let chat = GenAIApp.newChat();
```

### Adding Messages

Add messages to the chat either from the user or the system:

```js
// Add a user message
chat.addMessage("Hello, how can you assist me?");

// Add a system message (optional)
chat.addMessage("You are speaking with an AI assistant.", true);
```

### Adding Functions to the Chat

You can create and add functions to the chat that the AI can call during the conversation:

```js
// Create a new function
let myFunction = GenAIApp.newFunction()
  .setName("getWeather")
  .setDescription("Retrieve the current weather for a given city.")
  .addParameter("city", "string", "The name of the city.");

// Add the function to the chat
chat.addFunction(myFunction);
```

### Running the Chat

Once you've set up the chat and added the necessary components, run the conversation:

```js
let response = chat.run({
  model: "gemini", // Optional: set the model to use
  temperature: 0.5 // Optional: set response creativity
});

console.log(response);
```

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

### Enabling Web Browsing

You can allow the chatbot to perform web searches using the Google Custom Search API:

```js
chat.enableBrowsing(true, "search-engine-id");
```

### Enabling Vision Analysis

Enable the vision analysis feature for the chat:

```js
chat.enableVision(true);
```

### Retrieving Knowledge from an OpenAI Assistant

Retrieve contextual information from a specific assistant:

```js
chat.retrieveKnowledgeFromAssistant("assistant-id", "A description of available knowledge.");
```

### Analyzing Documents with an OpenAI Assistant

Analyze a document from Google Drive using an assistant:

```js
chat.analyzeDocumentWithAssistant("assistant-id", "drive-file-id");
```

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

---

Happy coding and enjoy building with the **GenAIApp** library!
