# GenAIApp

The **GenAIApp** library is a Google Apps Script library designed for creating, managing, and interacting with LLMs using Gemini and OpenAI APIs. It supports text conversations, web browsing, image and document analysis, function calling, vector-store retrieval, MCP connectors, and Google Workspace integrations.

## Table of Contents

- [Features](#features)
- [Samples](#samples)
  - [Quick Start](#quick-start)
  - [Getting Started](#getting-started)
  - [Content Analysis](#content-analysis)
  - [Function Calling](#function-calling)
  - [Advanced Features](#advanced-features)
  - [Integration](#integration)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [API Guide](#api-guide)
  - [Setting API Keys](#setting-api-keys)
  - [Creating a New Chat](#creating-a-new-chat)
  - [Adding Messages](#adding-messages)
  - [Adding Callable Functions to the Chat](#adding-callable-functions-to-the-chat)
  - [Enable Web Browsing Optional](#enable-web-browsing-optional)
  - [Enable OpenAI Server-Side Compaction Optional](#enable-openai-server-side-compaction-optional)
  - [Give a Web Page as a Knowledge Base Optional](#give-a-web-page-as-a-knowledge-base-optional)
  - [Add Image Optional](#add-image-optional)
  - [Add File to Chat Optional](#add-file-to-chat-optional)
  - [Add an MCP Connector Optional](#add-an-mcp-connector-optional)
  - [Running the Chat](#running-the-chat)
  - [FunctionObject Class](#functionobject-class)
  - [VectorStoreObject Class](#vectorstoreobject-class)
- [Reference](#reference)
  - [GenAIApp Factory](#genaiapp-factory)
  - [Chat](#chat)
  - [Function Object](#function-object)
  - [Vector Store Object](#vector-store-object)
  - [Connector Object](#connector-object)
- [Contributing](#contributing)
- [License](#license)

## Features

- **Chat Creation:** Create interactive chats that can send and receive messages using Gemini or OpenAI APIs.
- **Web Search Integration:** Perform web searches to enhance chatbot responses.
- **Image Analysis:** Retrieve image descriptions using Gemini and OpenAI vision models.
- **Function Calling:** Enable the chat to call predefined functions and use their results in conversations.
- **Vector Store Search:** Retrieve knowledge from OpenAI vector stores or Google Gemini File Search Stores for better contextual responses.
- **Document Analysis:** Analyze documents from Google Drive with support for multiple formats.
- **MCP Connectors:** Attach Google Workspace or custom Model Context Protocol connectors to securely retrieve additional context during a conversation.

## Samples

> **Start here:** The `samples/` directory contains copyable Google Apps Script examples organized by use case. Use this section as the fastest way to find a working pattern before reading the detailed API guide.

### Quick Start

Try [`samples/simple-chat.gs`](samples/simple-chat.gs) first. It is the recommended hello-world smoke test: set an OpenAI API key in Script Properties as `OPENAI_API_KEY`, paste or import the sample into Apps Script, and run `simpleChatSample()` to log a short model response.

### Getting Started

| Sample | Description | Demonstrates |
| --- | --- | --- |
| [`simple-chat.gs`](samples/simple-chat.gs) | Smallest possible GenAIApp chat request for a hello-world smoke test. | `setOpenAIAPIKey()`, `newChat()`, `addMessage()`, `run()` |
| [`system-prompts.gs`](samples/system-prompts.gs) | Sets assistant role, tone, and response format with a system message. | System messages with `addMessage(message, true)`, prompt shaping |
| [`configuration-options.gs`](samples/configuration-options.gs) | Shows common configuration and guardrail settings for production scripts. | `disableLogs()`, `warnIfResponseTokenUsageAbove()`, `enableCompaction()`, `setMaximumAPICalls()` |
| [`multi-model-usage.gs`](samples/multi-model-usage.gs) | Reuses the same prompt across OpenAI and Gemini model names for comparison. | `setOpenAIAPIKey()`, `setGeminiAPIKey()`, model selection in `run()` |
| [`vertex-ai-setup.gs`](samples/vertex-ai-setup.gs) | Authenticates Gemini through a linked Google Cloud project instead of an API key. | `setGeminiAuth()`, Vertex AI configuration, Gemini model execution |

### Content Analysis

| Sample | Description | Demonstrates |
| --- | --- | --- |
| [`image-analysis.gs`](samples/image-analysis.gs) | Sends both a public image URL and an Apps Script Blob to a vision-capable model. | `addImage()`, multimodal prompts, Blob inputs |
| [`document-analysis.gs`](samples/document-analysis.gs) | Summarizes PDFs or exported Google Workspace files from Drive and Blob inputs. | `addFile()`, Drive file IDs, Blob file analysis |
| [`knowledge-links.gs`](samples/knowledge-links.gs) | Injects a known web page as direct context without broad web search. | `addKnowledgeLink()`, page-grounded answers |
| [`web-browsing.gs`](samples/web-browsing.gs) | Allows real-time browsing with an optional trusted-domain restriction. | `enableBrowsing(true, url)`, current-information prompts |
| [`vector-store-rag.gs`](samples/vector-store-rag.gs) | Creates a full OpenAI vector store, uploads source content, queries it, and returns chunks. | `newVectorStore()`, `uploadAndAttachFile()`, `addVectorStores()`, `onlyReturnChunks()` |
| [`openai-vector-store-quickstart.gs`](samples/openai-vector-store-quickstart.gs) | Minimal OpenAI vector-store retrieval example. | `newVectorStore('openai')`, `uploadAndAttachFile()`, OpenAI file search |
| [`google-file-search-store-quickstart.gs`](samples/google-file-search-store-quickstart.gs) | Minimal Gemini File Search Store retrieval example. | `newGeminiFileSearchStore()`, `createFileSearchStore()`, `uploadAndImportDocument()` |

### Function Calling

| Sample | Description | Demonstrates |
| --- | --- | --- |
| [`function-calling-basics.gs`](samples/function-calling-basics.gs) | Registers a single callable tool and lets the model use Apps Script code. | `newFunction()`, `setName()`, `setDescription()`, `addParameter()`, `addFunction()` |
| [`function-calling-advanced.gs`](samples/function-calling-advanced.gs) | Extracts structured arguments or ends early after a tool result. | `onlyReturnArguments(true)`, `endWithResult(true)`, tool-routing patterns |

### Advanced Features

| Sample | Description | Demonstrates |
| --- | --- | --- |
| [`conversation-continuation.gs`](samples/conversation-continuation.gs) | Continues OpenAI Responses API and Gemini Interactions API conversations without resending the full transcript. | `retrieveLastResponseId()`, `setPreviousResponseId()`, `retrieveLastInteractionId()`, `setPreviousInteractionId()` |
| [`configuration-options.gs`](samples/configuration-options.gs) | Configures operational controls for long-running or budget-sensitive automations. | Token warnings, API call limits, compaction thresholds, logging controls |
| [`multi-model-usage.gs`](samples/multi-model-usage.gs) | Compares outputs from OpenAI and Gemini model names with one prompt. | Model IDs, provider switching |
| [`vector-store-rag.gs`](samples/vector-store-rag.gs) | Builds retrieval-augmented generation on OpenAI vector stores. | Vector-store lifecycle, chunking, attributes, retrieval responses |
| [`openai-vector-store-quickstart.gs`](samples/openai-vector-store-quickstart.gs) | Shows the shortest OpenAI vector-store create-upload-query flow. | OpenAI vector-store IDs, file upload, chat retrieval |
| [`google-file-search-store-quickstart.gs`](samples/google-file-search-store-quickstart.gs) | Shows the shortest Gemini File Search Store create-upload-query flow. | Gemini store resource names, direct document upload/import, chat retrieval |

### Integration

| Sample | Description | Demonstrates |
| --- | --- | --- |
| [`sheets-ai-assistant.gs`](samples/sheets-ai-assistant.gs) | Reads active Google Sheets data, asks AI for observations, and writes results back. | `SpreadsheetApp`, sheet-bound workflows, chat summarization |
| [`mcp-connectors.gs`](samples/mcp-connectors.gs) | Configures Gmail, Calendar, and Drive MCP connectors for Workspace-aware responses. | `newConnector()`, `setConnectorId()`, `setAuthorization()`, `addMCP()` |
| [`google-mcp-connector.gs`](samples/google-mcp-connector.gs) | Connects directly to Google's Native Gmail MCP endpoint. | `setServerUrl()`, native Google MCP endpoint setup, OAuth token authorization |

## Prerequisites

Choose the credentials that match the models you plan to use:

1. **OpenAI models:** Store an OpenAI API key for `GenAIApp.setOpenAIAPIKey()`.
2. **Gemini with an API key:** Store a Gemini API key for `GenAIApp.setGeminiAPIKey()`.
3. **Gemini through Vertex AI:** Link your Apps Script project to a Google Cloud project with Vertex AI enabled, then use `GenAIApp.setGeminiAuth(projectId, region)`.

For Vertex AI or Google Workspace MCP connectors, include the required OAuth scopes in your Apps Script manifest. Start with:

```js
"oauthScopes": [
  "https://www.googleapis.com/auth/cloud-platform",
  "https://www.googleapis.com/auth/script.external_request"
]
```

> **Note:** The `oauthScopes` array shown above is a starting point. Google Workspace MCP connectors (Gmail, Calendar, Drive) require additional service-specific scopes. For example, Gmail requires `https://www.googleapis.com/auth/gmail.readonly`, Calendar requires `https://www.googleapis.com/auth/calendar`, and Drive requires `https://www.googleapis.com/auth/drive.readonly`. The exact scopes depend on which connectors and operations you use, so you must add the corresponding scopes to the `oauthScopes` manifest entry.

## Installation

Setup is intentionally lightweight: drag and drop the **GenAIApp** library files into your Google Apps Script project, add the needed credentials in Script Properties, and start with [`samples/simple-chat.gs`](samples/simple-chat.gs).

## API Guide

### Setting API Keys

You need to set your API keys before starting any chat. See [`samples/simple-chat.gs`](samples/simple-chat.gs), [`samples/multi-model-usage.gs`](samples/multi-model-usage.gs), and [`samples/vertex-ai-setup.gs`](samples/vertex-ai-setup.gs) for complete setup examples.

```js
// Set Gemini API Key
GenAIApp.setGeminiAPIKey('your-gemini-api-key');

// Set Gemini Auth if using Google Cloud
GenAIApp.setGeminiAuth('your-gcp-project-id', 'your-region');

// Set OpenAI API Key if using OpenAI
GenAIApp.setOpenAIAPIKey('your-openai-api-key');

// Set global metadata passed with each request (optional)
GenAIApp.setGlobalMetadata('app', 'demo');

// Use a custom OpenAI-compatible endpoint (optional)
GenAIApp.setPrivateInstanceBaseUrl('https://your-endpoint.example.com');
```

### Creating a New Chat

To start a new chat, call the `newChat()` method. This creates a new `Chat` instance. For the smallest runnable version, use [`samples/simple-chat.gs`](samples/simple-chat.gs).

```js
const chat = GenAIApp.newChat();
```

### Adding Messages

You can add messages to your chat using the `addMessage()` method. Messages can be from the user or from the system. See [`samples/system-prompts.gs`](samples/system-prompts.gs) for a focused system-prompt example.

```js
// Add a user message
chat.addMessage('Hello, how are you?');

// Add a system message (optional)
chat.addMessage('Answer to the user in a professional way.', true);
```

### Adding Callable Functions to the Chat

You can create functions that the AI can call during the conversation. The `newFunction()` method creates a `FunctionObject`, and `addFunction()` attaches it to your chat. See [`samples/function-calling-basics.gs`](samples/function-calling-basics.gs) and [`samples/function-calling-advanced.gs`](samples/function-calling-advanced.gs) for runnable patterns.

```js
const myFunction = GenAIApp.newFunction()
  .setName('getWeather')
  .setDescription('Retrieve the current weather for a given city.')
  .addParameter('city', 'string', 'The name of the city.');

chat.addFunction(myFunction);
```

From the moment that you add a function to chat, we will use function-calling features.
For more information:
- [https://ai.google.dev/gemini-api/docs/function-calling](https://ai.google.dev/gemini-api/docs/function-calling)
- [https://platform.openai.com/docs/guides/gpt/function-calling](https://platform.openai.com/docs/guides/gpt/function-calling)

### Enable Web Browsing (Optional)

If you want to allow the chat to perform web searches and fetch web pages, enable browsing on your chat instance. See [`samples/web-browsing.gs`](samples/web-browsing.gs).

```js
chat.enableBrowsing(true);
```

To restrict browsing to a specific web page or domain, add the URL as the second argument.

```js
chat.enableBrowsing(true, 'https://support.google.com');
```

### Enable OpenAI Server-Side Compaction (Optional)

Use Responses API native compaction to let OpenAI compact long conversations automatically. See [`samples/configuration-options.gs`](samples/configuration-options.gs) for this alongside other operational settings.

```js
const chat = GenAIApp.newChat()
  .enableCompaction(true)
  .setCompactionThreshold(120000); // minimum: 1000
```

If you only need default behavior, enabling compaction is enough. The default threshold is `10000`.

```js
const chat = GenAIApp.newChat().enableCompaction(true);
```

### Give a Web Page as a Knowledge Base (Optional)

If you do not need broad web search and want to give the model a specific page to read before answering, use `addKnowledgeLink(url)`. See [`samples/knowledge-links.gs`](samples/knowledge-links.gs).

```js
chat.addKnowledgeLink('https://developers.google.com/apps-script/guides/libraries');
```

### Add Image (Optional)

To include an image in the conversation, use the `addImage()` method with a URL or a Blob. See [`samples/image-analysis.gs`](samples/image-analysis.gs).

```js
chat.addImage('https://example.com/image.jpg');
```

### Add File to Chat (Optional)

You can include the contents of a Google Drive file or a Blob in your conversation using the `addFile()` method. This works with both Gemini and OpenAI multimodal models. See [`samples/document-analysis.gs`](samples/document-analysis.gs).

```js
// Add a Google Drive file to the chat context using its Drive file ID
chat.addFile('your-google-drive-file-id');
```

### Add an MCP Connector (Optional)

Use Model Context Protocol (MCP) connectors to let OpenAI Responses API models reach structured data sources such as native Google Workspace endpoints or your own custom MCP servers. See [`samples/mcp-connectors.gs`](samples/mcp-connectors.gs) and [`samples/google-mcp-connector.gs`](samples/google-mcp-connector.gs).

> ⚠️ **Google Workspace Native MCP Requirements:**
> To connect to Google's official MCP endpoints, such as `https://drivemcp.googleapis.com/mcp/v1` or `https://calendarmcp.googleapis.com/mcp/v1`, your Google Apps Script must be linked to a **Standard Google Cloud Project**. In your GCP console, enable both the standard API, such as `drive.googleapis.com` for Drive or `calendar.googleapis.com` for Calendar, and the specific MCP API, such as `drivemcp.googleapis.com` or `calendarmcp.googleapis.com`.

```js
const chat = GenAIApp.newChat();

const nativeDriveConnector = GenAIApp.newConnector()
  .setLabel('Google_Native_Drive')
  .setDescription('Official Google Workspace MCP server for Google Drive')
  .setServerUrl('https://drivemcp.googleapis.com/mcp/v1')
  .setAuthorization(ScriptApp.getOAuthToken())
  .setRequireApproval('never');

chat.addMCP(nativeDriveConnector);

const customConnector = GenAIApp.newConnector()
  .setLabel('Salesforce CRM')
  .setDescription('Query opportunity data from Salesforce via MCP proxy')
  .setServerUrl('https://mcp.example.com/salesforce')
  .setAuthorization('Bearer ' + SALESFORCE_MCP_TOKEN)
  .setRequireApproval('always');

chat.addMCP(customConnector);
```

> **Note on Authorization Format:** Google native connectors using `ScriptApp.getOAuthToken()` do not require the `Bearer` prefix; the token is passed directly to `.setAuthorization()`. Custom MCP servers typically expect the full bearer-token format with a space following `Bearer`. Always check your custom server's authentication requirements.

#### Connector Configuration

- **Google Native endpoints:** Configure a connector using `.setServerUrl()` pointing to the desired service and pass the script's OAuth token via `.setAuthorization(ScriptApp.getOAuthToken())`.
- **Custom MCP servers:** Configure a connector with `.setLabel()`, `.setDescription()`, `.setServerUrl('https://...')`, and optionally `.setAuthorization()` if the server expects a bearer token.
- **Approval workflows:** `.setRequireApproval('never' | 'domain' | 'always')` lets you enforce end-user approval before the model calls the connector.

> ⚠️ **Model Availability:** MCP connectors are currently available only when you run the chat with OpenAI Responses API models. Check your provider's current model documentation before choosing a model override.

### Running the Chat

Once you have set up the chat and added the necessary components, start the conversation by calling `run()`. See [`samples/simple-chat.gs`](samples/simple-chat.gs) for a minimal run and [`samples/multi-model-usage.gs`](samples/multi-model-usage.gs) for model switching.

```js
const response = chat.run({
  model: 'your-model-name', // Optional: set the model to use
  temperature: 0.5, // Optional: set response creativity
  function_call: 'getWeather' // Optional: force the first API response to call a function
});

console.log(response);
```

GenAIApp supports compatible Gemini and OpenAI chat models. Model availability changes over time, so check your provider's current documentation before setting a model override.

⚠️ **Warning:** The `reasoning_effort` parameter is supported only by reasoning-capable OpenAI models and ignored by all others.

### FunctionObject Class

#### Creating a Function

The **FunctionObject** class represents a callable function for the chatbot. It is customizable with names, descriptions, parameters, and execution behavior. See [`samples/function-calling-basics.gs`](samples/function-calling-basics.gs).

```js
const functionObject = GenAIApp.newFunction()
  .setName('searchMovies')
  .setDescription('Search for movies based on a genre.')
  .addParameter('genre', 'string', 'The genre of movies to search for.');
```

#### Configuring Parameters

Function parameters can be configured as required or optional. See [`samples/function-calling-advanced.gs`](samples/function-calling-advanced.gs) for structured extraction patterns.

```js
// Adding required parameter
functionObject.addParameter('year', 'number', 'The year of the movie release.');

// Adding optional parameter
functionObject.addParameter('rating', 'number', 'The minimum rating of movies to return.', true);
```

### VectorStoreObject Class

#### Retrieving Knowledge from Vector Stores

Use a vector store when you want model answers grounded in uploaded source files. GenAIApp supports OpenAI vector stores and Google Gemini File Search Stores behind the same `VectorStoreObject` workflow. See [`samples/vector-store-rag.gs`](samples/vector-store-rag.gs) for a full OpenAI create-upload-query workflow, [`samples/openai-vector-store-quickstart.gs`](samples/openai-vector-store-quickstart.gs) for a short OpenAI example, and [`samples/google-file-search-store-quickstart.gs`](samples/google-file-search-store-quickstart.gs) for a short Gemini File Search Store example.

```js
// OpenAI: pass an OpenAI vector store ID.
const openAiStore = GenAIApp.newVectorStore('openai')
  .initializeFromId('vs_your_openai_vector_store_id');

const openAiChat = GenAIApp.newChat();
openAiChat.addVectorStores(openAiStore.getId());

// Google Gemini: pass the File Search Store resource name returned by getId(),
// for example: fileSearchStores/abc123.
const geminiStore = GenAIApp.newGeminiFileSearchStore()
  .initializeFromId('fileSearchStores/your-google-store-name');

const geminiChat = GenAIApp.newChat();
geminiChat.addVectorStores(geminiStore.getId());
```

For OpenAI, `addVectorStores()` sends `vector_store_ids` to the Responses API file-search tool. For Gemini, the same method sends `file_search_store_names` to the Gemini Interactions API, so use the full File Search Store resource name from `getId()`. Gemini uploads use the direct `uploadToFileSearchStore` media endpoint and wait for the returned operation to complete before the document is available.

To find out more, see the [OpenAI vector store search API](https://platform.openai.com/docs/api-reference/vector_stores/search) and the [Gemini File Search Store API reference](https://ai.google.dev/api/file-search/file-search-stores).

## Reference

### GenAIApp Factory

- `newChat()`: Create a new `Chat` instance. Start with [`samples/simple-chat.gs`](samples/simple-chat.gs).
- `newFunction()`: Create a new `FunctionObject`. See [`samples/function-calling-basics.gs`](samples/function-calling-basics.gs).
- `newConnector()`: Create a new `ConnectorObject` for MCP integrations. See [`samples/mcp-connectors.gs`](samples/mcp-connectors.gs).
- `newVectorStore([providerName])`: Create a new `VectorStoreObject`; omit `providerName` or pass `'openai'` for OpenAI, or pass `'gemini'` for Google Gemini File Search Stores. See [`samples/vector-store-rag.gs`](samples/vector-store-rag.gs), [`samples/openai-vector-store-quickstart.gs`](samples/openai-vector-store-quickstart.gs), and [`samples/google-file-search-store-quickstart.gs`](samples/google-file-search-store-quickstart.gs).
- `newGeminiFileSearchStore()`: Convenience factory for a Gemini File Search Store-backed `VectorStoreObject`. See [`samples/google-file-search-store-quickstart.gs`](samples/google-file-search-store-quickstart.gs).
- `setOpenAIAPIKey(apiKey)`: Set the OpenAI API key.
- `setGeminiAPIKey(apiKey)`: Set the Gemini API key.
- `setGeminiAuth(projectId, region)`: Use Vertex AI authentication. See [`samples/vertex-ai-setup.gs`](samples/vertex-ai-setup.gs).
- `setGlobalMetadata(key, value)`: Attach a key/value pair to every request.
- `setPrivateInstanceBaseUrl(baseUrl)`: Use a custom OpenAI-compatible endpoint.

### Chat

A `Chat` represents a conversation with the model.

- `addMessage(messageContent, [system])`: Add a user or system message. See [`samples/system-prompts.gs`](samples/system-prompts.gs).
- `addFunction(functionObject)`: Attach a `FunctionObject` for function calling. See [`samples/function-calling-basics.gs`](samples/function-calling-basics.gs).
- `addImage(imageInput)`: Include an image URL or Blob in the conversation. See [`samples/image-analysis.gs`](samples/image-analysis.gs).
- `addFile(fileInput)`: Include the content of a Google Drive file or Blob. See [`samples/document-analysis.gs`](samples/document-analysis.gs).
- `addMetadata(key, value)`: Add metadata sent with the next request.
- `getAttributes()`: Retrieve attributes from vector store search results.
- `onlyReturnChunks(bool)`: Return raw chunks from vector store searches. See [`samples/vector-store-rag.gs`](samples/vector-store-rag.gs).
- `setMaxChunks(maxChunks)`: Limit the number of chunks returned by vector stores.
- `getMessages()`: Get the messages as a JSON string.
- `getFunctions()`: Get the functions as a JSON string.
- `disableLogs(bool)`: Disable library logs. See [`samples/configuration-options.gs`](samples/configuration-options.gs).
- `enableBrowsing(bool, [url])`: Allow the model to browse the web, optionally restricted to a URL. See [`samples/web-browsing.gs`](samples/web-browsing.gs).
- `enableCompaction(enabled)`: Enable or disable OpenAI Responses API server-side compaction. See [`samples/configuration-options.gs`](samples/configuration-options.gs).
- `setCompactionThreshold(threshold)`: Set the compaction threshold. The default is `10000`, the minimum is `1000`, and values must be finite numbers.
- `addKnowledgeLink(url)`: Inject the content of a web page into the conversation. See [`samples/knowledge-links.gs`](samples/knowledge-links.gs).
- `addMCP(connectorObject)`: Attach one or more MCP connectors to the chat request. See [`samples/mcp-connectors.gs`](samples/mcp-connectors.gs).
- `setMaximumAPICalls(maxAPICalls)`: Limit the number of API calls in a run. See [`samples/configuration-options.gs`](samples/configuration-options.gs).
- `retrieveLastResponseId()`: Get the last OpenAI response ID returned by `run()`. See [`samples/conversation-continuation.gs`](samples/conversation-continuation.gs).
- `setPreviousResponseId(id)`: Reuse a previous OpenAI response ID to continue a conversation. See [`samples/conversation-continuation.gs`](samples/conversation-continuation.gs).
- `retrieveLastInteractionId()`: Get the last Gemini Interactions API interaction ID returned by `run()`. See [`samples/conversation-continuation.gs`](samples/conversation-continuation.gs).
- `setPreviousInteractionId(id)`: Reuse a previous Gemini interaction ID to continue a conversation. See [`samples/conversation-continuation.gs`](samples/conversation-continuation.gs).
- `warnIfResponseTokenUsageAbove(input_token_threshold)`: Log a warning if input tokens exceed the threshold. It is off by default.
- `addVectorStores(vectorStoreIds)`: Attach OpenAI vector store IDs or Gemini File Search Store resource names for retrieval. See [`samples/vector-store-rag.gs`](samples/vector-store-rag.gs), [`samples/openai-vector-store-quickstart.gs`](samples/openai-vector-store-quickstart.gs), and [`samples/google-file-search-store-quickstart.gs`](samples/google-file-search-store-quickstart.gs).
- `run([advancedParametersObject])`: Execute the chat and return the response. Supports `model`, `temperature`, `reasoning_effort`, `max_tokens`, and `function_call` parameters.

### Function Object

A `FunctionObject` represents a function that can be called by the chat.

- `setName(name)`: Set the function name.
- `setDescription(description)`: Set the function description.
- `addParameter(name, type, description, [isOptional])`: Add a parameter to the function. Parameters are required by default; set `isOptional` to `true` to make a parameter optional.
- `endWithResult(bool)`: End the conversation after the function is executed. See [`samples/function-calling-advanced.gs`](samples/function-calling-advanced.gs).
- `onlyReturnArguments(bool)`: End the conversation and return only the arguments. See [`samples/function-calling-advanced.gs`](samples/function-calling-advanced.gs).

### Vector Store Object

A `VectorStoreObject` represents an OpenAI vector store or a Google Gemini File Search Store. OpenAI is the default provider; use `GenAIApp.newVectorStore('gemini')` or `GenAIApp.newGeminiFileSearchStore()` for Google.

- `setName(newName)`: Set the OpenAI vector store name or Gemini display name.
- `setDescription(newDesc)`: Set the description stored on the wrapper.
- `setChunkingStrategy(maxChunkSize, chunkOverlap)`: Configure OpenAI chunking before uploads.
- `setEmbeddingModel(embeddingModel)`: Set the Gemini File Search Store embedding model resource name before creation.
- `createVectorStore()`: Create the provider-backed store.
- `createFileSearchStore()`: Alias for `createVectorStore()` when using Gemini.
- `initializeFromId(vectorStoreId)`: Initialize from an OpenAI vector store ID or a Gemini File Search Store resource name.
- `getId()`: Get the OpenAI vector store ID or Gemini File Search Store resource name.
- `getName()`: Get the local name/display name.
- `uploadAndAttachFile(blob, attributes)`: Upload a file to OpenAI or upload/import a document into Gemini.
- `uploadAndImportDocument(blob, attributes)`: Alias for `uploadAndAttachFile()` for Gemini-style naming.
- `listFiles()`: List OpenAI files or Gemini documents.
- `listDocuments()`: Alias for `listFiles()`.
- `deleteFile(fileId)`: Delete an OpenAI file or Gemini document.
- `deleteDocument(documentId)`: Alias for `deleteFile()`.
- `deleteVectorStore()`: Delete the store when supported. OpenAI stores can be deleted; Gemini store deletion is not implemented, so delete individual documents instead.

### Connector Object

A `ConnectorObject` represents a Google Workspace or custom MCP connector that can be attached to an OpenAI chat request.

- `setLabel(label)`: Set the identifier used in the chat payload. This is required for custom servers.
- `setDescription(description)`: Provide an optional description visible to the model.
- `setServerUrl(url)`: Use a custom MCP server hosted at the provided HTTPS URL.
- `setConnectorId('gmail' | 'calendar' | 'drive')`: Reference a Google Workspace MCP connector by its predefined ID.
- `setAuthorization(token)`: Override the default OAuth token, for example supply `Bearer ...`.
- `setRequireApproval('never' | 'domain' | 'always')`: Control whether the connector requires user approval before execution.

## Contributing

Contributions are welcome! If you find bugs, have feature requests, or want to contribute code, please submit an issue or pull request on this [GitHub repository](https://github.com/scriptit-fr/GenAIApp).

## License

The **GenAIApp** library is licensed under the Apache License, Version 2.0. You may not use this file except in compliance with the License. For details, see the [LICENSE](http://www.apache.org/licenses/LICENSE-2.0).

---

Happy coding and enjoy building with the **GenAIApp** library!
