# GenAIApp

The **GenAIApp** library is a Google Apps Script library designed for creating, managing, and interacting with LLMs using Gemini and OpenAI's API. The library provides features like text-based conversation, browsing the web, image analysis, and more, allowing you to build versatile AI chat applications that can integrate with various functionalities and external data sources.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Continuous Integration (CI/CD)](#continuous-integration-cicd)
- [Usage](#usage)
  - [Setting API Keys](#setting-api-keys)
  - [Creating a New Chat](#creating-a-new-chat)
  - [Adding Messages](#adding-messages)
  - [Adding Callable Functions to the Chat](#adding-callable-functions-to-the-chat)
  - [Enable Web Browsing (Optional)](#enable-web-browsing-optional)
  - [Enable OpenAI server-side compaction (Optional)](#enable-openai-server-side-compaction-optional)
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
  - [Example 7: Connect to a Custom MCP Server with setServerUrl()](#example-7--connect-to-a-custom-mcp-server-with-setserverurl)
  - [Example 8: Continue a Conversation with previous_response_id](#example-8--continue-a-conversation-with-previous_response_id)
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

## Continuous Integration (CI/CD)

This repository includes GitHub Actions, clasp, and Apps Script project infrastructure for CI-based validation and optional remote test execution. The workflow lives at `.github/workflows/test.yml`; helper files include `appsscript.json`, `.claspignore`, `.clasp.json.example`, and `.github/scripts/prepare-tests.sh`. The real `.clasp.json` file is intentionally ignored by Git and generated dynamically in CI from the `SCRIPT_ID` secret before `clasp push` or `clasp run` is executed.

### What the workflow validates

The **Test** workflow has two layers:

1. **Syntax validation** always runs. It checks out the repository, sets up Node.js, lists every `.gs` file under `src/`, copies each file to a temporary `.js` file, and runs `node --check` to catch JavaScript syntax errors. This job does **not** require clasp credentials, Apps Script authentication, script IDs, or API keys.
2. **Optional remote Apps Script tests** run only when explicitly enabled from the manual workflow inputs. When enabled and correctly configured, CI installs clasp, recreates clasp credentials, generates `.clasp.json`, pushes the repository to the Apps Script test deployment, and invokes configured Apps Script test functions with `clasp run`.

### Workflow triggers

The workflow runs automatically for:

- Pushes to the `main` branch.
- Pull requests targeting the `main` branch.

Automatic runs keep remote/authenticated tests disabled by default, so they provide always-on syntax validation without requiring repository secrets to be available to every event.

To run the workflow manually:

1. Open the repository on GitHub.
2. Go to **Actions**.
3. Select the **Test** workflow.
4. Click **Run workflow**.
5. Choose the branch and set `enable-api-key-tests` and/or `enable-vertex-ai-tests` to `true` only for the remote test modes you want to run.

### Required GitHub secrets

Configure secrets in **GitHub repository → Settings → Secrets and variables → Actions → Repository secrets**. Never commit these values to the repository.

| Secret or variable | Required when | How to obtain it |
| --- | --- | --- |
| `CLASPRC_JSON` | Any remote Apps Script test mode is enabled | Install and authenticate clasp locally with `clasp login`, then copy the full contents of `~/.clasprc.json` into a GitHub repository secret. |
| `SCRIPT_ID` | Any remote Apps Script test mode is enabled | Open the Apps Script project used for CI tests and copy its script ID from **Project Settings → Script ID**, or from the Apps Script URL. |
| `OPEN_AI_API_KEY` | `enable-api-key-tests=true` | Create/copy an OpenAI API key from your OpenAI account and store only the key value as the secret. |
| `GEMINI_API_KEY` | `enable-api-key-tests=true` | Create/copy a Gemini API key from Google AI Studio or the Google Cloud API credentials page and store only the key value as the secret. |
| `VERTEX_AI_PROJECT_ID` | `enable-vertex-ai-tests=true` | Use the Google Cloud project ID for the standard GCP project linked to the Apps Script project and with Vertex AI enabled. |
| `VERTEX_AI_LOCATION` | `enable-vertex-ai-tests=true` | Optional; defaults to `global` if not provided. Configure it as either a repository variable or secret. Use the Vertex AI location you want the tests to call, such as `global` or `us-central1`. |
| `VERTEX_AI_SERVICE_ACCOUNT_JSON` | `enable-vertex-ai-tests=true` | Create a Google Cloud service account with the permissions your Vertex AI test flow requires, create a JSON key if your organization allows key-based CI credentials, and store the complete JSON document as the secret. |

> **Note:** The current Apps Script Vertex AI smoke test authenticates through the linked Apps Script/GCP project via `ScriptApp.getOAuthToken()` after CI pushes the code. The workflow still validates `VERTEX_AI_SERVICE_ACCOUNT_JSON` when Vertex AI tests are requested so the repository has an explicit place for future Vertex AI CI credential needs.

#### Creating and maintaining `CLASPRC_JSON`

`CLASPRC_JSON` is the serialized clasp OAuth credential file used by CI. To generate it:

1. Install clasp locally if needed: `npm install --global @google/clasp`.
2. Run `clasp login`.
3. Complete the browser-based Google OAuth flow with the Google account that has access to the Apps Script test project.
4. Locate the generated file at `~/.clasprc.json`.
5. Copy the entire JSON file contents.
6. Create or update the GitHub repository secret named `CLASPRC_JSON` with that JSON content.

Regenerate and update `CLASPRC_JSON` when:

- CI reports clasp authentication failures, invalid credentials, or expired/revoked tokens.
- You change the Google account used for CI deployments.
- Required Apps Script OAuth scopes change and clasp needs to reauthorize.
- Google revokes the credential, an administrator resets access, or token refresh starts failing.
- Credentials expire. This can happen without warning, so periodic regeneration may be required even if no repository files changed.

### Workflow inputs and behavior

| Manual input | Default | What it controls | Required configuration |
| --- | --- | --- | --- |
| `enable-api-key-tests` | `false` | Runs remote Apps Script tests that use the OpenAI and Gemini API-key paths. | `CLASPRC_JSON`, `SCRIPT_ID`, `OPEN_AI_API_KEY`, and `GEMINI_API_KEY`. |
| `enable-vertex-ai-tests` | `false` | Runs remote Apps Script tests that use the Vertex AI authentication path. | `CLASPRC_JSON`, `SCRIPT_ID`, `VERTEX_AI_PROJECT_ID`, `VERTEX_AI_LOCATION`, and `VERTEX_AI_SERVICE_ACCOUNT_JSON`. |

Skip and failure behavior is intentional:

- If a mode is disabled, the workflow logs a message such as `Skipping API key tests: not enabled via workflow input` or `Skipping Vertex AI tests: not enabled via workflow input`.
- If both auth modes are disabled and clasp credentials are missing or invalid, the remote test steps are skipped and syntax validation can still pass.
- If any auth mode is enabled and `CLASPRC_JSON`, `SCRIPT_ID`, or clasp authentication is missing/invalid, the workflow fails with an actionable message telling you to configure or regenerate clasp credentials.
- If API-key tests are enabled but `OPEN_AI_API_KEY` or `GEMINI_API_KEY` is missing, the workflow fails before pushing or running remote tests.
- If Vertex AI tests are enabled but the required Vertex AI configuration is missing, the workflow fails before pushing or running remote tests.

Example scenarios:

- **Syntax validation only:** leave both inputs set to `false`. This is the default for manual runs and automatic `main`/PR runs.
- **Run only API-key tests:** set `enable-api-key-tests=true` and `enable-vertex-ai-tests=false`; configure `CLASPRC_JSON`, `SCRIPT_ID`, `OPEN_AI_API_KEY`, and `GEMINI_API_KEY`.
- **Run only Vertex AI tests:** set `enable-api-key-tests=false` and `enable-vertex-ai-tests=true`; configure `CLASPRC_JSON`, `SCRIPT_ID`, `VERTEX_AI_PROJECT_ID`, `VERTEX_AI_LOCATION`, and `VERTEX_AI_SERVICE_ACCOUNT_JSON`.
- **Run all remote tests:** set both inputs to `true` and configure every secret listed above.

The remote test functions can be customized with repository variables:

- `API_KEY_TEST_FUNCTIONS`: comma-separated Apps Script function names for API-key tests. Defaults to `testSimpleChatInstance`.
- `VERTEX_AI_TEST_FUNCTIONS`: comma-separated Apps Script function names for Vertex AI tests. Defaults to `testVertexAISimpleChat`, which is generated by the CI prepare script.

### CI to local setup mapping

Local Apps Script development and CI use the same script-level variable names, but they are populated differently. Locally, define the values directly in your Apps Script project or test file. In CI, `.github/scripts/prepare-tests.sh` injects `const` declarations at the top of `src/testFunctions.gs` in the temporary workflow checkout immediately before `clasp push`; those changes are not committed.

| GitHub secret/variable | Local Apps Script variable or setting | Used by |
| --- | --- | --- |
| `OPEN_AI_API_KEY` | `const OPEN_AI_API_KEY = '...'` | `GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY)` in tests. |
| `GEMINI_API_KEY` | `const GEMINI_API_KEY = '...'` | `GenAIApp.setGeminiAPIKey(GEMINI_API_KEY)` in tests. |
| `VERTEX_AI_PROJECT_ID` | `const VERTEX_AI_PROJECT_ID = '...'` | `GenAIApp.setGeminiAuth(VERTEX_AI_PROJECT_ID, VERTEX_AI_LOCATION)` in Vertex AI tests. |
| `VERTEX_AI_LOCATION` | `const VERTEX_AI_LOCATION = 'global'` or another region | Vertex AI location passed to `setGeminiAuth`. |
| `SCRIPT_ID` | Local `.clasp.json` `scriptId` | Determines which Apps Script project clasp pushes/runs against. |
| `CLASPRC_JSON` | Local `~/.clasprc.json` | Authenticates clasp as the selected Google account. |
| `VERTEX_AI_SERVICE_ACCOUNT_JSON` | Service account JSON kept outside Apps Script source | Reserved/validated for Vertex AI CI credentials and future workflow expansion. |

Important local/CI differences:

- Locally, you may already have `.clasp.json` and `~/.clasprc.json`; CI recreates both from secrets every run.
- Locally, test constants can be manually defined; CI injects them automatically and discards the modified test file when the runner is destroyed.
- CI uses `.claspignore` so only `appsscript.json` and `src/` contents are pushed to Apps Script.
- CI remote tests run in the Apps Script environment through `clasp run`, so failures may involve Apps Script permissions, OAuth scopes, linked Google Cloud project configuration, or API quotas in addition to JavaScript errors.

### Troubleshooting CI/CD

#### Clasp credential failures

In CI logs, inspect the **Validate clasp authentication** step. Credential problems commonly show up as:

- `clasp login --status` failures.
- Messages stating that clasp credentials are present but authentication failed, invalid, expired, or revoked.
- A follow-up failure that says auth tests were requested but clasp authentication failed.

To regenerate `CLASPRC_JSON`:

1. On your local machine, run `clasp logout` to clear existing clasp credentials.
2. Run `clasp login`.
3. Complete the OAuth flow in the browser.
4. Confirm local authentication works with `clasp login --status`.
5. Open `~/.clasprc.json` and copy the full JSON content.
6. In GitHub, go to **Settings → Secrets and variables → Actions → Repository secrets**.
7. Update the `CLASPRC_JSON` secret with the new JSON content.
8. Re-run the GitHub Actions workflow.

Before updating CI, it is useful to verify clasp works locally against the intended script:

```bash
clasp login --status
cat > .clasp.json <<'EOF'
{"scriptId":"YOUR_SCRIPT_ID","rootDir":"."}
EOF
clasp push --dry-run
```

#### Common failures and resolutions

| Symptom in CI logs | Likely cause | Resolution |
| --- | --- | --- |
| `401 Unauthorized`, invalid credentials, or expired token messages | `CLASPRC_JSON` is expired or revoked. | Regenerate `CLASPRC_JSON` with `clasp logout` and `clasp login`, then update the GitHub secret. |
| `Script not found` or permission denied for the script | `SCRIPT_ID` is wrong, or the Google account in `CLASPRC_JSON` does not have access to that Apps Script project. | Confirm the Apps Script project ID and share the project with the clasp-authenticated Google account. |
| `CLASPRC_JSON secret not configured - skipping clasp setup` | The secret is missing. | Add `CLASPRC_JSON` if you want remote Apps Script tests. Leave it unset for syntax-only CI. |
| `SCRIPT_ID secret not configured - skipping clasp setup` | The target Apps Script project ID is missing. | Add the `SCRIPT_ID` repository secret. |
| Missing `OPEN_AI_API_KEY` or `GEMINI_API_KEY` errors | API-key tests were enabled without required API-key secrets. | Add both API-key secrets or disable `enable-api-key-tests`. |
| Missing Vertex AI configuration errors | Vertex AI tests were enabled without required Vertex AI values. | Add `VERTEX_AI_PROJECT_ID`, `VERTEX_AI_LOCATION`, and `VERTEX_AI_SERVICE_ACCOUNT_JSON`, or disable `enable-vertex-ai-tests`. |
| Apps Script OAuth scope errors after `clasp run` | The Apps Script project needs reauthorization after scopes changed. | Re-run local clasp authentication and ensure the Apps Script project has the manifest scopes in `appsscript.json`; regenerate `CLASPRC_JSON` if needed. |
| API quota or permission errors from OpenAI, Gemini, or Vertex AI | External API credentials, linked GCP project, billing, API enablement, or quotas are not ready. | Verify the relevant API key/project/service account locally first, then re-run CI. |

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

### Enable OpenAI server-side compaction (optional)

Use Responses API native compaction to let OpenAI compact long conversations automatically.

```js
const chat = GenAIApp.newChat()
  .enableCompaction(true)
  .setCompactionThreshold(120000); // minimum: 1000
```

If you only need default behavior, enabling compaction is enough (default threshold is `10000`):

```js
const chat = GenAIApp.newChat().enableCompaction(true);
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

### Add an MCP Connector (Optional)

Use Model Context Protocol (MCP) connectors to let OpenAI Responses API models reach structured data sources such as native Google Workspace endpoints or your own custom MCP servers.

> ⚠️ **Google Workspace Native MCP Requirements:**
> To connect to Google's official MCP endpoints (e.g., `https://drivemcp.googleapis.com/mcp/v1` or `https://calendarmcp.googleapis.com/mcp/v1`), your Google Apps Script must be linked to a **Standard Google Cloud Project** (the default Apps Script project will return a *403 Forbidden* error).
> In your GCP console, you must enable both the standard API (e.g., `drive.googleapis.com` for Drive or `calendar.googleapis.com` for Calendar) **AND** the specific MCP API (e.g., `drivemcp.googleapis.com` or `calendarmcp.googleapis.com`).

```javascript
const chat = GenAIApp.newChat();

// Google Native Workspace Connector (e.g., Google Drive)
const nativeDriveConnector = GenAIApp.newConnector()
  .setLabel('Google_Native_Drive')
  .setDescription('Official Google Workspace MCP server for Google Drive')
  .setServerUrl('https://drivemcp.googleapis.com/mcp/v1')
  .setAuthorization(ScriptApp.getOAuthToken())
  .setRequireApproval('never');

chat.addMCP(nativeDriveConnector);

// Custom Internal MCP Server
const customConnector = GenAIApp.newConnector()
  .setLabel('Salesforce CRM')
  .setDescription('Query opportunity data from Salesforce via MCP proxy')
  .setServerUrl('https://mcp.example.com/salesforce')
  .setAuthorization('Bearer ' + SALESFORCE_MCP_TOKEN)
  .setRequireApproval('always');

chat.addMCP(customConnector);
```

> **Note on Authorization Format:**
> Google native connectors using `ScriptApp.getOAuthToken()` do NOT require the "Bearer " prefix — the token is passed directly to `.setAuthorization()`. Custom MCP servers (like the Salesforce example using `'Bearer ' + SALESFORCE_MCP_TOKEN`) typically expect the full "Bearer <TOKEN>" format. Always check your custom server's authentication requirements.

#### Connector Configuration

* **Google Native endpoints:** Configure a connector using `.setServerUrl()` pointing to the desired service (e.g., `"https://drivemcp.googleapis.com/mcp/v1"` or `"https://calendarmcp.googleapis.com/mcp/v1"`) and pass the script's OAuth token via `.setAuthorization(ScriptApp.getOAuthToken())`.
* **Custom MCP servers:** Configure a connector with `.setLabel()`, `.setDescription()`, and `.setServerUrl("https://...")`, and optionally `.setAuthorization()` if the server expects a bearer token.
* **Approval workflows:** `.setRequireApproval('never' | 'domain' | 'always')` lets you enforce end-user approval before the model calls the connector.

> ⚠️ **Model Availability:** MCP connectors are currently available only when you run the chat with OpenAI Responses API models (for example, `o4-mini`, `o3`, or `gpt-5.4`).

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
2. OpenAI: "gpt-5.4" | "o4-mini" | "o3" | "gpt-5"

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

// Use Google's Native MCP Endpoint
const gmailConnector = GenAIApp.newConnector()
  .setServerUrl('https://gmailmcp.googleapis.com/mcp/v1') 
  .setLabel('Google_Native_Gmail')
  .setDescription('Official Google Workspace MCP server for Gmail')
  .setAuthorization(ScriptApp.getOAuthToken())
  .setRequireApproval('never');

chat.addMCP(gmailConnector);

const summary = chat.run({ model: 'gpt-5.4', max_tokens: 10000 });
Logger.log(summary);
```

In this example, the connector points directly to Google's Native MCP infrastructure. It requires a linked GCP project with both **Gmail** API and **Gmail MCP API** enabled.

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

const report = chat.run({ model: 'gpt-5.4' });
Logger.log(report);
```

The `setServerUrl()` method points the connector to your MCP gateway, while `setAuthorization()` injects a bearer token or API
key that the proxy expects. Combine these settings with `.setRequireApproval('always')` if you want end users to explicitly
authorize every connector invocation.

### Example 8 : Continue a Conversation with previous_response_id

```javascript
GenAIApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

// First request
const firstChat = GenAIApp.newChat();
firstChat.addMessage("Explain what Google Apps Script libraries are in 3 short bullet points.");

const firstAnswer = firstChat.run({ model: "gpt-5.4" });
Logger.log(firstAnswer);

// Save the response id returned by the OpenAI Responses API
const previousResponseId = firstChat.retrieveLastResponseId();
Logger.log(`Previous response id: ${previousResponseId}`);

// Follow-up request using previous_response_id
const secondChat = GenAIApp.newChat();
secondChat
  .setPreviousResponseId(previousResponseId)
  .addMessage("Now rewrite your previous answer for a beginner in one short paragraph.");

const secondAnswer = secondChat.run({ model: "gpt-5.4" });
Logger.log(secondAnswer);
```


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
- `enableCompaction(enabled)`: Enable/disable OpenAI Responses API server-side compaction (`false` by default).
- `setCompactionThreshold(threshold)`: Set the compaction threshold (`10000` by default, minimum `1000`; finite numbers only).
- `addKnowledgeLink(url)`: Inject the content of a web page into the conversation.
- `addMCP(connectorObject)`: Attach one or more MCP connectors to the chat request.
- `setMaximumAPICalls(maxAPICalls)`: Limit the number of API calls in a run.
- `retrieveLastResponseId()`: Get the last OpenAI response ID returned by `run()`.
- `setPreviousResponseId(id)`: Reuse a previous OpenAI response ID to continue a conversation.
- `warnIfResponseTokenUsageAbove(input_token_threshold)`: Logs a warning if the input tokens are greater than the threshold. Is not on by default.
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