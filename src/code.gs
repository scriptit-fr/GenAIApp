/*
 GenAIApp
 https://github.com/scriptit-fr/GenAIApp
 
 Copyright (c) 2024 Guillemine Allavena - Romain Vialard
 
 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at
 
 http://www.apache.org/licenses/LICENSE-2.0
 
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */

const GenAIApp = (function () {
  let openAIKey = "";
  let geminiKey = "";
  let gcpProjectId = "";
  let region = "";

  let restrictSearch;

  let verbose = true;

  let response_id;

  const apiBaseUrl = "https://api.openai.com";
  let privateInstanceBaseUrl;

  const globalMetadata = {};
  const addedVectorStores = {};

  const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20MB in bytes

  /**
   * @class
   * Class representing a chat.
   */
  class Chat {
    constructor() {
      let messages = []; // messages for OpenAI API
      let contents = []; // contents for Gemini API
      const tools = [];
      let model = "gpt-5"; // default 
      //  OpenAI & Gemini models support a temperature value between 0.0 and 2.0. Models have a default temperature of 1.0.
      let temperature = 1;
      let max_tokens = 1000;
      let browsing = false;
      let reasoning_effort = "medium";
      let knowledgeLink = [];

      let previous_response_id;

      let maxNumOfChunks = 10;
      let onlyChunks = false;
      let retrievedAttributes = [];

      const messageMetadata = {};
      let maximumAPICalls = 30;
      let numberOfAPICalls = 0;

      /**
       * Add a message to the chat.
       * @param {string} messageContent - The message to be added.
       * @param {boolean} [system] - OPTIONAL - True if message from system, False for user. 
       * @returns {Chat} - The current Chat instance.
       */
      this.addMessage = function (messageContent, system) {
        let role = "user";
        if (system) {
          role = "system";
        }
        messages.push({
          role: role,
          content: messageContent
        });
        contents.push({
          role: "user", // Gemini uses 'user' role for both 'user' and 'system' messages
          parts: {
            text: messageContent
          }
        })
        return this;
      };

      /**
       * Add a function to the chat.
       * @param {FunctionObject} functionObject - The function to be added.
       * @returns {Chat} - The current Chat instance.
       */
      this.addFunction = function (functionObject) {
        tools.push({
          type: "function",
          function: functionObject
        });
        return this;
      };

      /**
       * Add an image to the chat.
       * @param {string | Blob} imageInput - The publicly accessible URL of an image or a Blob
       * @returns {Chat} - The current Chat instance.
       */
      this.addImage = function (imageInput) {
        if (typeof imageInput == 'string') {
          // assume it's an url
          // Must be converted to base64 for Gemini
          const response = UrlFetchApp.fetch(imageInput);
          const blob = response.getBlob();
          const base64Image = Utilities.base64Encode(blob.getBytes());
          contents.push({
            role: "user",
            parts: [
              {
                inline_data: {
                  mime_type: blob.getContentType(),
                  data: base64Image
                }
              }
            ]
          });
          // With OpenAI the image url can directly be sent to the model
          messages.push({
            role: "user",
            content: [{
              type: "input_image",
              image_url: imageInput
            }]
          });
        }
        else if (typeof imageInput.getBytes === 'function' &&
          typeof imageInput.getContentType === 'function') {
          // the input is a Blob, to be handled by the addFile() method
          this.addFile(imageInput);
        }
        else {
          throw new Error('Invalid image input provided to addImage() method. Please provide the url of a publicly available image or a Blob.');
        }
        return this;
      };

      /**
       * Adds the content of a file in the prompt
       * @param {string | Blob} fileInput - the ID of a Google Drive or a Blob
       * @returns {Chat} - The current Chat instance.
       */
      this.addFile = function (fileInput) {
        let fileInfo;
        let blobToBase64;

        if (typeof fileInput == 'string') {
          // assume the input is a Google Drive ID
          fileInfo = this._getBlobFromGoogleDrive(fileInput);
          blobToBase64 = Utilities.base64Encode(fileInfo.blob.getBytes());
        }
        else if (typeof fileInput.getBytes === 'function' &&
          typeof fileInput.getContentType === 'function') {
          // the input is a Blob
          const fileBytes = fileInput.getBytes();
          const fileSize = fileBytes.length;
          if (fileSize > MAX_FILE_SIZE) {
            throw new Error(`File too large (${fileSize} bytes). Maximum allowed size is ${MAX_FILE_SIZE} bytes.`);
          }
          blobToBase64 = Utilities.base64Encode(fileBytes);
          fileInfo = {
            mimeType: fileInput.getContentType(),
            fileName: fileInput.getName()
          };
        }
        else {
          throw new Error('Invalid file input provided to addFile() method. Please provide a valid Google Drive file ID or a Blob.');
        }

        // OpenAI
        const contentObj = {};
        if (fileInfo.mimeType.startsWith("image/")) {
          contentObj.type = "input_image";
          contentObj.image_url = `data:${fileInfo.mimeType};base64,${blobToBase64}`;
        }
        else {
          contentObj.type = "input_file";
          contentObj.file_data = `data:${fileInfo.mimeType};base64,${blobToBase64}`;
          contentObj.filename = fileInfo.fileName;
        }
        messages.push({
          role: "user",
          content: [contentObj]
        });

        // Gemini
        contents.push({
          role: 'user',
          parts: [{
            inline_data: {
              mime_type: fileInfo.mimeType,
              data: blobToBase64
            }
          }]
        });
        return this;
      };

      /**
       * Adds an item (key/value pair) to the metadata that will be passed to the OpenAI API.
       * 
       * @param {string} key - The key of the object that should be added.
       * @param {string} value - The value of the object that should be added.
       */
      this.addMetadata = function (key, value = null) {
        messageMetadata[key] = value;
        return this;
      }

      /**
       * Returns the retrieveAttributes list, containing all of the attributes from the chunks that were retrieved through the search() method.
       * @returns {Array} - The list of the attributes from the chunks that were retrieved.
       */
      this.getAttributes = function () {
        return retrievedAttributes;
      }

      /**
       * Defines wether the vector store search should return the raw chunks or send them to the chat.
       * @param {boolean} bool - A boolean to set or not the flag.
       */
      this.onlyReturnChunks = function (bool) {
        if (bool) {
          onlyChunks = true;
        }
        return this;
      }

      /**
       * Sets the limit for how many chunks should be returned by the vector store.
       * @param {number} maxChunks - The number of chunks to return.
       */
      this.setMaxChunks = function (maxChunks) {
        maxNumOfChunks = maxChunks;
        return this;
      }

      /**
       * Get the messages of the chat.
       * returns {string[]} - The messages of the chat.
       */
      this.getMessages = function () {
        return JSON.stringify(messages);
      };

      /**
       * Get the tools of the chat.
       * returns {FunctionObject[]} - The tools of the chat.
       */
      this.getFunctions = function () {
        return JSON.stringify(tools);
      };

      /**
       * Disable logs generated by this library
       * @returns {Chat} - The current Chat instance.
       */
      this.disableLogs = function (bool) {
        if (bool) {
          verbose = false;
        }
        return this;
      };

      /**
       * OPTIONAL
       * 
       * Allow openAI to browse the web.
       * @param {true} scope - set to true to enable full browsing
       * @param {string} [url] - A specific site you want to restrict the search on . 
       * @returns {Chat} - The current Chat instance.
       */
      this.enableBrowsing = function (scope, url) {
        if (scope) {
          browsing = true;
        }
        if (url) {
          restrictSearch = url;
        }
        return this;
      };

      /**
       * Includes the content of a web page in the prompt sent to openAI
       * @param {string} url - the url of the webpage you want to fetch
       * @returns {Chat} - The current Chat instance.
       */
      this.addKnowledgeLink = function (url) {
        if (typeof url === 'string') {
          knowledgeLink.push(url);
        }
        else if (Array.isArray(url)) {
          knowledgeLink.push(...url);
        }
        return this;
      };

      /**
       * If you want to limit the number of calls to the OpenAI API
       * A good way to avoid infinity loops and manage your budget.
       * @param {number} maxAPICalls - 
       */
      this.setMaximumAPICalls = function (maxAPICalls) {
        maximumAPICalls = maxAPICalls;
        return this;
      };

      /**
       * Returns the response Id currently set for the class.
       */
      this.retrieveLastResponseId = function () {
        return response_id;
      };

      /**
       * Sets the previous response Id attribute for the chat (used by Open AI to keep track of conversations)
       * @param {string} previousResponseId - The id of the previous Chat GPT response.
       */
      this.setPreviousResponseId = function (previousResponseId) {
        previous_response_id = previousResponseId;
        return this;
      };

      /**
       * Uses the provided vector store ids (up to 5) with the file search tool for simple RAG.
       * @param {string} vectorStoreIds - A vector store ID or a comma separated list of vector store IDs 
       */
      this.addVectorStores = function (vectorStoreIds) {
        const ids = vectorStoreIds.split(',').map(id => id.trim());
        ids.forEach(id => {
          if (id) addedVectorStores[id] = 1;
        });
        return this;
      };


      this._toJson = function () {
        return {
          messages: messages,
          tools: tools,
          model: model,
          temperature: temperature,
          max_tokens: max_tokens,
          browsing: browsing,
          maximumAPICalls: maximumAPICalls,
          numberOfAPICalls: numberOfAPICalls
        };
      };

      /**
       * Start the chat conversation.
       * Sends all your messages and eventual function to chat GPT.
       * Will return the last chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @param {Object} [advancedParametersObject] OPTIONAL - For more advanced settings and specific usage only. {model, temperature, function_call}
       * @param {"gemini-2.5-pro" | "gemini-2.5-flash" | "gpt-5" | "gpt-4.1" | "o4-mini" | "o3"} [advancedParametersObject.model]
       * @param {number} [advancedParametersObject.temperature]
       * @param {"low" | "medium" | "high"} [advancedParametersObject.reasoning_effort] Only needed for OpenAI reasoning models, defaults to medium
       * @param {number} [advancedParametersObject.max_tokens]
       * @param {string} [advancedParametersObject.function_call]
       * @returns {object} - the last message of the chat 
       */
      this.run = function (advancedParametersObject) {
        model = advancedParametersObject?.model ?? model;
        temperature = advancedParametersObject?.temperature ?? temperature;
        max_tokens = advancedParametersObject?.max_tokens ?? max_tokens;
        reasoning_effort = advancedParametersObject?.reasoning_effort ?? reasoning_effort;

        if (model.includes("gemini")) {
          if (!geminiKey && !gcpProjectId) {
            throw Error("[GenAIApp] - Please set your Gemini API key or GCP project auth using GenAIApp.setGeminiAPIKey(YOUR_GEMINI_API_KEY) or GenAIApp.setGeminiAuth(YOUR_PROJECT_ID, REGION)");
          }
        }
        else {
          if (!openAIKey) {
            throw Error("[GenAIApp] - Please set your OpenAI API key using GenAIApp.setOpenAIAPIKey(yourAPIKey)");
          }
        }

        if ((model.startsWith("o") || model.includes("gemini") || model.includes("gpt-5")) && browsing && max_tokens < 10000) {
          console.warn(`[GenAIApp] - Browsing enabled on ${model} with max_tokens=${max_tokens} (< 10000). This will likely truncate the response. Consider chat.run({ max_tokens: 20000 }).`);
        }

        if (knowledgeLink.length > 0) {
          let knowledge = "";
          knowledgeLink.forEach(url => {
            const urlContent = _urlFetch(url);
            knowledge += `${url}: \n\n ${urlContent}\n\n`;
          })
          if (!knowledge) {
            throw Error(`[GenAIApp] - The webpage of at least one of the URLs didn't respond, please change the url of the addKnowledgeLink() function.`);
          }
          messages.push({
            role: "system",
            content: `Information to help with your response : ${knowledge}`
          });
          contents.push({
            role: "user",
            parts: {
              text: `Information to help with your response : ${knowledge}`
            }
          })
          knowledgeLink = [];
        }

        let payload;
        if (model.includes("gemini")) {
          payload = this._buildGeminiPayload(advancedParametersObject);
        }
        else {
          payload = this._buildOpenAIPayload();
        }

        let responseMessage;
        if (numberOfAPICalls <= maximumAPICalls) {
          let endpointUrl = apiBaseUrl + "/v1/responses";
          if (privateInstanceBaseUrl) {
            endpointUrl = privateInstanceBaseUrl + "/v1/responses?api-version=preview";
          }
          if (model.includes("gemini")) {
            if (geminiKey) {
              // Public endpoint / Generative Language API
              // https://console.cloud.google.com/apis/api/generativelanguage.googleapis.com
              endpointUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
            }
            else {
              // Enterprise endpoint / Vertex AI API
              // https://console.cloud.google.com/apis/api/aiplatform.googleapis.com
              // requires scope "https://www.googleapis.com/auth/cloud-platform.read-only" in access token
              if (region) {
                endpointUrl = `https://${region}-aiplatform.googleapis.com/v1/projects/${gcpProjectId}/locations/${region}/publishers/google/models/${model}:generateContent`;
              }
              else {
                endpointUrl = `https://aiplatform.googleapis.com/v1/projects/${gcpProjectId}/locations/global/publishers/google/models/${model}:generateContent`;
              }
            }
          }
          responseMessage = _callGenAIApi(endpointUrl, payload);
          numberOfAPICalls++;
        }
        else {
          throw new Error(`[GenAIApp] - Too many calls to genAI API: ${numberOfAPICalls}`);
        }

        if (Array.isArray(responseMessage)) {
          const fileSearchCall = responseMessage.filter(item => item.type === "file_search_call");
          if (fileSearchCall.length > 0) {
            // In case of file search, handle the possible use of onlyReturnChunks() and store files attributes
            // https://platform.openai.com/docs/guides/tools-file-search
            retrievedAttributes = [];
            const retrievedChunks = fileSearchCall.flatMap(call =>
              Array.isArray(call?.results) ? call.results : []
            );
            if (verbose) {
              console.log(`[GenAIApp] - File search performed: retrieved ${retrievedChunks.length} chunks (max_num_results=${maxNumOfChunks}).`);
            }
            for (const chunk of retrievedChunks) {
              if (chunk?.attributes != null) {
                retrievedAttributes.push(chunk.attributes);
              }
            }
            if (onlyChunks) {
              return retrievedChunks;
            }
          }
        }

        if (tools.length > 0) {
          // Check if AI model wanted to call a function
          if (model.includes("gemini")) {
            if (responseMessage?.parts?.some(p => p?.functionCall)) {
              contents = _handleGeminiToolCalls(responseMessage, tools, contents);
              // check if endWithResults or onlyReturnArguments
              if (contents[contents.length - 1].role == "model") {
                if (contents[contents.length - 1].parts.text == "endWithResult") {
                  if (verbose) {
                    console.log("[GenAIApp] - Conversation stopped because end function has been called");
                  }
                  // Do not return anything specific as the goal is simply to end here.
                  return "OK";
                }
                else if (contents[contents.length - 1].parts.text == "onlyReturnArguments") {
                  if (verbose) {
                    console.log("[GenAIApp] - Conversation stopped because argument return has been enabled - No function has been called");
                  }
                  // return the argument(s) of the last function called
                  return contents[contents.length - 2].parts
                    .find(p => p && p.functionCall)
                    ?.functionCall.args;
                }
              }
            }
            else {
              // if no function has been found, stop here
              const part = responseMessage?.parts?.find(p => !p.thought && p.text);
              return part?.text || null;
            }
          }
          else {
            const functionCalls = responseMessage.filter(item => item.type === "function_call");
            if (functionCalls.length > 0) {
              messages = _handleOpenAIToolCalls(responseMessage, tools, messages);
              // check if endWithResults or onlyReturnArguments
              if (messages[messages.length - 1].role == "system") {
                if (messages[messages.length - 1].content == "endWithResult") {
                  if (verbose) {
                    console.log("[GenAIApp] - Conversation stopped because end function has been called");
                  }
                  // Do not return anything specific as the goal is simply to end here.
                  return "OK";
                }
                else if (messages[messages.length - 1].content == "onlyReturnArguments") {
                  if (verbose) {
                    console.log("[GenAIApp] - Conversation stopped because argument return has been enabled - No function has been called");
                  }
                  // return the argument(s) of the last function called
                  return _parseResponse(messages[messages.length - 3].arguments);
                }
              }
              // Use the previous_response_id parameter to pass reasoning items from previous responses
              // This allows the model to continue its reasoning process to produce better results in the most token-efficient manner.
              // https://platform.openai.com/docs/guides/reasoning#keeping-reasoning-items-in-context
              previous_response_id = response_id;
            }
            else {
              // if no function has been found, stop here
              const messageItem = responseMessage?.find?.(item => item.type === "message");
              return messageItem?.content?.find(part => part?.text)?.text || null;
            }
          }
          if (advancedParametersObject) {
            return this.run(advancedParametersObject);
          }
          else {
            return this.run();
          }
        }
        else {
          if (model.includes("gemini")) {
            const part = responseMessage?.parts?.find(p => !p.thought && p.text);
            return part?.text || null;
          }
          else {
            return responseMessage
              .find(item => item.type === "message")?.content
              ?.find(part => part?.text)?.text || null;
          }
        }
      }

      /**
       * Builds and returns a payload for an OpenAI API call, incorporating advanced parameters and 
       * tool-specific configurations for browsing, image description, and assistant functionalities. 
       * Configures tool choices based on recent interactions, message content, and options in 
       * `advancedParametersObject`.
       *
       * @private
       * @returns {Object} - The payload object, configured with messages, model settings, and tool selections 
       *                     for OpenAI's API.
       * @throws {Error} If an incompatible model is selected with certain functionalities (e.g., Gemini model with assistant).
       */
      this._buildOpenAIPayload = function () {
        if (globalMetadata) {
          Object.assign(messageMetadata, globalMetadata);
        }

        let systemInstructions = "";
        const userMessages = [];

        for (const message of messages) {
          if (message.role === "system") {
            systemInstructions += message.content + "\n";
          }
          else {
            userMessages.push(message);
          }
        }

        const payload = {
          model: model,
          instructions: systemInstructions,
          input: userMessages,
          max_output_tokens: max_tokens,
          previous_response_id: previous_response_id,
          tools: [],
          metadata: messageMetadata
        };

        if (tools.length > 0) {
          // the user has added functions, enable function calling
          const toolsPayload = tools.map(tool => ({
            type: "function",
            name: tool.function._toJson().name,
            description: tool.function._toJson().description,
            parameters: tool.function._toJson().parameters,
          }));
          payload.tools = toolsPayload;

          if (!payload.tool_choice) {
            payload.tool_choice = 'auto';
          }
        }

        if (model.startsWith('o') || model.includes("gpt-5")) {
          payload.reasoning = {
            "effort": reasoning_effort
          }
        }

        if (browsing) {
          payload.tools.push({
            type: "web_search_preview"
          });

          if (restrictSearch) {
            messages.push({
              role: "user", // upon testing, this instruction has better results with user role instead of system
              content: `You are only able to search for information on ${restrictSearch}, restrict your search to this website only.`
            });
          }
        }

        if (Object.keys(addedVectorStores).length > 0 && numberOfAPICalls < 1) {
          const fileSearchTool = {
            type: "file_search",
            vector_store_ids: Object.keys(addedVectorStores),
            max_num_results: maxNumOfChunks
          };

          payload.tools = payload.tools || [];
          payload.tools.push(fileSearchTool);

          payload.include = ["file_search_call.results"];
        }
        return payload;
      }

      /**
       * Builds and returns a payload for a Gemini API call, configuring content, model parameters, 
       * and tool settings based on advanced options and feature flags such as browsing. 
       * Adapts the payload for specific function calls and tools.
       *
       * @private
       * @param {Object} advancedParametersObject - An object with optional advanced parameters, 
       *                                            such as function call preferences.
       * @returns {Object} - The configured payload object for the Gemini API, including content, model settings, 
       *                     generation configuration, and available tools.
       * @throws {Error} If an incompatible feature is selected (e.g., assistant usage with the Gemini model).
       */
      this._buildGeminiPayload = function (advancedParametersObject) {
        const payload = {
          'contents': contents,
          'model': model,
          'generationConfig': {
            maxOutputTokens: max_tokens,
            temperature: temperature,
          },
          'tool_config': {
            function_calling_config: {
              mode: "AUTO"
            }
          },
          tools: []
        };

        if (advancedParametersObject?.function_call) {
          payload.tool_config.function_calling_config.mode = "ANY";
          payload.tool_config.function_calling_config.allowed_function_names = advancedParametersObject.function_call;
          delete advancedParametersObject.function_call;
        }

        if (tools.length > 0) {
          // the user has added functions, enable function calling
          const payloadTools = Object.keys(tools).map(t => {
            const toolFunction = tools[t].function._toJson();

            const parameters = toolFunction.parameters;
            if (parameters?.type) {
              toolFunction.parameters.type = parameters.type.toUpperCase();
            }

            return {
              name: toolFunction.name,
              description: toolFunction.description,
              parameters: toolFunction.parameters
            };
          });

          payload.tools = [{
            functionDeclarations: payloadTools
          }];
        }

        if (browsing) {
          tools.push({
            google_search: "",
          });
          payload.tools.push({
            url_context: {}
          });
          payload.tools.push({
            google_search: {}
          });
        }

        return payload;
      }

      /**
       * Get a blob from a Google Drive file ID
       *
       * @private
       * @param {string} fileId - The ID of the Google Drive file
       * @returns {{
       * blob: Blob,
       * fileName: fileName,
       * mimeType: mimeType
       * }} - The data as a blob.
       */
      this._getBlobFromGoogleDrive = function (fileId) {
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();
        const fileName = file.getName();
        const fileSize = file.getSize();
        let blob;
        // Gemini has a 20MB limit for API requests
        if (fileSize > MAX_FILE_SIZE) {
          throw new Error(`[GenAIApp] - File too large (${fileSize} bytes). Maximum allowed size is ${MAX_FILE_SIZE} bytes.`);
        }

        switch (mimeType) {
          case 'application/pdf':
          case 'text/plain':
          case 'image/png':
          case 'image/jpeg':
          case 'image/gif':
          case 'image/webp':
            blob = file.getBlob();
            break;

          case 'application/vnd.google-apps.spreadsheet':
          case 'application/vnd.google-apps.document':
          case 'application/vnd.google-apps.presentation':
            let fileBlobUrl;
            switch (mimeType) {
              case 'application/vnd.google-apps.spreadsheet':
                fileBlobUrl =
                  'https://docs.google.com/spreadsheets/d/' +
                  fileId +
                  '/export?format=pdf';
                break;
              case 'application/vnd.google-apps.document':
                fileBlobUrl =
                  'https://docs.google.com/document/d/' +
                  fileId +
                  '/export?format=pdf';
                break;
              case 'application/vnd.google-apps.presentation':
                fileBlobUrl =
                  'https://docs.google.com/presentation/d/' +
                  fileId +
                  '/export?format=pdf';
                break;
            }
            const token = ScriptApp.getOAuthToken();
            const response = UrlFetchApp.fetch(fileBlobUrl, {
              headers: {
                Authorization: 'Bearer ' + token
              }
            });
            blob = response.getBlob();
            break;

          default:
            throw new Error('[GenAIApp] - Invalid file type provided to addFile() method.');
        }

        return {
          fileName: fileName,
          mimeType: mimeType,
          blob: blob
        };
      }
    }
  }

  /**
 * @class
 * Class representing an Open AI Vector Store.
 */
  class VectorStoreObject {
    constructor() {
      let name = "";
      let description = "";
      let id = null;
      let max_chunk_size = 800;
      let chunk_overlap = 400;

      /**
       * Sets the vector store's name
       * @param {string} newName - The name to assign to the vector store.
       * @returns {VectorStoreObject}
       */
      this.setName = function (newName) {
        name = newName;
        return this;
      };

      /**
       * Sets the description of the vector store.
       * @param {string} newDesc - The description to assign to the vector store.
       * @returns {VectorStoreObject}
       */
      this.setDescription = function (newDesc) {
        description = newDesc;
        return this;
      };

      /**
       * Sets the chunking strategy for the Vector Store.
       * @param {number} maxChunkSize - The maximum token size of a chunk (max is 4096, defaults to 800).
       * @param {number} chunkOverlap - The chunk overlap to apply. Cannot exceed half of the maxChunkSize (defaults to 400).
       */
      this.setChunkingStrategy = function (maxChunkSize, chunkOverlap) {
        max_chunk_size = maxChunkSize;
        chunk_overlap = chunkOverlap;
        return this
      }

      /**
       * Creates the Open AI vector store. A name must be assigned before calling this function.
       * @returns {VectorStoreObject}
       */
      this.createVectorStore = function () {
        if (!name) throw new Error("[GenAIApp] - Please specify your Vector Store name using the GenAiApp.newVectorStore().setName() method before creating it.");
        try {
          id = _createOpenAiVectorStore(name);
        }
        catch (e) {
          console.error(`Error creating the vector store : ${e}`);
        }
        return this;
      };

      /**
       * Initializes a new vector store object from an existing Open AI vector store id. This allows a user to interact with an existing vector store.
       * 
       * @param {string} vectorStoreId - The Open AI API vector store id. 
       */
      this.initializeFromId = function (vectorStoreId) {
        try {
          const vectorStoreName = _retrieveVectorStoreInformation(vectorStoreId);
          name = vectorStoreName;
          id = vectorStoreId;
        }
        catch (e) {
          console.error(`[GenAIApp] - Could not initialize vector store object from id : ${e}`);
        }
        return this;
      }

      /**
       * Returns the vector store id.
       * @returns {string} - The id of the vector store.
       */
      this.getId = function () {
        return id;
      };

      /**
       * Uploads a file to Open AI storage and attaches it to the vector store.
       * @param {Blob} blob - File to upload.
       * @param {Object} attributes - The JSON object containing the attributes to attach to the vector store. Per Open AI's documentation, must contain a max of 16 key-value pairs (both strings, up to 64 characters for keys, and up to 500 characters for values).
       * @returns {object} - The raw JSON chunks returned by the vector store.
       */
      this.uploadAndAttachFile = function (blob, attributes = {}) {
        if (!id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object with GenAiApp.newVectorStore().setName().initializeFromId() or GenAiApp.newVectorStore().setName().createVectorStore() before attaching files.");
        try {
          const uploadedFileId = _uploadFileToOpenAIStorage(blob);
          const attachedFileId = _attachFileToVectorStore(uploadedFileId, id, attributes, max_chunk_size, chunk_overlap);
          return attachedFileId;
        }
        catch (e) {
          Logger.log({
            message: `Unable to upload and attach the file to the vector store : ${e}`,
            fileBlob: blob
          });
        }
      };

      /**
       * Lists the files attached to the vector store.
       * @returns {Array} - An array containing the files attached to the vector store.
       */
      this.listFiles = function () {
        if (!id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object with GenAiApp.newVectorStore().setName().initializeFromId() or GenAiApp.newVectorStore().setName().createVectorStore() before listing files.");
        try {
          const listedFiles = _listFilesInVectorStore(id);
          return listedFiles;
        }
        catch (e) {
          Logger.log({
            message: `An error occured when trying to list files from vector store : ${e}`,
            vectorStoreId: id
          });
        }
      };

      /**
       * Deletes a file from the vector store.
       * @param {string} fileId - The ID of the file to delete.
       */
      this.deleteFile = function (fileId) {
        if (!fileId) throw new Error("[GenAIApp] - Please pass an Open AI storage file ID to the deleteFile(fileId) function. You can retrieve the file ID through the Open AI Files API or directly through the platform.");
        if (!id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object with GenAiApp.newVectorStore().setName().initializeFromId() or GenAiApp.newVectorStore().setName().createVectorStore() before deleting files.");
        try {
          const deleteId = _deleteFileInVectorStore(id, fileId);
          return deleteId;
        }
        catch (e) {
          Logger.log({
            message: `An error occured when trying to delete the file from the vector store : ${e}`,
            vectorStoreId: id,
            fileId: fileId
          });
        }

      };

      /**
       * Deletes the vector store from Open AI.
       * @returns {string} - The delete ID.
       */
      this.deleteVectorStore = function () {
        if (!id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object with GenAiApp.newVectorStore().setName().initializeFromId() or GenAiApp.newVectorStore().setName().createVectorStore() before being deleted.");
        try {
          const deleteId = _deleteVectorStore(id);
          id = null;
          return deleteId;
        }
        catch (e) {
          Logger.log({
            message: `An error occured when trying to delete the vector store : ${e}`,
            vectorStoreId: id
          });
        }
      };

      /**
       * Returns the JSON object with name, description, and ID.
       * @returns {Object}
       */
      this._toJson = function () {
        return {
          name: name,
          description: description,
          id: id
        };
      };
    }
  }

  /**
   * @class
   * Class representing a function known by function calling model
   */
  class FunctionObject {

    constructor() {
      let name = '';
      let description = '';
      const properties = {};
      const required = [];
      const argumentsInRightOrder = [];
      let endingFunction = false;
      let onlyArgs = false;

      /**
       * Sets the name of a function.
       * @param {string} nameOfYourFunction - The name to set for the function.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setName = function (nameOfYourFunction) {
        name = nameOfYourFunction;
        return this;
      };

      /**
       * Sets the description of a function.
       * @param {string} descriptionOfYourFunction - The description to set for the function.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setDescription = function (descriptionOfYourFunction) {
        description = descriptionOfYourFunction;
        return this;
      };

      /**
       * OPTIONAL
       * If enabled, the conversation with the chat will automatically end when this function is called.
       * Default : false, eg the function is sent to the chat that will decide what the next action shoud be accordingly. 
       * @param {boolean} bool - Whether or not you wish for the option to be enabled. 
       * @returns {FunctionObject} - The current Function instance.
       */
      this.endWithResult = function (bool) {
        if (bool) {
          endingFunction = true;
        }
        return this;
      }

      /**
       * Adds a property (an argument) to the function.
       * Note: Parameters are required by default. Set 'isOptional' to true to make a parameter optional.
       *
       * @param {string} name - The property name.
       * @param {string} type - The property type.
       * @param {string} description - The property description.
       * @param {boolean} [isOptional] - To set if the argument is optional (default: false).
       * @returns {FunctionObject} - The current Function instance.
       */

      this.addParameter = function (name, type, description, isOptional = false) {
        let itemsType;

        if (String(type).includes("Array")) {
          const startIndex = type.indexOf("<") + 1;
          const endIndex = type.indexOf(">");
          itemsType = type.slice(startIndex, endIndex);
          type = "array";
        }

        properties[name] = {
          type: type,
          description: description
        };

        if (type === "array") {
          if (itemsType) {
            properties[name]["items"] = {
              type: itemsType
            }
          }
          else {
            throw Error("[GenAIApp] - Please precise the type of the items contained in the array when calling addParameter. Use format Array.<itemsType> for the type parameter.");
          }
        }

        argumentsInRightOrder.push(name);
        if (!isOptional) {
          required.push(name);
        }
        return this;
      }

      /**
       * OPTIONAL
       * If enabled, the conversation will automatically end when this function is called and the chat will return the arguments in a stringified JSON object.
       * Default : false
       * @param {boolean} bool - Whether or not you wish for the option to be enabled. 
       * @returns {FunctionObject} - The current Function instance.
       */
      this.onlyReturnArguments = function (bool) {
        if (bool) {
          onlyArgs = true;
        }
        return this;
      }

      this._toJson = function () {
        return {
          name: name,
          description: description,
          parameters: {
            type: "object",
            properties: properties,
            required: required
          },
          argumentsInRightOrder: argumentsInRightOrder,
          endingFunction: endingFunction,
          onlyArgs: onlyArgs
        };
      };
    }
  }

  /**
   * Makes an API call to the specified GenAI endpoint (either OpenAI or Google) with a payload
   * and handles authentication, retries on rate limits and server errors, and response parsing.
   * This function is designed for internal use and includes exponential backoff for retries.
   *
   * @private
   * @param {string} endpoint - The API endpoint URL to call, e.g., OpenAI or Google GenAI endpoint.
   * @param {Object} payload - The request payload to send in JSON format, including request data like max_tokens.
   * @returns {object} - The response message from the GenAI API.
   * @throws {Error} If the API call fails after the maximum number of retries.
   */
  function _callGenAIApi(endpoint, payload) {
    let authMethod = 'Bearer ' + openAIKey;
    if (endpoint.includes("google")) {
      if (geminiKey) {
        // Header name different for Google API key
        authMethod = null;
      }
      else {
        authMethod = 'Bearer ' + ScriptApp.getOAuthToken();
      }
    }
    const maxRetries = 5;
    let retries = 0;
    let success = false;

    let responseMessage, finish_reason;
    while (retries < maxRetries && !success) {
      const headers = {
        'Content-Type': 'application/json',
      };
      if (authMethod) {
        headers['Authorization'] = authMethod;
      }
      else if (geminiKey) {
        // use an HTTP header instead of including the API key in the query parameters.
        headers['x-goog-api-key'] = geminiKey;
      }
      const options = {
        method: 'post',
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      let response;
      // if the ErrorHandler library is loaded and supports backoff, use it (https://github.com/RomainVialard/ErrorHandler)
      if (typeof ErrorHandler !== 'undefined' && typeof ErrorHandler.urlFetchWithExpBackOff === 'function') {
        response = ErrorHandler.urlFetchWithExpBackOff(endpoint, options);
      }
      else {
        try {
          response = UrlFetchApp.fetch(endpoint, options);
        }
        catch (err) {
          if (verbose) {
            console.warn(`[GenAIApp] - Network error calling ${payload.model}: ${err.message}. Retrying (${retries + 1}/${maxRetries})`);
          }
          const delay = Math.pow(2, retries) * 1000;
          Utilities.sleep(delay);
          retries++;
          continue;
        }
      }
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        // The request was successful, exit the loop.
        const parsedResponse = JSON.parse(response.getContentText());
        if (endpoint.includes("google")) {
          const firstCandidate = parsedResponse.candidates?.[0];
          responseMessage = firstCandidate?.content || null;
          finish_reason = firstCandidate?.finishReason || null;
        }
        else {
          responseMessage = parsedResponse.output;
          response_id = parsedResponse.id;
          finish_reason = parsedResponse.status;
        }
        if (finish_reason == "length" || finish_reason == "incomplete" || finish_reason == "MAX_TOKENS") {
          console.warn(`[GenAIApp] - ${payload.model} response could not be completed because of an insufficient amount of tokens. To resolve this issue, you can increase the amount of tokens like this : chat.run({max_tokens: XXXX}).`);
        }
        success = true;
      }
      else if (responseCode === 429) {
        console.warn(`[GenAIApp] - Rate limit reached when calling ${payload.model}, will automatically retry in a few seconds.`);
        // Rate limit reached, wait before retrying.
        const delay = Math.pow(2, retries) * 1000; // Delay in milliseconds, starting at 1 second.
        Utilities.sleep(delay);
        retries++;
      }
      else if (responseCode === 503 || responseCode === 500 || responseCode === 502) {
        // The server is temporarily unavailable, or an issue occured on OpenAI servers. wait before retrying.
        // https://platform.openai.com/docs/guides/error-codes/api-errors
        const delay = Math.pow(2, retries) * 1000; // Delay in milliseconds, starting at 1 second.
        Utilities.sleep(delay);
        retries++;
        if (verbose) {
          console.warn(`[GenAIApp] - Request to ${payload.model} failed with response code ${responseCode} - ${response.getContentText()}, retrying (${retries}/${maxRetries})`);
        }
      }
      else {
        // The request failed for another reason, log the error and exit the loop.
        console.error(`[GenAIApp] - Request to ${payload.model} failed with response code ${responseCode} - ${response.getContentText()}`);
        break;
      }
    }

    if (!success) {
      throw new Error(`[GenAIApp] - Failed to call ${payload.model} after ${retries} retries.`);
    }

    if (verbose) {
      Logger.log({
        message: `[GenAIApp] - Got response from ${payload.model}`,
        responseMessage: responseMessage
      });
    }
    return responseMessage;
  }

  /**
   * Processes tool calls from a Gemini response message, managing the sequence of function executions, argument handling, 
   * and special actions for web search and URL fetch calls. It dynamically builds a conversation flow and manages the end
   * condition based on tool specifications.
   *
   * @private
   * @param {Object} responseMessage - The response message from Gemini containing tool calls.
   * @param {Array} tools - List of available tools, each with metadata including function names and argument requirements.
   * @param {Array} contents - Array representing the conversational content, updated with each tool call and its result.
   * @returns {Array} - The updated contents array, representing the conversation flow with function calls and responses.
   */
  function _handleGeminiToolCalls(responseMessage, tools, contents) {
    // Append function call to contents
    // The thought signature is also sent back
    // https://ai.google.dev/gemini-api/docs/function-calling?example=meeting#thinking
    // Note: thoughtSignature seems to be only included in Generative Language API, not Vertex AI API
    contents.push(responseMessage);

    const parts = (responseMessage && responseMessage.parts) || [];
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i] || {};
      if (!part.functionCall || !part.functionCall.name) continue;

      const functionName = part.functionCall.name;
      const functionArgs = part.functionCall.args;

      let argsOrder = [];
      let endWithResult = false;
      let onlyReturnArguments = false;

      for (const t in tools) {
        const currentFunction = tools[t].function._toJson();
        if (currentFunction.name == functionName) {
          argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order
          endWithResult = currentFunction.endingFunction;
          onlyReturnArguments = currentFunction.onlyArgs;
          break;
        }
      }

      // No actual call to the function
      if (onlyReturnArguments) {
        contents.push({
          "role": "model",
          "parts": {
            text: "onlyReturnArguments"
          }
        });
        return contents;
      }

      let functionResponse = _callFunction(functionName, functionArgs, argsOrder);
      if (verbose) {
        console.log(`[GenAIApp] - function ${functionName}() called by Gemini.`);
      }
      if (typeof functionResponse != "string") {
        if (typeof functionResponse == "object") {
          functionResponse = JSON.stringify(functionResponse);
        }
        else {
          functionResponse = String(functionResponse);
        }
      }

      // Append result of the function execution to contents
      // https://ai.google.dev/gemini-api/docs/function-calling?example=meeting#step-4
      contents.push({
        role: 'user',
        parts: [{
          functionResponse: {
            name: functionName,
            response: { functionResponse }
          }
        }]
      });

      if (endWithResult) {
        // User defined that if this function has been called, we do not call back the AI endpoint.
        contents.push({
          "role": "model",
          "parts": {
            text: "endWithResult"
          }
        });
        return contents;
      }
    }
    return contents;
  }

  /**
   * Processes OpenAI tool calls from a response message, handling function executions, argument ordering, and
   * managing specific actions like web searches and URL fetches. This function updates the conversation flow with 
   * each tool call result and manages special conditions based on tool specifications.
   *
   * @private
   * @param {Object} responseMessage - The response message from OpenAI, containing details about tool calls.
   * @param {Array} tools - Array of tool objects, each containing metadata about functions, argument orders, and conditions.
   * @param {Array} messages - Array representing the conversation flow, which is updated with tool call results and system messages.
   * @returns {Array} - The updated messages array, representing the conversation flow with function calls, results, and system responses.
   */
  function _handleOpenAIToolCalls(responseMessage, tools, messages) {
    responseMessage.forEach(item => messages.push(item));
    for (const tool_call of responseMessage) {
      if (tool_call.type == "function_call") {

        const functionName = tool_call.name;
        const functionArgs = _parseResponse(tool_call.arguments);

        let argsOrder = [];
        let endWithResult = false;
        let onlyReturnArguments = false;

        for (const t in tools) {
          const currentFunction = tools[t].function._toJson();
          if (currentFunction.name == functionName) {
            argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order
            endWithResult = currentFunction.endingFunction;
            onlyReturnArguments = currentFunction.onlyArgs;
            break;
          }
        }

        // No actual call to the function
        if (onlyReturnArguments) {
          messages.push(tool_call);
          messages.push({
            "role": "system",
            "content": "onlyReturnArguments"
          });
          return messages;
        }

        // Call the function
        let functionResponse = _callFunction(functionName, functionArgs, argsOrder);
        if (typeof functionResponse != "string") {
          if (typeof functionResponse == "object") {
            functionResponse = JSON.stringify(functionResponse);
          }
          else {
            functionResponse = String(functionResponse);
          }
        }
        if (verbose) {
          console.log(`[GenAIApp] - function ${functionName}() called by OpenAI.`);
        }

        if (endWithResult) {
          // User defined that if this function has been called, we do not call back the AI endpoint.
          messages.push(tool_call);
          messages.push({
            "type": "function_call_output",
            "call_id": tool_call.call_id,
            "output": functionResponse
          });
          messages.push({
            "role": "system",
            "content": "endWithResult"
          });
          return messages;
        }
        else {
          // Reset the previous messages, 
          // we will rely instead on the previous_response_id parameter to pass reasoning items from previous responses
          // This allows the model to continue its reasoning process to produce better results in the most token-efficient manner.
          // https://platform.openai.com/docs/guides/reasoning#keeping-reasoning-items-in-context
          if (!messages.every(msg => msg.type === "function_call_output" || msg.role === "system")) {
            // Reset only if it contains other messages than function_call_output 
            // to allow for parallel function calling
            // https://platform.openai.com/docs/guides/function-calling#parallel-function-calling
            // Preserve only system messages
            messages = messages.filter(msg => msg.role === "system");
          }
          // Provide function call results to the model
          messages.push({
            "type": "function_call_output",
            "call_id": tool_call.call_id,
            "output": functionResponse
          });
        }
      }
    }
    return messages;
  }

  /**
   * Calls an internal or dynamically referenced function based on the provided function name, 
   * with arguments parsed and ordered as specified. Handles specific functions directly and 
   * supports dynamic function calling from the global context.
   *
   * @private
   * @param {string} functionName - The name of the function to call, which may be an internal or global function.
   * @param {Object} jsonArgs - An object containing the arguments as key-value pairs, extracted from the tool call.
   * @param {Array} argsOrder - An array specifying the required order of arguments for the function.
   * @returns {*} - The result of the function call, or a success message if the function executes without a return.
   * @throws {Error} If the specified function is not found or is not a function.
   */
  function _callFunction(functionName, jsonArgs, argsOrder) {
    // Parse JSON arguments
    const argsObj = jsonArgs;
    const argsArray = argsOrder.map(argName => argsObj[argName]);

    // Call the function dynamically
    if (globalThis[functionName] instanceof Function) {
      const functionResponse = globalThis[functionName].apply(null, argsArray);
      if (functionResponse) {
        return functionResponse;
      }
      else {
        return "The function has been sucessfully executed but has nothing to return";
      }
    }
    else {
      throw Error("[GenAIApp] - Function not found or not a function: " + functionName);
    }
  }

  /**
   * Attempts to parse a JSON response string into an object. If the response contains errors, 
   * such as missing values, colons, or braces, it applies corrective measures to reconstruct 
   * and parse the JSON structure. Logs warnings if parsing fails after corrections.
   *
   * @private
   * @param {string} response - The JSON response string to parse, potentially with formatting issues.
   * @returns {Object|null} - The parsed JSON object if successful, or `null` if parsing fails.
   */
  function _parseResponse(response) {
    try {
      const parsedReponse = JSON.parse(response);
      return parsedReponse;
    }
    catch (e) {
      // Normalize to the substring between the first '{' and the last '}' to avoid fixed-index assumptions
      let text = String(response);
      const start = text.indexOf('{');
      const end = text.lastIndexOf('}');
      if (start === -1 || end === -1 || end < start) {
        console.warn('[GenAIApp] - Malformed JSON: missing object boundaries');
        return null;
      }
      text = text.slice(start, end + 1).trim();
      const lines = text.split('\n');
      for (let i = 1; i < lines.length - 1; i++) {
        let line = lines[i].trim();
        if (!line) continue; // skip empty lines
        if (!/"[^"]+"\s*:/.test(line)) continue; // not a key-value line
        // For other lines, check for missing values or colons
        line = line.trimEnd(); // Strip trailing white spaces
        if (line[line.length - 1] !== ',') {
          if (line[line.length - 1] == ':') {
            // If line has the format "property":, add null,
            lines[i] = line + ' ""';
          }
          else if (!line.includes(':')) {
            lines[i] = line + ': ""';
          }
          else if (line[line.length - 1] !== '"') {
            lines[i] = line + '"';
          }
        }
      }
      // Reconstruct the response
      response = lines.join('\n').trim();

      // Try parsing the corrected response
      try {
        const parsedResponse = JSON.parse(response);
        return parsedResponse;
      }
      catch (e) {
        // If parsing still fails, log the error and return null.
        console.warn('[GenAIApp] - Error parsing corrected response: ' + e.message);
        return null;
      }
    }
  }

  /**
   * Uploads a file to OpenAI and returns the file ID.
   * 
   * @param {string} optionalAttachment - The optional attachment ID from Google Drive.
   * @returns {string} The OpenAI file ID.
   */
  function _uploadFileToOpenAI(optionalAttachment) {
    const file = DriveApp.getFileById(optionalAttachment);
    const mimeType = file.getMimeType();
    let fileBlobUrl;

    switch (mimeType) {
      case "application/vnd.google-apps.spreadsheet":
        fileBlobUrl = 'https://docs.google.com/spreadsheets/d/' + optionalAttachment + '/export?format=xlsx';
        break;
      case "application/vnd.google-apps.document":
        fileBlobUrl = 'https://docs.google.com/document/d/' + optionalAttachment + '/export?format=docx';
        break;
      case "application/vnd.google-apps.presentation":
        fileBlobUrl = 'https://docs.google.com/presentation/d/' + optionalAttachment + '/export/pptx';
        break;
    }

    const token = ScriptApp.getOAuthToken();

    // Fetch the file from Google Drive using the generated URL and OAuth token
    let response = UrlFetchApp.fetch(fileBlobUrl, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    const fileBlob = response.getBlob();

    const openAIFileEndpoint = 'https://api.openai.com/v1/files';

    const formData = {
      'file': fileBlob,
      'purpose': 'assistants'
    };

    const uploadOptions = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + openAIKey
      },
      'payload': formData,
      'muteHttpExceptions': true
    };

    response = UrlFetchApp.fetch(openAIFileEndpoint, uploadOptions);
    const uploadedFileResponse = JSON.parse(response.getContentText());
    if (uploadedFileResponse.error) {
      throw new Error('[GenAIApp] - Error: ' + uploadedFileResponse.error.message);
    }
    return uploadedFileResponse.id;
  }

  /**
   * Fetches the content of a specified URL, logging the action if verbose mode is enabled.
   * Converts HTML content to Markdown format if the response is successful. If an error occurs
   * during fetching or access is denied, returns an error message.
   *
   * @private
   * @param {string} url - The URL of the webpage to fetch.
   * @returns {string|null} - The page content in Markdown format if successful, `null` if the response code is not 200, 
   *                          or an error message in JSON format if access is denied or an error occurs.
   */
  function _urlFetch(url) {
    if (verbose) {
      console.log(`[GenAIApp] - Clicked on link : ${url}`);
    }
    let response;
    try {
      response = UrlFetchApp.fetch(url);
    }
    catch (e) {
      console.warn(`[GenAIApp] - Error fetching the URL: ${e.message}`);
      return JSON.stringify({
        error: "Failed to fetch the URL : You are not authorized to access this website. Try another one."
      });
    }
    if (response.getResponseCode() == 200) {
      let pageContent = response.getContentText();
      pageContent = _convertHtmlToMarkdown(pageContent);
      return pageContent;
    }
    else {
      return null;
    }
  }

  /**
   * Converts an HTML string to Markdown format, removing unnecessary tags and attributes, 
   * and handling common HTML elements such as anchors, headings, lists, tables, images, 
   * inline code, and preformatted text. Also removes script and style content, and 
   * cleans up excess whitespace.
   *
   * @private
   * @param {string} htmlString - The HTML content to convert to Markdown.
   * @returns {string} - The converted Markdown representation of the HTML content.
   */
  function _convertHtmlToMarkdown(htmlString) {
    // Remove <script> tags and their content
    htmlString = htmlString.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
    // Remove <style> tags and their content
    htmlString = htmlString.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');
    // Remove on* attributes (e.g., onclick, onload)
    htmlString = htmlString.replace(/ ?on[a-z]*=".*?"/gi, '');
    // Remove style attributes
    htmlString = htmlString.replace(/ ?style=".*?"/gi, '');
    // Remove class attributes
    htmlString = htmlString.replace(/ ?class=".*?"/gi, '');

    // Convert &nbsp; to spaces
    htmlString = htmlString.replace(/&nbsp;/g, ' ');

    // Improved anchor tag conversion
    htmlString = htmlString.replace(/<a [^>]*href="(.*?)"[^>]*>(.*?)<\/a>/g, '[$2]($1)');

    // Strong
    htmlString = htmlString.replace(/<strong>(.*?)<\/strong>/g, '**$1**');
    // Emphasize
    htmlString = htmlString.replace(/<em>(.*?)<\/em>/g, '_$1_');
    // Headers
    for (let i = 1; i <= 6; i++) {
      const regex = new RegExp(`<h${i}>(.*?)<\/h${i}>`, 'g');
      htmlString = htmlString.replace(regex, `${'#'.repeat(i)} $1`);
    }

    // Blockquote
    htmlString = htmlString.replace(/<blockquote>(.*?)<\/blockquote>/g, '> $1');

    // Unordered list
    htmlString = htmlString.replace(/<ul>(.*?)<\/ul>/g, '$1');
    htmlString = htmlString.replace(/<li>(.*?)<\/li>/g, '- $1');

    // Ordered list
    htmlString = htmlString.replace(/<ol>(.*?)<\/ol>/g, '$1');
    htmlString = htmlString.replace(/<li>(.*?)<\/li>/g, (match, p1, offset, string) => {
      const number = string.substring(0, offset).match(/<li>/g).length;
      return `${number}. ${p1}`;
    });

    // Handle table headers (if they exist)
    htmlString = htmlString.replace(/<thead>([\s\S]*?)<\/thead>/g, (match, content) => {
      const headerRow = content.replace(/<th>(.*?)<\/th>/g, '| $1 ').trim() + '|';
      const separatorRow = headerRow.replace(/[^|]+/g, match => {
        return '-'.repeat(match.length);
      }).replace(/\| -/g, '|--');
      return headerRow + '\n' + separatorRow;
    });

    // Handle table rows
    htmlString = htmlString.replace(/<tbody>([\s\S]*?)<\/tbody>/g, (match, content) => {
      return content.replace(/<tr>([\s\S]*?)<\/tr>/g, (match, trContent) => {
        return trContent.replace(/<td>(.*?)<\/td>/g, '| $1 ').trim() + '|';
      });
    });

    // Inline code
    htmlString = htmlString.replace(/<code>(.*?)<\/code>/g, '`$1`');
    // Preformatted text
    htmlString = htmlString.replace(/<pre>(.*?)<\/pre>/g, '```\n$1\n```');

    // Images - Updated to use Markdown syntax
    htmlString = htmlString.replace(/<img src="(.+?)" alt="(.*?)" ?\/?>/g, '![$2]($1)'); // Markdown syntax for images is ![alt text](image URL).

    // Remove remaining HTML tags
    htmlString = htmlString.replace(/<[^>]*>/g, '');

    // Trim excessive white spaces between words/phrases
    htmlString = htmlString.replace(/ +/g, ' ');

    // Remove whitespace followed by newline patterns
    htmlString = htmlString.replace(/ \n/g, '\n');

    // Normalize the line endings to just \n
    htmlString = htmlString.replace(/\r\n/g, '\n');

    // Collapse multiple contiguous newline characters down to a single newline
    htmlString = htmlString.replace(/\n{2,}/g, '\n');

    // Trim leading and trailing white spaces and newlines
    htmlString = htmlString.trim();

    return htmlString;
  }

  /**
   * Makes the API call to Open AI to create a new vector store.
   * 
   * @param {string} vectorStoreName - The vectorStoreName to help build the vector store's name.
   * @returns {string} id - The id of the vector store that was just created.
   */
  function _createOpenAiVectorStore(vectorStoreName) {
    const url = apiBaseUrl + "/v1/vector_stores";

    const payload = {
      name: `VectorStore for ${vectorStoreName}`,
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + openAIKey,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result && result.id) {
      Logger.log({
        message: `[GenAIApp] - Vector store successfully created.`,
        id: result.id
      });

      const id = result.id;
      return id;
    }
    else {
      console.log(`[GenAIApp] - Failed to create vector store. Response: ${response.getContentText()}`);
      throw new Error("Fail to create vector store");
    }
  }

  /**
   * Retrieves information avout a specific Vector Store from Open AI's API.
   * 
   * @param {string} vectorStoreId - The Open AI API vector store Id.  
   */
  function _retrieveVectorStoreInformation(vectorStoreId) {
    const url = apiBaseUrl + '/v1/vector_stores/' + vectorStoreId;
    const options = {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + openAIKey,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      }
    };
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    console.log(`[GenAIApp] - Succesfully retrieved Vector Store information from Open AI : ${result}`);
    if (result.status == "completed") {
      return result.name;
    }
  }

  /**
   * Uploads a file to the Open AI storage.
   * 
   * @param {Blob} blob - The file blob.
   * @returns {string} id - The id of the uploaded file.
   */
  function _uploadFileToOpenAIStorage(blob) {
    const url = apiBaseUrl + "/v1/files";
    const headers = {
      'Authorization': 'Bearer ' + openAIKey
    };

    const form = {
      'purpose': "user_data",
      'file': blob
    };

    const options = {
      'method': "post",
      'headers': headers,
      'payload': form
    };

    try {
      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() == 200) {
        const json = JSON.parse(response.getContentText());
        Logger.log({
          message: `[GenAIApp] - File successfully uploaded to OpenAI`,
          id: json.id
        });
        const fileId = json.id;
        return fileId;
      }
      else {
        console.error(`[GenAIApp] - Unexpected error: ${response.getContentText()} (Status Code: ${response.getResponseCode()})`);
        throw new Error(`[GenAIApp] - Failed to upload file. Status Code: ${response.getResponseCode()}`);
      }
    }
    catch (error) {
      // Handle network errors or unexpected exceptions
      console.error(`[GenAIApp] - An error occurred while uploading the file to OpenAI: ${error.message}`);
      // Optionally, rethrow the error to be handled by the caller
      throw error;
    }
  }

  /**
   * Attaches a file to a specified vector store in OpenAI.
   *
   * @param {string} fileId - The unique identifier of the file to attach.
   * @param {string} vectorStoreId - The unique identifier of the vector store.
   * @param {Object} attributes - JSON object with the attributes of the file (set of up to 16 key-value pairs).
   * @returns {object} A vector store file object (JSON object).
   * @throws {Error} Throws an error if the attachment fails or if a network error occurs.
   */
  function _attachFileToVectorStore(fileId, vectorStoreId, attributes, max_chunk_size, chunk_overlap) {
    const url = apiBaseUrl + `/v1/vector_stores/${vectorStoreId}/files`;
    const payload = {
      "file_id": fileId,
      "attributes": attributes,
      "chunking_strategy": {
        "type": "static",
        "static": {
          "max_chunk_size_tokens": max_chunk_size,
          "chunk_overlap_tokens": chunk_overlap
        }
      }
    };

    const options = {
      method: 'post',
      'contentType': 'application/json',
      'headers': {
        'Authorization': 'Bearer ' + openAIKey,
        'OpenAI-Beta': 'assistants=v2'
      },
      'payload': JSON.stringify(payload)
    };
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    return data;
  }

  /**
   * Retrieves all file from a specified vector store in OpenAI.
   *
   * This function fetches files in batches of 100 using pagination. It continues to
   * request additional batches until all file have been retrieved. The file IDs are
   * stored as objects in an array.
   *
   * @param {string} vectorStoreId - The unique identifier of the vector store from which to list files.
   * @returns {Array} An array where each element is a file object from the vector store.
   * @throws {Error} Throws an error if there is an issue fetching the file IDs.
   */
  function _listFilesInVectorStore(vectorStoreId) {
    const baseUrl = apiBaseUrl + '/v1/vector_stores';
    const files = [];
    let hasMoreFiles = true;
    let after;

    while (hasMoreFiles) {
      try {
        // Get file IDs from the source vector store
        let url = `${baseUrl}/${vectorStoreId}/files?limit=100`;
        if (after) {
          url += `&after=${after}`;
        }

        const options = {
          'method': 'get',
          'headers': {
            'Authorization': 'Bearer ' + openAIKey,
            'OpenAI-Beta': 'assistants=v2'
          },
        };

        const response = UrlFetchApp.fetch(url, options);
        const storageData = JSON.parse(response.getContentText());

        if (storageData && storageData.data) {
          storageData.data.forEach(file => {
            files.push(file);
          });

          console.log(`[GenAIApp] - Fetched ${storageData.data.length} files`);

          if (storageData.data.length < 100) {
            hasMoreFiles = false;
          }
          else {
            after = storageData.data[storageData.data.length - 1].id;
          }
        }
        else {
          console.log('[GenAIApp] - No file IDs found in the vector store storage');
          hasMoreFiles = false;
        }
      }
      catch (e) {
        console.log(`[GenAIApp] - Error fetching files IDs: ${e.message}`);
        hasMoreFiles = false;
      }
    }

    return files;
  }

  /**
   * Deletes a file from a specified vector store in OpenAI.
   *
   * This function sends a DELETE request to the OpenAI API to remove a file from the specified vector store.
   * If an error occurs during the request, it is logged to the console.
   *
   * @param {string} vectorStoreId - The unique identifier of the vector store.
   * @param {string} fileId - The unique identifier of the file to delete.
   */
  function _deleteFileInVectorStore(vectorStoreId, fileId) {
    const url = apiBaseUrl + `/v1/vector_stores/${vectorStoreId}/files/${fileId}`;

    const options = {
      'method': 'delete',
      'headers': {
        'Authorization': 'Bearer ' + openAIKey,
        'OpenAI-Beta': 'assistants=v2'
      },
    };

    try {
      // Delete the file from the vector store
      UrlFetchApp.fetch(url, options);
    }
    catch (error) {
      console.error(`[GenAIApp] - Failed to delete file with ID: ${fileId}`, error);
    }
  }

  /**
   * Deletes a specific vector store from OpenAI by its ID.
   *
   * Sends a DELETE request to the OpenAI API to remove a specific vector store.
   *
   * @param {string} vectorStoreId - The unique identifier of the vector store to delete.
   * @returns {string} The unique identifier of the deleted vector store.
   * @throws {Error} Throws an error if the deletion fails or if there is an issue with the API request.
   */
  function _deleteVectorStore(vectorStoreId) {
    const url = apiBaseUrl + '/v1/vector_stores/' + vectorStoreId;

    const options = {
      method: 'delete',
      headers: {
        'Authorization': 'Bearer ' + openAIKey,
        'OpenAI-Beta': 'assistants=v2'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result && result.id) {
      Logger.log({
        message: `[GenAIApp] - Vector store successfully deleted.`,
        id: result.id
      });
      return result.id;
    }
    else {
      console.error(`[GenAIApp] - Failed to delete Vector store. Response: ${response.getContentText()}`);
      throw new Error("[GenAIApp] - Fail to delete Vector store");
    }
  }

  return {
    /**
     * Create a new chat.
     * @returns {Chat} - A new Chat instance.
     */
    newChat: function () {
      return new Chat();
    },

    /**
     * Create a new function.
     * @returns {FunctionObject} - A new Function instance.
     */
    newFunction: function () {
      return new FunctionObject();
    },

    /**
     * Create a new Vector Store.
     * @returns {VectorStoreObject} - A new Vector Store instance.
     */
    newVectorStore: function () {
      return new VectorStoreObject();
    },

    /**
     * Mandatory in order to use OpenAI models
     * @param {string} apiKey - Your openAI API key.
     */
    setOpenAIAPIKey: function (apiKey) {
      openAIKey = apiKey;
    },

    /**
     * To use Gemini models with an API key
     * @param {string} apiKey - Your Gemmini API key.
     */
    setGeminiAPIKey: function (apiKey) {
      geminiKey = apiKey;
    },

    /**
     * To use Gemini models without an API key
     * Requires Vertex AI enabled on a GCP project linked to your Google Apps Script project
     * @param {string} gcp_project_id - Your GCP project ID
     * @param {string} [gcp_project_region] - Your GCP project region (ex: us-central1, leave empty for global)
     */
    setGeminiAuth: function (gcp_project_id, gcp_project_region) {
      gcpProjectId = gcp_project_id;
      region = gcp_project_region;
    },

    /**
     * To set a global metadata key/value pair that will be passed along every message to the API.
     * @param {string} globalMetadataKey - The key of the key/value pair.
     * @param {string} globalMetadataValue - The value of the key/value pair.
     */
    setGlobalMetadata: function (globalMetadataKey, globalMetadataValue) {
      globalMetadata[globalMetadataKey] = globalMetadataValue;
    },

    /**
     * To set a specific API URL like Azure or Google Cloud for using Open AI models.
     * @param {string} baseUrl - The base url to be used for the API calls.
     */
    setPrivateInstanceBaseUrl: function (baseUrl) {
      privateInstanceBaseUrl = baseUrl;
    }
  }
})();
