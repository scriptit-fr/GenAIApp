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
      const mcpConnectors = [];
      let model = "gpt-5.6-terra"; // default
      //  OpenAI & Gemini models support a temperature value between 0.0 and 2.0. Models have a default temperature of 1.0.
      let temperature = 1;
      let max_tokens = 1000;
      let browsing = false;
      let reasoning_effort = "medium";
      let knowledgeLink = [];
      this._codeInterpreterEnabled = false;
      this._codeInterpreterContainerId = null;
      this._generatedFiles = [];
      this._lastContainerId = null;
      this._lastGeneratedDriveFileUrl = null;
      let compaction_enabled = false;
      let compaction_threshold = 10000;
      let tool_combination_enabled = false;

      let previous_response_id;
      let last_response_id = null;
      let previous_interaction_id;
      let last_gemini_interaction_id = null;
      let last_gemini_content_count = 0;

      let maxNumOfChunks = 10;
      let onlyChunks = false;
      let retrievedAttributes = [];

      const messageMetadata = {};
      let maximumAPICalls = 30;
      let numberOfAPICalls = 0;

      this._lastUsage = null;
      this._inputTokenWarningThreshold = null;

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
        else if (isBlobLike(imageInput)) {
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
        else if (isBlobLike(fileInput)) {
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
          Object.assign(
            contentObj,
            createOpenAIInputImageContent(fileInfo.mimeType, blobToBase64)
          );
        }
        else {
          Object.assign(
            contentObj,
            createOpenAIInputFileContent(fileInfo.mimeType, blobToBase64, fileInfo.fileName)
          );
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
       * Enables OpenAI Code Interpreter support for this chat.
       * @param {string} [containerId] - OPTIONAL - Explicit container ID to reuse.
       * @returns {Chat} - The current Chat instance.
       * @example
       * const chat = GenAIApp.newChat()
       *   .addFile(DriveApp.getFileById("YOUR_FILE_ID").getBlob())
       *   .enableCodeInterpreter()
       *   .addMessage("Process this file and generate an updated version.");
       * chat.run();
       * const generatedFiles = chat.getGeneratedFiles();
       * const blob = chat.downloadGeneratedFile(0);
       * DriveApp.createFile(blob);
       */
      this.enableCodeInterpreter = function (containerId) {
        this._codeInterpreterEnabled = true;
        if (containerId) {
          this._codeInterpreterContainerId = containerId;
        }
        return this;
      };

       /** OPTIONAL
       *
       * Enable or disable server-side tool invocations for Gemini (Tool Combination).
       * @param {boolean} enabled - True to enable tool combination.
       * @returns {Chat} - The current Chat instance.
       */
      this.enableToolCombination = function (enabled) {
        tool_combination_enabled = enabled;
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
        return last_response_id;
      };

      /**
       * Returns the last Gemini Interactions API interaction Id for this chat.
       */
      this.retrieveLastInteractionId = function () {
        return last_gemini_interaction_id;
      };

      /**
       * Defines the input token threshold that should trigger a warning log.
       * @param {number} input_token_threshold - Input token threshold for warning.
       * @returns {Chat} - The current Chat instance.
       */
      this.warnIfResponseTokenUsageAbove = function (input_token_threshold) {
        if (typeof input_token_threshold !== 'number' || !Number.isFinite(input_token_threshold) || input_token_threshold < 0) {
          throw new RangeError('[GenAIApp] - input token warning threshold must be a finite number >= 0.');
        }
        this._inputTokenWarningThreshold = input_token_threshold;
        return this;
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
       * Sets the previous Gemini Interactions API interaction Id used to continue a conversation.
       * @param {string} previousInteractionId - The id of the previous Gemini interaction.
       */
      this.setPreviousInteractionId = function (previousInteractionId) {
        previous_interaction_id = previousInteractionId;
        last_gemini_interaction_id = previousInteractionId || last_gemini_interaction_id;
        return this;
      };

      /**
       * Enable or disable server-side context compaction for OpenAI Responses API requests.
       * @param {boolean} enabled - True to enable compaction.
       */
      this.enableCompaction = function (enabled) {
        compaction_enabled = enabled;
        return this;
      };

      /**
       * Set the token threshold used by OpenAI server-side compaction.
       * @param {number} threshold - Token threshold that triggers compaction.
       */
      this.setCompactionThreshold = function (threshold) {
        if (typeof threshold !== 'number' || !Number.isFinite(threshold) || threshold < 1000) {
          throw new Error('[GenAIApp] - compaction threshold must be a number with minimum value 1000 (tokens).');
        }
        compaction_threshold = threshold;
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

      /**
       * Add a Google or custom MCP connector to the chat request.
       *
       * @param {ConnectorObject} connectorObject - The connector to be added.
       *
       * @returns {Chat} The current Chat instance (for chaining).
       */
      this.addMCP = function addMCP(connectorObject) {
        mcpConnectors.push(connectorObject);
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
          compaction_enabled: compaction_enabled,
          compaction_threshold: compaction_threshold,
          maximumAPICalls: maximumAPICalls,
          numberOfAPICalls: numberOfAPICalls,
          last_gemini_interaction_id: last_gemini_interaction_id
        };
      };

      /**
       * Start the chat conversation.
       * Sends all your messages and eventual function to chat GPT.
       * Will return the last chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @param {Object} [advancedParametersObject] OPTIONAL - For more advanced settings and specific usage only. {model, temperature, function_call}
       * @param {"gemini-3.1-pro-preview" | "gemini-3.1-flash-lite" | "gemini-3.5-flash" | "gpt-5.6-sol" | "gpt-5.6-terra" | "gpt-5.6-luna"} [advancedParametersObject.model]
       * @param {number} [advancedParametersObject.temperature]
       * @param {"low" | "medium" | "high"} [advancedParametersObject.reasoning_effort] Only needed for OpenAI reasoning models, defaults to medium
       * @param {number} [advancedParametersObject.max_tokens]
       * @param {string} [advancedParametersObject.function_call]
       * @returns {object} - the last message of the chat
       */
      this.run = function (advancedParametersObject) {
        this._lastUsage = null;
        last_response_id = null;
        this._generatedFiles = [];
        this._lastContainerId = null;
        this._lastGeneratedDriveFileUrl = null;

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

        if ((model.includes("gemini") || model.startsWith("gpt-5")) && browsing && max_tokens < 10000) {
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
              endpointUrl = `https://generativelanguage.googleapis.com/v1beta/interactions`;
            }
            else {
              // Enterprise endpoint / Vertex AI API
              // https://console.cloud.google.com/apis/api/aiplatform.googleapis.com
              // requires scope "https://www.googleapis.com/auth/cloud-platform.read-only" in access token
              if (!region || model.includes("gemini-3")) { // Gemini 3 requires global endpoint when using Vertex AI API
                endpointUrl = `https://aiplatform.googleapis.com/v1beta1/projects/${gcpProjectId}/locations/global/interactions`;
              }
              else {
                endpointUrl = `https://${region}-aiplatform.googleapis.com/v1beta1/projects/${gcpProjectId}/locations/${region}/interactions`;
              }
            }
          }
          responseMessage = _callGenAIApi(endpointUrl, payload);
          if (responseMessage?.usage) {
            this._lastUsage = responseMessage.usage;
            const inputTokens = this._lastUsage?.input_tokens ?? this._lastUsage?.total_input_tokens;
            if (this._inputTokenWarningThreshold !== null
              && inputTokens > this._inputTokenWarningThreshold) {
              console.warn(`[GenAIApp] - Warning: input token usage (${inputTokens}) exceeded configured threshold (${this._inputTokenWarningThreshold}) for response ${responseMessage.id}`);
            }
          }
          this._generatedFiles = this._extractContainerFileCitations(responseMessage);
          if (this._generatedFiles.length > 0) {
            this._lastContainerId = this._generatedFiles[0].containerId;
            const blob = this._downloadContainerFile(
              this._generatedFiles[0].containerId,
              this._generatedFiles[0].fileId,
              this._generatedFiles[0].filename
            );
            const createdFile = DriveApp.createFile(blob);
            this._lastGeneratedDriveFileUrl = createdFile.getUrl();
          }

          // OpenAI Responses API and Gemini Interactions API return a top-level "id".
          if (!model.includes("gemini")) {
            last_response_id = responseMessage?.id ?? null;
          }
          else {
            last_gemini_interaction_id = responseMessage?.id ?? last_gemini_interaction_id;
            previous_interaction_id = last_gemini_interaction_id || previous_interaction_id;
            last_gemini_content_count = contents.length;
          }
          numberOfAPICalls++;
        }
        else {
          throw new Error(`[GenAIApp] - Too many calls to genAI API: ${numberOfAPICalls}`);
        }
        if (Array.isArray(responseMessage.output)) {
          const fileSearchCall = responseMessage.output.filter(item => item.type === "file_search_call");
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
            const functionCalls = _extractGeminiFunctionCalls(responseMessage);
            if (functionCalls.length > 0) {
              contents = _handleGeminiToolCalls(responseMessage, tools, contents);
              // check if endWithResults or onlyReturnArguments
              if (contents._geminiEndWithResult) {
                if (verbose) {
                  console.log("[GenAIApp] - Conversation stopped because end function has been called");
                }
                return "OK";
              }
              else if (contents._geminiOnlyReturnArguments !== undefined) {
                if (verbose) {
                  console.log("[GenAIApp] - Conversation stopped because argument return has been enabled - No function has been called");
                }
                return contents._geminiOnlyReturnArguments;
              }
            }
            else {
              // if no function has been found, stop here
              return _extractGeminiResponseText(responseMessage);
            }
          }
          else {
            const functionCalls = responseMessage.output.filter(item => item.type === "function_call");
            if (functionCalls.length > 0) {
              messages = _handleOpenAIToolCalls(responseMessage.output, tools, messages);
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

              previous_response_id = responseMessage.id;
            }
            else {
              // if no function has been found, stop here
              if (this._lastGeneratedDriveFileUrl) {
                return this._lastGeneratedDriveFileUrl;
              }
              return _extractOpenAIResponseText(responseMessage);
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
            return _extractGeminiResponseText(responseMessage);
          }
          else {
            if (this._lastGeneratedDriveFileUrl) {
              return this._lastGeneratedDriveFileUrl;
            }
            return _extractOpenAIResponseText(responseMessage);
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
        let payload = {
          model: model,
          max_output_tokens: max_tokens,
          parallel_tool_calls: true,
          tools: []
        };
        if (model.startsWith("gpt-5")) {
          payload.reasoning = {
            "effort": reasoning_effort
          }
        }

        // Use the previous_response_id parameter to pass reasoning items from previous responses
        // This allows the model to continue its reasoning process to produce better results in the most token-efficient manner.
        // https://platform.openai.com/docs/guides/reasoning#keeping-reasoning-items-in-context
        if (previous_response_id) {
          payload.previous_response_id = previous_response_id;
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
        if (systemInstructions !== "") {
          payload.instructions = systemInstructions;
        }
        payload.input = userMessages;

        if (globalMetadata && Object.keys(globalMetadata).length > 0) {
          Object.assign(messageMetadata, globalMetadata);
          payload.metadata = messageMetadata;
        }

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

        if (mcpConnectors.length > 0) {
          // Parallel function calling is not possible when using built-in tools.
          payload.parallel_tool_calls = false;
          mcpConnectors.forEach(connector => {
            payload.tools.push(connector._toJson());
          });
        }

        if (compaction_enabled) {
          payload.context_management = [{
            type: "compaction",
            compact_threshold: compaction_threshold
          }];
        }

        if (browsing) {
          // Parallel function calling is not possible when using built-in tools.
          payload.parallel_tool_calls = false;
          payload.tools.push({
            type: "web_search"
          });
          if (restrictSearch) {
            messages.push({
              role: "user", // upon testing, this instruction has better results with user role instead of system
              content: `You are only able to search for information on ${restrictSearch}, restrict your search to this website only.`
            });
          }
        }

        if (Object.keys(addedVectorStores).length > 0 && numberOfAPICalls < 1) {
          // Parallel function calling is not possible when using built-in tools.
          payload.parallel_tool_calls = false;
          const fileSearchTool = {
            type: "file_search",
            vector_store_ids: Object.keys(addedVectorStores),
            max_num_results: maxNumOfChunks
          };
          payload.tools.push(fileSearchTool);
          payload.include = ["file_search_call.results"];
        }

        if (this._codeInterpreterEnabled) {
          payload.parallel_tool_calls = false;
          if (this._codeInterpreterContainerId) {
            payload.tools.push({
              type: "container",
              container_id: this._codeInterpreterContainerId
            });
          }
          else {
            payload.tools.push({
              type: "code_interpreter",
              container: {
                type: "auto"
              }
            });
          }
        }
        return payload;
      }

      this._extractContainerFileCitations = function (response) {
        if (!response || !Array.isArray(response.output)) {
          return [];
        }
        const citationsById = {};

        const addCitation = (containerId, fileId, filename) => {
          if (!containerId || !fileId) return;
          citationsById[fileId] = {
            containerId: containerId,
            fileId: fileId,
            filename: filename || citationsById[fileId]?.filename || null
          };
        };

        const walk = (node) => {
          if (!node) return;
          if (Array.isArray(node)) {
            node.forEach(walk);
            return;
          }
          if (typeof node !== "object") return;

          if (node.type === "container_file_citation") {
            addCitation(node.container_id, node.file_id, node.filename);
          }
          if (node.container_file_citation) {
            const c = node.container_file_citation;
            addCitation(c.container_id, c.file_id, c.filename);
          }
          if (node.type === "container_file" && node.file_id) {
            addCitation(node.container_id, node.file_id, node.filename);
          }

          Object.keys(node).forEach(key => walk(node[key]));
        };

        walk(response.output);
        return Object.keys(citationsById).map(fileId => citationsById[fileId]);
      };

      this._downloadContainerFile = function (containerId, fileId, filename) {
        const endpointUrl = `${apiBaseUrl}/v1/containers/${containerId}/files/${fileId}/content`;
        const response = _callGenAIApi(endpointUrl, null, "GET", true);
        const blob = response.getBlob();
        const contentType = response.getHeaders()["Content-Type"] || blob.getContentType();
        if (contentType) {
          blob.setContentType(contentType);
        }
        if (filename) {
          blob.setName(filename);
        }
        return blob;
      };

      /**
       * Returns generated files from the last run call.
       * @returns {{containerId: string, fileId: string, filename: string}[]} Generated files metadata.
       */
      this.getGeneratedFiles = function () {
        return this._generatedFiles;
      };

      /**
       * Downloads a generated file from the last run.
       * @param {string|number} fileIdOrIndex - File ID or index from getGeneratedFiles().
       * @returns {Blob} Downloaded file blob that can be stored with DriveApp.createFile(blob).
       * @example
       * const chat = GenAIApp.newChat()
       *   .addFile(DriveApp.getFileById("YOUR_FILE_ID").getBlob())
       *   .enableCodeInterpreter()
       *   .addMessage("Process this file and generate an updated version.");
       * chat.run();
       * const files = chat.getGeneratedFiles();
       * const blob = chat.downloadGeneratedFile(files[0].fileId);
       * DriveApp.createFile(blob);
       */
      this.downloadGeneratedFile = function (fileIdOrIndex) {
        let targetFile;
        if (fileIdOrIndex === undefined || fileIdOrIndex === null) {
          targetFile = this._generatedFiles[0];
        }
        else if (typeof fileIdOrIndex === "string" && fileIdOrIndex.trim() === "") {
          targetFile = this._generatedFiles[0];
        }
        if (typeof fileIdOrIndex === "number") {
          targetFile = this._generatedFiles[fileIdOrIndex];
        }
        else if (typeof fileIdOrIndex === "string") {
          targetFile = this._generatedFiles.find(file => file.fileId === fileIdOrIndex);
        }
        if (!targetFile) {
          throw new Error("[GenAIApp] - Generated file not found. Provide a valid file ID or index from getGeneratedFiles().");
        }
        return this._downloadContainerFile(targetFile.containerId, targetFile.fileId, targetFile.filename);
      };

      this.getContainerId = function () {
        if (this._lastContainerId) {
          return this._lastContainerId;
        }
        return this._generatedFiles[0]?.containerId || this._codeInterpreterContainerId || null;
      };

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
          model: model,
          input: _geminiContentsToInteractionInput(previous_interaction_id ? contents.slice(last_gemini_content_count) : contents),
          generation_config: {
            max_output_tokens: max_tokens,
            temperature: temperature
          },
          tools: []
        };

        // Continue Gemini conversations using the Interactions API state handle instead of resending
        // the full previous contents array.
        if (previous_interaction_id) {
          payload.previous_interaction_id = previous_interaction_id;
        }

        if (tool_combination_enabled) {
          payload.include_server_side_tool_invocations = true;
        }

        if (advancedParametersObject?.function_call) {
          // The Gemini Interactions API expects forced tool selection under
          // generation_config.tool_choice. Sending tool_choice at the top level
          // is rejected as an unknown parameter.
          payload.generation_config.tool_choice = {
            allowed_tools: {
              mode: "any",
              tools: Array.isArray(advancedParametersObject.function_call)
                ? advancedParametersObject.function_call
                : [advancedParametersObject.function_call]
            }
          };
          delete advancedParametersObject.function_call;
        }

        if (tools.length > 0) {
          // the user has added functions, enable function calling
          payload.tools = tools.map(t => {
            const toolFunction = t.function._toJson();

            const parameters = toolFunction.parameters;
            if (parameters?.type) {
              toolFunction.parameters.type = parameters.type.toUpperCase();
            }

            return {
              type: "function",
              name: toolFunction.name,
              description: toolFunction.description,
              parameters: toolFunction.parameters
            };
          });
        }

        if (Object.keys(addedVectorStores).length > 0 && numberOfAPICalls < 1) {
          payload.tools.push({
            "type": "file_search",
            file_search_store_names: Object.keys(addedVectorStores)
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
   * Internal provider for OpenAI vector stores.
   */
  class OpenAIVectorStoreProvider {
    constructor(state) {
      this.type = "openai";
      this.state = state;
    }

    create() {
      if (!this.state.name) throw new Error("[GenAIApp] - Please specify your Vector Store name using the GenAiApp.newVectorStore().setName() method before creating it.");
      return {
        id: _createOpenAiVectorStore(this.state.name),
        name: this.state.name
      };
    }

    retrieve(vectorStoreId) {
      return {
        id: vectorStoreId,
        name: _retrieveVectorStoreInformation(vectorStoreId)
      };
    }

    list() {
      return _listFilesInVectorStore(this.state.id);
    }

    upload(blob, attributes) {
      const uploadedFileId = _uploadFileToOpenAIStorage(blob);
      return _attachFileToVectorStore(uploadedFileId, this.state.id, attributes, this.state.max_chunk_size, this.state.chunk_overlap);
    }

    deleteItem(fileId) {
      return _deleteFileInVectorStore(this.state.id, fileId);
    }

    deleteStore() {
      return _deleteVectorStore(this.state.id);
    }
  }

  /**
   * @class
   * Internal provider for Gemini File Search Stores.
   */
  class GeminiFileSearchStoreProvider {
    constructor(state) {
      this.type = "gemini";
      this.state = state;
    }

    create() {
      const store = _createGeminiFileSearchStore(this.state.name, this.state.embeddingModel);
      return {
        id: store.name,
        name: store.displayName || this.state.name,
        embeddingModel: store.embeddingModel || this.state.embeddingModel
      };
    }

    retrieve(storeName) {
      const store = _retrieveGeminiFileSearchStoreInformation(storeName);
      return {
        id: store.name,
        name: store.displayName || "",
        embeddingModel: store.embeddingModel || ""
      };
    }

    list() {
      return _listGeminiFileSearchStoreDocuments(this.state.id);
    }

    upload(blob, attributes) {
      return _uploadAndImportGeminiFileSearchStoreDocument(this.state.id, blob, attributes);
    }

    deleteItem(documentName) {
      return _deleteGeminiFileSearchStoreDocument(documentName);
    }

    deleteStore() {
      throw new Error("[GenAIApp] - Deleting a Gemini File Search Store is not implemented in GenAIApp yet. Use deleteDocument/deleteFile to remove documents from the store.");
    }
  }

  /**
   * @class
   * Class representing a Vector Store, backed by an internal provider.
   */
  class VectorStoreObject {
    constructor(providerName = "openai") {
      const state = {
        name: "",
        description: "",
        id: null,
        max_chunk_size: 800,
        chunk_overlap: 400,
        embeddingModel: ""
      };
      const provider = providerName === "gemini"
        ? new GeminiFileSearchStoreProvider(state)
        : new OpenAIVectorStoreProvider(state);

      /**
       * Sets the vector store's name or display name.
       * @param {string} newName - The name to assign to the vector store.
       * @returns {VectorStoreObject}
       */
      this.setName = function (newName) {
        state.name = newName;
        return this;
      };

      /**
       * Sets the description of the vector store.
       * @param {string} newDesc - The description to assign to the vector store.
       * @returns {VectorStoreObject}
       */
      this.setDescription = function (newDesc) {
        state.description = newDesc;
        return this;
      };

      /**
       * Sets the chunking strategy for OpenAI vector-store uploads.
       * @param {number} maxChunkSize - The maximum token size of a chunk (max is 4096, defaults to 800).
       * @param {number} chunkOverlap - The chunk overlap to apply. Cannot exceed half of the maxChunkSize (defaults to 400).
       * @returns {VectorStoreObject}
       */
      this.setChunkingStrategy = function (maxChunkSize, chunkOverlap) {
        state.max_chunk_size = maxChunkSize;
        state.chunk_overlap = chunkOverlap;
        return this;
      };

      /**
       * Sets the embedding model for Gemini File Search Stores.
       * @param {string} embeddingModel - The embedding model resource name.
       * @returns {VectorStoreObject}
       */
      this.setEmbeddingModel = function (embeddingModel) {
        state.embeddingModel = embeddingModel;
        return this;
      };

      /**
       * Creates the provider-backed vector store.
       * @returns {VectorStoreObject}
       */
      this.createVectorStore = function () {
        try {
          const store = provider.create();
          state.id = store.id;
          state.name = store.name || state.name;
          state.embeddingModel = store.embeddingModel || state.embeddingModel;
        }
        catch (e) {
          console.error(`[GenAIApp] - Error creating the ${provider.type} vector store: ${e}`);
        }
        return this;
      };

      /**
       * Creates a Gemini File Search Store. Alias for createVectorStore().
       * @returns {VectorStoreObject}
       */
      this.createFileSearchStore = function () {
        return this.createVectorStore();
      };

      /**
       * Initializes a vector store object from an existing provider store id/resource name.
       * @param {string} vectorStoreId - The provider store id or resource name.
       * @returns {VectorStoreObject}
       */
      this.initializeFromId = function (vectorStoreId) {
        try {
          const store = provider.retrieve(vectorStoreId);
          state.id = store.id;
          state.name = store.name || state.name;
          state.embeddingModel = store.embeddingModel || state.embeddingModel;
        }
        catch (e) {
          console.error(`[GenAIApp] - Could not initialize ${provider.type} vector store object from id: ${e}`);
        }
        return this;
      };

      /**
       * Returns the vector store id/resource name.
       * @returns {string}
       */
      this.getId = function () {
        return state.id;
      };

      /**
       * Returns the vector store name/display name.
       * @returns {string}
       */
      this.getName = function () {
        return state.name;
      };

      /**
       * Uploads a file/blob and attaches or imports it into the vector store.
       * @param {Blob} blob - File to upload.
       * @param {Object} attributes - Metadata attributes to store on the item.
       * @returns {Object}
       */
      this.uploadAndAttachFile = function (blob, attributes = {}) {
        if (!state.id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object before attaching files.");
        try {
          return provider.upload(blob, attributes);
        }
        catch (e) {
          Logger.log({
            message: `Unable to upload and attach the file to the ${provider.type} vector store: ${e}`,
            fileBlob: blob
          });
          throw e;
        }
      };

      /**
       * Uploads and imports a document into the vector store. Alias for uploadAndAttachFile().
       * @param {Blob} blob - File to upload.
       * @param {Object} attributes - Metadata attributes to store on the item.
       * @returns {Object}
       */
      this.uploadAndImportDocument = function (blob, attributes = {}) {
        return this.uploadAndAttachFile(blob, attributes);
      };

      /**
       * Lists vector store files/documents.
       * @returns {Array}
       */
      this.listFiles = function () {
        if (!state.id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object before listing files.");
        try {
          return provider.list();
        }
        catch (e) {
          Logger.log({
            message: `An error occured when trying to list items from ${provider.type} vector store: ${e}`,
            vectorStoreId: state.id
          });
          throw e;
        }
      };

      /**
       * Lists vector store documents. Alias for listFiles().
       * @returns {Array}
       */
      this.listDocuments = function () {
        return this.listFiles();
      };

      /**
       * Deletes a file/document from the vector store.
       * @param {string} itemId - The provider item id/resource name to delete.
       * @returns {Object}
       */
      this.deleteFile = function (itemId) {
        if (!itemId) throw new Error("[GenAIApp] - Please pass a vector store item ID to deleteFile(itemId).");
        if (!state.id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object before deleting files.");
        try {
          return provider.deleteItem(itemId);
        }
        catch (e) {
          Logger.log({
            message: `An error occured when trying to delete the item from ${provider.type} vector store: ${e}`,
            vectorStoreId: state.id,
            itemId: itemId
          });
          throw e;
        }
      };

      /**
       * Deletes a document from the vector store. Alias for deleteFile().
       * @param {string} documentId - The provider document id/resource name to delete.
       * @returns {Object}
       */
      this.deleteDocument = function (documentId) {
        return this.deleteFile(documentId);
      };

      /**
       * Deletes the vector store when supported by the provider.
       * @returns {string|Object}
       */
      this.deleteVectorStore = function () {
        if (!state.id) throw new Error("[GenAIApp] - Please create or initialize your Vector Store object before being deleted.");
        try {
          const deleteId = provider.deleteStore();
          state.id = null;
          return deleteId;
        }
        catch (e) {
          Logger.log({
            message: `An error occured when trying to delete the ${provider.type} vector store: ${e}`,
            vectorStoreId: state.id
          });
          throw e;
        }
      };

      /**
       * Returns the JSON object with provider, name, description, and ID.
       * @returns {Object}
       */
      this._toJson = function () {
        return {
          provider: provider.type,
          name: state.name,
          description: state.description,
          id: state.id,
          embeddingModel: state.embeddingModel
        };
      };
    }
  }

  /**
   * @class
   * Class representing an MCP Connector.
   */
  class ConnectorObject {
    constructor() {
      let serverLabel = "";
      let serverDescription = null;
      let serverUrl = null;
      let connectorId = null;
      let allowedTools = null;
      let authorization = ScriptApp.getOAuthToken();
      let requireApproval = "never";

      /**
       * Sets the label used to identify the connector.
       * @param {string} label - The label to assign to the connector.
       * @returns {ConnectorObject}
       */
      this.setLabel = function (label) {
        if (typeof label !== "string" || label.trim() === "") {
          throw Error("[GenAIApp] - Please provide a non-empty server label.");
        }
        serverLabel = label.trim().replace(/[^a-zA-Z0-9_-]/g, "_");
        return this;
      };

      /**
       * Sets the optional description for the connector.
       * @param {string} description - The description to assign.
       * @returns {ConnectorObject}
       */
      this.setDescription = function (description) {
        if (typeof description !== "string") {
          throw Error("[GenAIApp] - The server description must be a string.");
        }
        serverDescription = description;
        return this;
      };

      /**
       * Configures the connector to use a custom MCP server URL.
       * @param {string} url - The HTTPS URL of the MCP server.
       * @returns {ConnectorObject}
       */
      this.setServerUrl = function (url) {
        if (typeof url !== "string" || url.trim() === "") {
          throw Error("[GenAIApp] - Please provide a non-empty server URL.");
        }
        const trimmedUrl = url.trim();
        if (!trimmedUrl.toLowerCase().startsWith("https://")) {
          throw Error("[GenAIApp] - Invalid server URL");
        }
        serverUrl = trimmedUrl;
        return this;
      };

      /**
       * Configures the connector to use one of the predefined Google connectors (Gmail, Calendar, Drive).
       * @param {"gmail"|"calendar"|"drive"} connectorType - The Google connector identifier.
       * @returns {ConnectorObject}
       */
      this.setConnectorId = function (connectorType) {
        if (typeof connectorType !== "string" || connectorType.trim() === "") {
          throw Error("[GenAIApp] - Please specify the Google MCP connector you want to use: 'gmail', 'calendar' or 'drive'");
        }

        const normalizedType = connectorType.toLowerCase();
        const validTypes = ["gmail", "calendar", "drive"];
        if (!validTypes.includes(normalizedType)) {
          throw Error(`[GenAIApp] - Invalid Google connector type: ${connectorType}. Accepted types are 'gmail', 'calendar' and 'drive'`);
        }

        const normalizedConnector = connectorType.toLowerCase().trim();
        const connectorIds = {
          gmail: "connector_gmail",
          calendar: "connector_googlecalendar",
          drive: "connector_googledrive",
        };
        const serverLabels = {
          gmail: "gmail",
          calendar: "google_calendar",
          drive: "google_drive"
        };

        if (!connectorIds[normalizedConnector]) {
          throw Error("[GenAIApp] - Unsupported Google MCP connector provided.");
        }

        connectorId = connectorIds[normalizedConnector];
        serverLabel = serverLabels[normalizedConnector];
        return this;
      };

      /**
       * Sets the authorization token for the connector.
       * @param {string|null} token - The access token used to authorize the connector.
       * @returns {ConnectorObject}
       */
      this.setAuthorization = function (token) {
        if (token === null || token === undefined) {
          authorization = null;
          return this;
        }

        if (typeof token !== "string" || token.trim().length < 10) {
          throw Error("[GenAIApp] - Invalid authorization token provided.");
        }

        authorization = token.trim();
        return this;
      };

      /**
       * Sets the approval requirement for the connector.
       * @param {"never"|"domain"|"always"} approval - The approval requirement.
       * @returns {ConnectorObject}
       */
      this.setRequireApproval = function (approval) {
        const allowedValues = ["never", "domain", "always"];
        if (approval === undefined || approval === null) {
          requireApproval = "never";
          return this;
        }

        if (allowedValues.indexOf(approval) === -1) {
          throw Error("[GenAIApp] - Invalid requireApproval value. Use 'never', 'domain', or 'always'.");
        }

        requireApproval = approval;
        return this;
      };

      /**
       * Sets the allowed mcp tools for the connector.
       * @param {Array<string>} allowedToolsArray - Allowed mcp tools to be called.
       * @returns {ConnectorObject}
       */
      this.setAllowedTools = function (allowedToolsArray) {
        if (!Array.isArray(allowedToolsArray)) {
          throw Error("[GenAIApp] - allowedTools must be an array.");
        }

        if (allowedToolsArray.length === 0) {
          throw Error("[GenAIApp] - allowedTools array cannot be empty.");
        }

        if (!allowedToolsArray.every(tool => typeof tool === "string" && tool.trim() !== "")) {
          throw Error("[GenAIApp] - All items in allowedTools must be non-empty strings.");
        }

        allowedTools = allowedToolsArray.map(tool => tool.trim());
        return this;
      };

      /**
       * Returns the JSON representation for the connector.
       * @returns {Object}
       */
      this._toJson = function () {
        if (!serverUrl && !connectorId) {
          throw Error("[GenAIApp] - Please configure the connector using useServerUrl() or setConnectorId().");
        }

        const connector = {
          type: "mcp",
          require_approval: requireApproval
        };

        if (serverDescription) {
          connector.server_description = serverDescription;
        }

        if (serverUrl) {
          connector.server_url = serverUrl;
          connector.server_label = serverLabel || "custom_mcp";
        }
        else {
          connector.connector_id = connectorId;
          connector.server_label = serverLabel || connectorId;
        }

        if (allowedTools) {
          connector.allowed_tools = allowedTools;
        }

        if (authorization) {
          connector.authorization = authorization;
        }

        return connector;
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
  function _callGenAIApi(endpoint, payload, method = "post", returnRawResponse = false) {
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

    const hasMcpConnectors = !!payload && Array.isArray(payload.tools) && payload.tools.some(t => t && t.type === "mcp");

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
        method: method.toLowerCase(),
        headers: headers,
        timeoutSeconds: 30 * 60,
        muteHttpExceptions: true
      };
      if (payload !== null && payload !== undefined) {
        options.payload = JSON.stringify(payload);
      }

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
        if (returnRawResponse) {
          return response;
        }
        // The request was successful, exit the loop.
        const parsedResponse = JSON.parse(response.getContentText());
        if (endpoint.includes("google")) {
          responseMessage = parsedResponse;
          finish_reason = parsedResponse.status
            || parsedResponse.finishReason
            || parsedResponse.finish_reason
            || (_extractGeminiFunctionCalls(parsedResponse).length > 0 ? "tool_calls" : "completed");
        }
        else {
          responseMessage = parsedResponse;
          finish_reason = parsedResponse.status;
        }
        if (finish_reason == "length" || finish_reason == "incomplete" || finish_reason == "MAX_TOKENS") {
          console.warn(`[GenAIApp] - ${payload.model} response could not be completed because of an insufficient amount of tokens. To resolve this issue, you can increase the amount of tokens like this : chat.run({max_tokens: XXXX}).`);
        }
        success = true;
      }
      else if (responseCode === 400 && hasMcpConnectors) {
        // Retry on context_length_exceeded ONLY when MCP connectors are present.
        let errJson = null;
        try { errJson = JSON.parse(response.getContentText()); } catch (e) { }
        const errCode = errJson?.error?.code;
        if (errCode === "context_length_exceeded") {
          // No need to wait before retrying
          retries++;
          console.warn(`[GenAIApp] - Context length exceeded when calling ${payload.model} with MCP connectors, retrying (${retries}/${maxRetries}).`);
          // No payload changes, no shrinking: exact same request again.
          continue;
        }
        // If it's a different 400, fall through to the generic error handler below.
        console.error(`[GenAIApp] - Request to ${payload.model} failed with response code ${responseCode} - ${response.getContentText()}`);
        break;
      }
      else if (responseCode === 429) {
        // Rate limit reached, wait before retrying.
        const delay = Math.pow(2, retries) * 1000; // Delay in milliseconds, starting at 1 second.
        Utilities.sleep(delay);
        retries++;
        console.warn(`[GenAIApp] - Rate limit reached when calling ${payload.model}, retrying (${retries}/${maxRetries}).`);
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
  function _geminiContentsToInteractionInput(contents) {
    return (contents || []).flatMap(content => {
      const textStep = {
        type: content.role === "model" ? "model_output" : "user_input",
        content: []
      };
      const resultSteps = [];
      const parts = Array.isArray(content.parts) ? content.parts : [content.parts];
      parts.forEach(part => {
        if (!part) return;
        if (part.text) {
          textStep.content.push({ type: "text", text: part.text });
        }
        else if (part.inline_data) {
          textStep.content.push({
            type: part.inline_data.mime_type?.startsWith("image/") ? "image" : "document",
            mime_type: part.inline_data.mime_type,
            data: part.inline_data.data
          });
        }
        else if (part.functionResponse) {
          resultSteps.push({
            type: "function_result",
            call_id: part.functionResponse.call_id,
            name: part.functionResponse.name,
            result: [{ type: "text", text: part.functionResponse.response?.functionResponse ?? "" }]
          });
        }
      });
      return textStep.content.length > 0 ? [textStep, ...resultSteps] : resultSteps;
    });
  }

  function _extractGeminiResponseText(responseMessage) {
    const steps = responseMessage?.steps || [];
    const texts = [];
    steps.forEach(step => {
      if (step?.type !== "model_output") return;
      const content = Array.isArray(step.content) ? step.content : [step.content];
      content.forEach(item => {
        if (!item) return;
        if (typeof item === "string") texts.push(item);
        else if (item.text) texts.push(item.text);
        else if (item.type === "text" && item.content) texts.push(item.content);
        else if (item.type === "output_text" && item.text) texts.push(item.text);
      });
      if (step.text) texts.push(step.text);
      if (step.output_text) texts.push(step.output_text);
    });
    if (texts.length > 0) return texts.join("\n");

    // Backward-compatible fallback for legacy content-shaped responses.
    const parts = responseMessage?.parts || [];
    const part = parts.find(p => !p.thought && p.text);
    return part?.text || responseMessage?.output_text || null;
  }

  function _extractGeminiFunctionCalls(responseMessage) {
    const calls = [];
    const steps = responseMessage?.steps || [];
    steps.forEach(step => {
      if (step?.type !== "function_call") return;
      calls.push({
        id: step.id || step.call_id || step.function_call_id,
        name: step.name || step.function?.name || step.functionCall?.name,
        args: step.args || step.arguments || step.function?.arguments || step.functionCall?.args || {}
      });
    });
    if (calls.length > 0) return calls;

    // Backward-compatible fallback for legacy content-shaped responses.
    const parts = responseMessage?.parts || [];
    parts.forEach(part => {
      if (part?.functionCall?.name) {
        calls.push({
          id: part.functionCall.id,
          name: part.functionCall.name,
          args: part.functionCall.args || {}
        });
      }
    });
    return calls;
  }

  /**
   * Processes tool calls from a Gemini Interactions API response message.
   *
   * @private
   * @param {Object} responseMessage - The response message from Gemini containing tool-call steps.
   * @param {Array} tools - List of available tools, each with metadata including function names and argument requirements.
   * @param {Array} contents - Array representing the conversational content, updated for backward compatibility.
   * @returns {Object} - Tool continuation input and state flags for the Chat run loop.
   */
  function _handleGeminiToolCalls(responseMessage, tools, contents) {
    const functionCalls = _extractGeminiFunctionCalls(responseMessage);
    const functionResults = [];
    let shouldEndWithResult = false;
    let onlyReturnArguments = null;

    functionCalls.forEach(functionCall => {
      const functionName = functionCall.name;
      const functionArgs = functionCall.args || {};
      if (!functionName) return;

      let argsOrder = [];
      let endWithResult = false;
      let onlyArgs = false;

      for (const t in tools) {
        const currentFunction = tools[t].function._toJson();
        if (currentFunction.name == functionName) {
          argsOrder = currentFunction.argumentsInRightOrder;
          endWithResult = currentFunction.endingFunction;
          onlyArgs = currentFunction.onlyArgs;
          break;
        }
      }

      // No actual call to the function
      if (onlyArgs) {
        onlyReturnArguments = functionArgs;
        return;
      }

      if (endWithResult) {
        shouldEndWithResult = true;
      }

      let functionResponse = _callFunction(functionName, functionArgs, argsOrder);
      if (verbose) {
        console.log(`[GenAIApp] - function ${functionName}() called by Gemini.`);
      }
      if (typeof functionResponse != "string") {
        functionResponse = typeof functionResponse == "object" ? JSON.stringify(functionResponse) : String(functionResponse);
      }

      functionResults.push({
        call_id: functionCall.id,
        name: functionName,
        response: { functionResponse }
      });
    });

    if (functionResults.length > 0) {
      contents.push({
        role: 'user',
        parts: functionResults.map(result => ({
          functionResponse: result
        }))
      });
    }

    contents._geminiEndWithResult = shouldEndWithResult;
    delete contents._geminiOnlyReturnArguments;
    if (onlyReturnArguments !== null) {
      contents._geminiOnlyReturnArguments = onlyReturnArguments;
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

            // handle array of blobs (or mixed arrays)
            if (Array.isArray(functionResponse)) {
              if (functionResponse.length > 0 && functionResponse.every(isBlobLike)) {
                functionResponse = functionResponse.map(blobToResponseInputFileContent);
              } else {
                // non-blob arrays
                functionResponse = JSON.stringify(functionResponse);
              }
            }
            // single-object handling
            else {
              // check if response is a blob
              if (isBlobLike(functionResponse)) {
                functionResponse = [blobToResponseInputFileContent(functionResponse)];
              }
              else {
                functionResponse = JSON.stringify(functionResponse);
              }
            }
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
   * Extracts assistant text from OpenAI Responses API output.
   * Prioritizes messages marked as `final_answer` over intermediate `commentary`.
   * Falls back to the last available assistant message text when no explicit final answer exists.
   * 
   * Logs a warning when compaction was used for the response.
   *
   * @private
   * @param {Array} response - The `response` array returned by OpenAI.
   * @returns {string|null} - The selected assistant text, or `null` if no text is found.
   */
  function _extractOpenAIResponseText(response) {
    const output = response?.output;

    if (!Array.isArray(output)) {
      return null;
    }

    const compactionItems = output.filter(item => item?.type === "compaction");
    if (compactionItems.length > 0) {
      compactionItems.forEach(item => {
        console.warn(`[GenAIApp] Compaction was used for response ${response?.id ?? null}`);
      });
    }

    const messageItems = output.filter(item => item?.type === "message");
    if (messageItems.length === 0) {
      return null;
    }

    const getText = (messageItem) => {
      const textPart = messageItem?.content?.find(part => part?.text);
      return textPart?.text || null;
    };

    const finalAnswerMessage = messageItems.find(item => item?.status === "final_answer");
    if (finalAnswerMessage) {
      return getText(finalAnswerMessage);
    }

    return getText(messageItems[messageItems.length - 1]);
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

  // Returns true for Blob-like objects exposing Apps Script Blob methods.
  const isBlobLike = (x) =>
    x &&
    typeof x === "object" &&
    typeof x.getBytes === "function" &&
    typeof x.getContentType === "function";

  // OpenAI-only helper: creates a Responses API input_file content object.
  const createOpenAIInputFileContent = (mimeType, base64Data, filename) => ({
    type: "input_file",
    file_data: `data:${mimeType};base64,${base64Data}`,
    filename: filename
  });

  // OpenAI-only helper: creates a Responses API input_image content object.
  const createOpenAIInputImageContent = (mimeType, base64Data) => ({
    type: "input_image",
    image_url: `data:${mimeType};base64,${base64Data}`
  });

  // OpenAI-only helper for Blob-like values returned by function calling.
  const blobToResponseInputFileContent = (blob) => {
    const mimeType = blob.getContentType();
    const base64Data = Utilities.base64Encode(blob.getBytes());

    if (mimeType && mimeType.startsWith("image/")) {
      return createOpenAIInputImageContent(mimeType, base64Data);
    }

    return createOpenAIInputFileContent(
      mimeType,
      base64Data,
      blob.getName()
    );
  };

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


  /**
   * Builds Gemini REST headers using the same authentication branch as _callGenAIApi().
   * @param {Object} extraHeaders - Additional headers to include.
   * @returns {Object}
   */
  function _getGeminiRestHeaders(extraHeaders = {}) {
    const headers = Object.assign({}, extraHeaders);
    if (geminiKey) {
      headers['x-goog-api-key'] = geminiKey;
    }
    else {
      headers['Authorization'] = 'Bearer ' + ScriptApp.getOAuthToken();
    }
    return headers;
  }

  /**
   * Calls Gemini REST APIs with JSON parsing and status-code error handling.
   * @param {string} url - Endpoint URL.
   * @param {Object} options - UrlFetch options.
   * @returns {Object}
   */
  function _callGeminiFileSearchStoreApi(url, options) {
    options.muteHttpExceptions = true;
    let response;
    if (typeof ErrorHandler !== 'undefined' && typeof ErrorHandler.urlFetchWithExpBackOff === 'function') {
      response = ErrorHandler.urlFetchWithExpBackOff(url, options);
    }
    else {
      response = UrlFetchApp.fetch(url, options);
    }

    const responseText = response.getContentText();
    const responseCode = response.getResponseCode();
    let result = {};
    if (responseText) {
      result = JSON.parse(responseText);
    }

    if (responseCode >= 200 && responseCode < 300) {
      return result;
    }

    const errorMessage = result && result.error && result.error.message ? result.error.message : responseText;
    throw new Error(`[GenAIApp] - Gemini File Search Store API error (${responseCode}): ${errorMessage}`);
  }

  /**
   * Creates a Gemini File Search Store.
   * @param {string} displayName - Human-readable display name.
   * @param {string} embeddingModel - Optional embedding model resource name.
   * @returns {Object}
   */
  function _createGeminiFileSearchStore(displayName, embeddingModel) {
    const payload = {};
    if (displayName) payload.displayName = displayName;
    if (embeddingModel) payload.embeddingModel = embeddingModel;

    const url = 'https://generativelanguage.googleapis.com/v1beta/fileSearchStores';
    const store = _callGeminiFileSearchStoreApi(url, {
      method: 'post',
      contentType: 'application/json',
      headers: _getGeminiRestHeaders({ 'Content-Type': 'application/json' }),
      payload: JSON.stringify(payload)
    });

    if (!store || !store.name) {
      throw new Error('[GenAIApp] - Gemini File Search Store creation did not return a store name.');
    }

    return store;
  }

  /**
   * Retrieves Gemini File Search Store information by resource name.
   * @param {string} storeName - Gemini File Search Store resource name.
   * @returns {Object}
   */
  function _retrieveGeminiFileSearchStoreInformation(storeName) {
    const url = `https://generativelanguage.googleapis.com/v1beta/${storeName}`;
    return _callGeminiFileSearchStoreApi(url, {
      method: 'get',
      headers: _getGeminiRestHeaders()
    });
  }

  /**
   * Normalizes Gemini customMetadata into an object.
   * @param {Array|Object} customMetadata - Gemini custom metadata.
   * @returns {Object}
   */
  function _normalizeGeminiCustomMetadata(customMetadata) {
    if (!customMetadata) return {};
    if (!Array.isArray(customMetadata)) return customMetadata;
    return customMetadata.reduce((metadata, entry) => {
      if (entry && entry.key) {
        metadata[entry.key] = entry.stringValue != null
          ? entry.stringValue
          : entry.value != null
            ? entry.value
            : entry.numericValue != null
              ? entry.numericValue
              : entry.boolValue;
      }
      return metadata;
    }, {});
  }

  /**
   * Converts a plain metadata object to Gemini customMetadata entries.
   * @param {Object} attributes - Metadata attributes.
   * @returns {Array}
   */
  function _buildGeminiCustomMetadata(attributes = {}) {
    return Object.keys(attributes)
      .filter(key => attributes[key] !== undefined && attributes[key] !== null)
      .map(key => ({ key: key, stringValue: String(attributes[key]) }));
  }

  /**
   * Lists documents in a Gemini File Search Store, following pagination.
   * @param {string} storeName - Gemini File Search Store resource name.
   * @returns {Array}
   */
  function _listGeminiFileSearchStoreDocuments(storeName) {
    const documents = [];
    let pageToken = null;

    do {
      let url = `https://generativelanguage.googleapis.com/v1beta/${storeName}/documents?pageSize=20`;
      if (pageToken) url += `&pageToken=${encodeURIComponent(pageToken)}`;

      const result = _callGeminiFileSearchStoreApi(url, {
        method: 'get',
        headers: _getGeminiRestHeaders()
      });

      (result.documents || []).forEach(document => {
        const customMetadata = _normalizeGeminiCustomMetadata(document.customMetadata);
        documents.push(Object.assign({}, document, {
          id: document.name,
          name: document.name,
          displayName: document.displayName,
          customMetadata: customMetadata,
          attributes: { url: customMetadata.url || customMetadata.publicUrl }
        }));
      });
      pageToken = result.nextPageToken;
    } while (pageToken);

    return documents;
  }

  /**
   * Uploads a blob directly into a File Search Store and waits for completion.
   *
   * Prefer the uploadToFileSearchStore media endpoint over the files.upload + importFile
   * two-step flow because importFile can reject otherwise valid Gemini API-key requests.
   * The direct upload endpoint still returns a File Search Store operation, so callers keep
   * the same completion/error semantics.
   *
   * @param {string} storeName - Gemini File Search Store resource name.
   * @param {Blob} blob - File blob.
   * @param {Object} attributes - Metadata to store on the document.
   * @returns {Object}
   */
  function _uploadAndImportGeminiFileSearchStoreDocument(storeName, blob, attributes = {}) {
    const operation = _uploadToGeminiFileSearchStore(storeName, blob, attributes);
    return _pollGeminiFileSearchStoreOperation(operation.name);
  }

  /**
   * Uploads raw media directly to a Gemini File Search Store.
   * @param {string} storeName - Gemini File Search Store resource name.
   * @param {Blob} blob - File blob.
   * @param {Object} attributes - Metadata to store on the document.
   * @returns {Object}
   */
  function _uploadToGeminiFileSearchStore(storeName, blob, attributes = {}) {
    const url = `https://generativelanguage.googleapis.com/upload/v1beta/${storeName}:uploadToFileSearchStore`;
    const fileName = blob.getName ? blob.getName() : 'document.html';
    const mimeType = blob.getContentType ? blob.getContentType() : 'text/html';

    return _callGeminiFileSearchStoreApi(url, {
      method: 'post',
      headers: _getGeminiRestHeaders(),
      payload: {
        metadata: Utilities.newBlob(JSON.stringify({
          displayName: fileName,
          mimeType: mimeType,
          customMetadata: _buildGeminiCustomMetadata(attributes)
        }), 'application/json'),
        file: blob
      }
    });
  }

  /**
   * Polls a Gemini File Search Store operation until it is done.
   * @param {string} operationName - Operation resource name.
   * @returns {Object}
   */
  function _pollGeminiFileSearchStoreOperation(operationName) {
    if (!operationName) throw new Error('[GenAIApp] - Gemini File Search Store upload did not return an operation name.');
    const operationUrl = `https://generativelanguage.googleapis.com/v1beta/${operationName}`;
    const maxPolls = 30;
    const pollIntervalMs = 2000;

    for (let poll = 0; poll < maxPolls; poll++) {
      const operation = _callGeminiFileSearchStoreApi(operationUrl, {
        method: 'get',
        headers: _getGeminiRestHeaders()
      });

      if (operation.done) {
        if (operation.error) {
          throw new Error(`[GenAIApp] - Gemini File Search Store upload failed: ${operation.error.message || JSON.stringify(operation.error)}`);
        }
        return operation.response || operation;
      }

      Logger.log({
        message: `[GenAIApp] - Waiting for Gemini File Search Store upload operation ${operationName} (${poll + 1}/${maxPolls})`
      });
      Utilities.sleep(pollIntervalMs);
    }

    throw new Error(`[GenAIApp] - Gemini File Search Store upload operation timed out: ${operationName}`);
  }

  /**
   * Deletes a Gemini File Search Store document using force=true.
   * @param {string} documentName - Document resource name.
   * @returns {Object}
   */
  function _deleteGeminiFileSearchStoreDocument(documentName) {
    const url = `https://generativelanguage.googleapis.com/v1beta/${documentName}?force=true`;
    return _callGeminiFileSearchStoreApi(url, {
      method: 'delete',
      headers: _getGeminiRestHeaders()
    });
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
     * Create a new connector.
     * @returns {ConnectorObject} - A new Connector instance.
     */
    newConnector: function () {
      return new ConnectorObject();
    },

    /**
     * Create a new provider-backed Vector Store.
     * @param {"openai"|"gemini"} [providerName] - The vector store provider, defaults to OpenAI.
     * @returns {VectorStoreObject} - A new Vector Store instance.
     */
    newVectorStore: function (providerName) {
      return new VectorStoreObject(providerName || "openai");
    },

    /**
     * Create a new Gemini File Search Store wrapper.
     * @returns {VectorStoreObject} - A new Gemini File Search Store instance.
     */
    newGeminiFileSearchStore: function () {
      return new VectorStoreObject("gemini");
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