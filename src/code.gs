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

    let googleCustomSearchAPIKey = "";
    let restrictSearch;

    let verbose = true;

    const noResultFromWebSearchMessage = `Your search did not match any documents. 
  Try with different, more general or fewer keywords.`;

    /**
     * @class
     * Class representing a function known by function calling model
     */
    class FunctionObject {

        constructor() {
            let name = '';
            let description = '';
            let properties = {};
            let required = [];
            let argumentsInRightOrder = [];
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
                    let startIndex = type.indexOf("<") + 1;
                    let endIndex = type.indexOf(">");
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
                        throw Error("Please precise the type of the items contained in the array when calling addParameter. Use format Array.<itemsType> for the type parameter.");
                        return
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

    let webSearchFunction = new FunctionObject()
        .setName("webSearch")
        .setDescription("Perform a web search via the Google Custom Search JSON API. Returns an array of search results (including the URL, title and plain text snippet for each result)")
        .addParameter("q", "string", "the query for the web search.");

    let urlFetchFunction = new FunctionObject()
        .setName("urlFetch")
        .setDescription("Fetch the viewable content of a web page. HTML tags will be stripped, returning a text-only version.")
        .addParameter("url", "string", "The URL to fetch.");

    let imageDescriptionFunction = new FunctionObject()
        .setName("getImageDescription")
        .setDescription("To retrieve the description of an image.")
        .addParameter("imageUrl", "string", "The URL of the image.")
        .addParameter("highFidelity", "boolean", `Default: false. To improve the image quality, not needed in most cases.`, isOptional = true);

    /**
     * @class
     * Class representing a chat.
     */
    class Chat {
        constructor() {
            let messages = []; // messages for OpenAI API
            let contents = []; // contents for Gemini API
            let tools = [];
            let model = "gpt-3.5-turbo"; // default 
            let temperature = 0.5;
            let max_tokens = 300;
            let browsing = false;
            let vision = false;
            let onlyRetrieveSearchResults = false;
            let knowledgeLink;
            let assistantIdentificator;
            let vectorStore;
            let attachmentIdentificator;
            let assistantTools;

            let webSearchQueries = [];
            let webPagesOpened = [];

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
             * NOT FOR GEMINI
            * Add an image to the chat. Will automatically include an image with gpt-4-turbo-2024-04-09.
            * @param {string} imageUrl - The URL of the image to add.
            * @returns {Chat} - The current Chat instance.
            */
            this.addImage = function (imageUrl) {
                messages.push(
                    {
                        role: "user",
                        content: [
                            {
                                type: "image_url",
                                image_url: { url: imageUrl }
                            }
                        ]
                    }
                );
                vision = true;
                return this;
            };

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
             * @param {true|"only_retrieve_results"} scope - set to true to enable full browsing, or "only_retrieve_results" to search Google without opening web pages. 
             * @param {string} [urlOrsearchEngineId] - A specific site you want to restrict the search on or a Search engine ID. 
             * @returns {Chat} - The current Chat instance.
             */
            this.enableBrowsing = function (scope, urlOrsearchEngineId) {
                if (scope) {
                    browsing = true;
                    if (scope == "only_retrieve_results") {
                        onlyRetrieveSearchResults = true;
                    }
                }
                if (urlOrsearchEngineId) {
                    restrictSearch = urlOrsearchEngineId;
                }
                return this;
            };

            /**
             * OPTIONAL
             * 
             * Allow genAI to call vision model.
             * @param {true} scope - set to true to enable vision. 
             * @returns {Chat} - The current Chat instance.
             */
            this.enableVision = function (scope) {
                if (scope) {
                    vision = true;
                }
                return this;
            };

            /**
             * OPTIONAL
             * 
             * Enable a thread run with an genAI assistant.
             * @param {string} assistantId - your assistant id
             * @param {string} vectorStoreDescription - a small description of the available knowledge from this assistant
             * @returns {Chat} - The current Chat instance.
             */
            this.retrieveKnowledgeFromAssistant = function (assistantId, vectorStoreDescription) {
                assistantIdentificator = assistantId;
                vectorStore = vectorStoreDescription;
                return this;
            }

            /**
             * OPTIONAL
             * 
             * Enable a thread run with an genAI assistant.
             * @param {string} assistantId - your assistant id
             * @param {string} attachmentId - the Drive ID of the document you want to attach
             * @returns {Chat} - The current Chat instance.
             */
            this.analyzeDocumentWithAssistant = function (assistantId, attachmentId) {
                assistantIdentificator = assistantId;
                attachmentIdentificator = attachmentId;
                return this;
            }

            /**
             * Includes the content of a web page in the prompt sent to openAI
             * @param {string} url - the url of the webpage you want to fetch
             * @returns {Chat} - The current Chat instance.
             */
            this.addKnowledgeLink = function (url) {
                knowledgeLink = url;
                return this;
            };

            /**
             * If you want to limit the number of calls to the OpenAI API
             * A good way to avoid infinity loops and manage your budget.
             * @param {number} maxAPICalls - 
             */
            this.setMaximumAPICalls = function (maxAPICalls) {
                maximumAPICalls = maxAPICalls;
            };

            this._toJson = function () {
                return {
                    messages: messages,
                    tools: tools,
                    model: model,
                    temperature: temperature,
                    max_tokens: max_tokens,
                    browsing: browsing,
                    webSearchQueries: webSearchQueries,
                    webPagesOpened: webPagesOpened,
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
             * @param {"gemini-1.5-pro-002" | "gemini-1.5-pro" | "gemini-1.5-flash-002" | "gemini-1.5-flash" | "gpt-3.5-turbo" | "gpt-3.5-turbo-16k" | "gpt-4" | "gpt-4-32k" | "gpt-4-1106-preview" | "gpt-4-turbo-preview" | "gpt-4o" | "o1" | "o1-mini"} [advancedParametersObject.model]
             * @param {number} [advancedParametersObject.temperature]
             * @param {number} [advancedParametersObject.max_tokens]
             * @param {string} [advancedParametersObject.function_call]
             * @returns {object} - the last message of the chat 
             */
            this.run = function (advancedParametersObject) {

                if (browsing && !googleCustomSearchAPIKey) {
                    throw Error("Please set your Google custom search API key using GenAIApp.setGoogleSearchAPIKey(yourAPIKey)");
                }

                if (advancedParametersObject) {
                    if (advancedParametersObject.model) {
                        model = advancedParametersObject.model;
                        if (model.includes("gemini")) {
                            if (!geminiKey && (!region || !gcpProjectId)) {
                                throw Error("Please set your Gemini API key or GCP project auth using GenAIApp.setGeminiAPIKey(YOUR_GEMINI_API_KEY) or GenAIApp.setGeminiAuth(YOUR_PROJECT_ID, REGION)");
                            }
                        } else {
                            if (!openAIKey) {
                                if (googleCustomSearchAPIKey) {
                                    throw Error("Careful to use setOpenAIAPIKey to set your OpenAI API key and not setGoogleSearchAPIKey.");
                                }
                                else {
                                    throw Error("Please set your OpenAI API key using GenAIApp.setOpenAIAPIKey(yourAPIKey)");
                                }
                            }
                        }
                    }
                    if (advancedParametersObject.temperature) {
                        temperature = advancedParametersObject.temperature;
                    }
                    if (advancedParametersObject.max_tokens) {
                        max_tokens = advancedParametersObject.max_tokens;
                    }
                }

                if (knowledgeLink) {
                    let knowledge = _urlFetch(knowledgeLink);
                    if (!knowledge) {
                        throw Error(`The webpage ${knowledgeLink} didn't respond, please change the url of the addKnowledgeLink() function.`);
                    }
                    messages.push({
                        role: "system",
                        content: `Information to help with your response (publicly available here: ${knowledgeLink}):\n\n${knowledge}`
                    });
                    contents.push({
                        role: "user",
                        parts: {
                            text: `Information to help with your response (publicly available here: ${knowledgeLink}):\n\n${knowledge}`
                        }
                    })
                    knowledgeLink = null;
                }

                let payload;
                if (model.includes("gemini")) {
                    payload = this._buildGeminiPayload(advancedParametersObject);
                }
                else {
                    payload = this._buildOpenAIPayload(advancedParametersObject);
                }

                let responseMessage;
                if (numberOfAPICalls <= maximumAPICalls) {
                    let endpointUrl = "https://api.openai.com/v1/chat/completions";
                    if (model.includes("gemini")) {
                        if (geminiKey) {
                            endpointUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${geminiKey}`;
                        }
                        else {
                            endpointUrl = `https://${region}-aiplatform.googleapis.com/v1/projects/${gcpProjectId}/locations/${region}/publishers/google/models/${model}:generateContent`;
                        }
                    }
                    responseMessage = _callGenAIApi(endpointUrl, payload);
                    numberOfAPICalls++;
                } else {
                    throw new Error(`Too many calls to genAI API: ${numberOfAPICalls}`);
                }

                if (tools.length > 0) {
                    // Check if AI model wanted to call a function
                    if (model.includes("gemini")) {
                        if (responseMessage?.parts?.[0]?.functionCall) {
                            contents = _handleGeminiToolCalls(responseMessage, tools, contents, webSearchQueries, webPagesOpened);
                            // check if endWithResults or onlyReturnArguments
                            if (contents[contents.length - 1].role == "model") {
                                if (contents[contents.length - 1].parts.text == "endWithResult") {
                                    if (verbose) {
                                        console.log("Conversation stopped because end function has been called");
                                    }
                                    return contents[contents.length - 2].parts.text; // the last chat completion
                                } else if (contents[contents.length - 1].parts.text == "onlyReturnArguments") {
                                    if (verbose) {
                                        console.log("Conversation stopped because argument return has been enabled - No function has been called");
                                    }
                                    return contents[contents.length - 2].parts[0].functionCall.args; // the argument(s) of the last function called
                                }
                            }
                        }
                        else {
                            // if no function has been found, stop here
                            return responseMessage.parts[0].text;
                        }
                    }
                    else {
                        if (responseMessage.tool_calls) {
                            messages = _handleOpenAIToolCalls(responseMessage, tools, messages, webSearchQueries, webPagesOpened);
                            // check if endWithResults or onlyReturnArguments
                            if (messages[messages.length - 1].role == "system") {
                                if (messages[messages.length - 1].content == "endWithResult") {
                                    if (verbose) {
                                        console.log("Conversation stopped because end function has been called");
                                    }
                                    return messages[messages.length - 2]; // the last chat completion
                                } else if (messages[messages.length - 1].content == "onlyReturnArguments") {
                                    if (verbose) {
                                        console.log("Conversation stopped because argument return has been enabled - No function has been called");
                                    }
                                    return _parseResponse(messages[messages.length - 3].tool_calls[0].function.arguments); // the argument(s) of the last function called
                                }
                            }
                        }
                        else {
                            // if no function has been found, stop here
                            return responseMessage.content;
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
                        return responseMessage.parts[0].text;
                    }
                    else {
                        return responseMessage.content;
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
             * @param {Object} advancedParametersObject - An object containing additional parameters for the API call,
             *                                            such as function call preferences.
             * @returns {Object} - The payload object, configured with messages, model settings, and tool selections 
             *                     for OpenAI's API.
             * @throws {Error} If an incompatible model is selected with certain functionalities (e.g., Gemini model with assistant).
             */
            this._buildOpenAIPayload = function (advancedParametersObject) {

                if (model.includes("o1")) {
                    // Developer messages are the new system messages: Starting with o1-2024-12-17, o1 models support developer messages rather than system messages, to align with the chain of command behavior described in the model spec. 
                    messages.forEach(message => {
                        if (message.role == "system") {
                            message.role == "developer";
                        }
                    })
                }

                let payload = {
                    'messages': messages,
                    'model': model,
                    'max_completion_tokens': max_tokens,
                    'temperature': temperature,
                    'user': Session.getTemporaryActiveUserKey()
                };

                if (browsing) {
                    if (messages[messages.length - 1].role !== "tool") {
                        tools.push({
                            type: "function",
                            function: webSearchFunction
                        });
                        let messageContent = `You are able to perform search queries on Google using the function webSearch and read the search results. `;
                        if (!onlyRetrieveSearchResults) {
                            messageContent += "Then you can select a search result and read the page content using the function urlFetch. ";
                            tools.push({
                                type: "function",
                                function: urlFetchFunction
                            });
                        }
                        messages.push({
                            role: "system",
                            content: messageContent
                        });
                        payload.tool_choice = {
                            type: "function",
                            function: { name: "webSearch" }
                        };
                    }
                    else if (messages[messages.length - 1].role == "tool" &&
                        messages[messages.length - 1].name === "webSearch" &&
                        !onlyRetrieveSearchResults) {
                        if (messages[messages.length - 1].content !== noResultFromWebSearchMessage) {
                            // force genAI to call the function urlFetch after retrieving results for a particular search
                            payload.tool_choice = {
                                type: "function",
                                function: { name: "urlFetch" }
                            };
                        }
                    }
                }

                if (assistantIdentificator) {
                    if (model.includes("gemini")) {
                        throw Error("To use OpenAI's assitant, please select a different model than Gemini");
                    }
                    // This function is created only here to adapt the function description to the vector store content
                    let runOpenAIAssistantFunction = new FunctionObject()
                        .setName("runOpenAIAssistant")
                        .setDescription(`To retrieve information from : ${vectorStore}`)
                        .addParameter("assistantId", "string", "The ID of the assistant")
                        .addParameter("prompt", "string", "The question you want to ask the assistant")
                        .endWithResult(true);

                    if (attachmentIdentificator) {
                        runOpenAIAssistantFunction.setDescription("To analyze a file with code interpreter")
                        runOpenAIAssistantFunction.addParameter("attachmentId", "string", "the Id of the file attached");
                    }

                    if (numberOfAPICalls == 0) {

                        tools.push({
                            type: "function",
                            function: runOpenAIAssistantFunction
                        });

                        if (attachmentIdentificator) {
                            messages.push({
                                role: "system",
                                content: `You can use the assistant ${assistantIdentificator} to analyze this file: "${attachmentIdentificator}"`
                            });
                        } else {
                            messages.push({
                                role: "system",
                                content: `You can use the assistant ${assistantIdentificator} to retrieve information from : ${vectorStore}`
                            });
                        }

                        payload.tool_choice = {
                            type: "function",
                            function: { name: "runOpenAIAssistant" }
                        };
                    }
                }

                if (vision && numberOfAPICalls == 0 && !model.includes("gemini")) {
                    tools.push({
                        type: "function",
                        function: imageDescriptionFunction
                    });
                    let messageContent = `You are able to retrieve images description using the getImageDescription function.`;
                    messages.push({
                        role: "system",
                        content: messageContent
                    });
                    payload.tool_choice = {
                        type: "function",
                        function: { name: "getImageDescription" }
                    };
                }

                if (tools.length > 0) {
                    // the user has added functions, enable function calling
                    let payloadTools = Object.keys(tools).map(t => ({
                        type: "function",
                        function: {
                            name: tools[t].function._toJson().name,
                            description: tools[t].function._toJson().description,
                            parameters: tools[t].function._toJson().parameters
                        }
                    }));
                    payload.tools = payloadTools;

                    if (!payload.tool_choice) {
                        payload.tool_choice = 'auto';
                    }

                    if (advancedParametersObject?.function_call &&
                        payload.tool_choice.function?.name !== "urlFetch" &&
                        payload.tool_choice.function?.name !== "webSearch") {
                        // the user has set a specific function to call
                        let tool_choosing = {
                            type: "function",
                            function: {
                                name: advancedParametersObject.function_call
                            }
                        };
                        payload.tool_choice = tool_choosing;
                    } else if (messages[messages.length - 1].role == "tool" && messages[messages.length - 1].name == _urlFetch) {
                        // Once we've opened a web page,
                        // let the model decide what to do
                        // eg: model can either be satisfied with the info found in the web page or decide to open another web page
                        payload.tool_choice = 'auto'
                    }
                }
                return payload;
            }

            /**
             * Builds and returns a payload for a Gemini API call, configuring content, model parameters, 
             * and tool settings based on advanced options and feature flags such as browsing and vision. 
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
                let payload = {
                    'contents': contents,
                    'model': model,
                    'generationConfig': {
                        maxOutputTokens: max_tokens,
                        temperature: temperature
                    },
                    'tool_config': {
                        function_calling_config: {
                            mode: "AUTO"
                        }
                    }
                };

                if (advancedParametersObject?.function_call
                    && advancedParametersObject.function_call !== "urlFetch"
                    && advancedParametersObject.function_call !== "webSearch") {

                    if (model == "gemini-1.5-pro" || model == "gemini-1.5-flash") {
                        payload.tool_config.function_calling_config.mode = "ANY";
                        payload.tool_config.function_calling_config.allowed_function_names = advancedParametersObject.function_call;
                        delete advancedParametersObject.function_call;
                    }
                    else {
                        // https://cloud.google.com/vertex-ai/generative-ai/docs/model-reference/function-calling
                        console.warn(`Unable to force Gemini call to ${advancedParametersObject.function_call} : "function calling with controlled generation" or "forced function calling" is  available with only the Gemini 1.5 Pro and Gemini 1.5 Flash models.`);
                    }

                }

                if (browsing) {
                    if (contents[contents.length - 1].role !== "tool") {
                        tools.push({
                            type: "function",
                            function: webSearchFunction
                        });
                        let messageContent = `You are able to perform search queries on Google using the function webSearch and read the search results. `;
                        if (!onlyRetrieveSearchResults) {
                            messageContent += "Then you can select a search result and read the page content using the function urlFetch. ";
                            tools.push({
                                type: "function",
                                function: urlFetchFunction
                            });
                        }
                        contents.push({
                            role: "user",
                            parts: {
                                text: messageContent
                            },
                        });
                    }
                }

                if (assistantIdentificator) {
                    throw Error("To use OpenAI's assistant, please select a different model than Gemini");
                }

                if (vision && numberOfAPICalls == 0) {
                    tools.push({
                        type: "function",
                        function: imageDescriptionFunction
                    });
                    let messageContent = `You are able to retrieve images description using the getImageDescription function.`;
                    contents.push({
                        role: "system",
                        parts: {
                            text: messageContent
                        }
                    });
                }

                if (tools.length > 0) {
                    // the user has added functions, enable function calling
                    let payloadTools = Object.keys(tools).map(t => {
                        let toolFunction = tools[t].function._toJson();

                        const parameters = toolFunction.parameters;
                        if (parameters && parameters.type) {
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
                    }]

                }
                return payload;
            }
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
            authMethod = 'Bearer ' + ScriptApp.getOAuthToken();
        }
        let maxRetries = 5;
        let retries = 0;
        let success = false;

        let responseMessage, finish_reason;
        while (retries < maxRetries && !success) {
            let options = {
                'method': 'post',
                'headers': {
                    'Content-Type': 'application/json',
                    'Authorization': authMethod
                },
                'payload': JSON.stringify(payload),
                'muteHttpExceptions': true
            };

            let response = UrlFetchApp.fetch(endpoint, options);
            let responseCode = response.getResponseCode();

            if (responseCode === 200) {
                // The request was successful, exit the loop.
                let parsedResponse = JSON.parse(response.getContentText());;
                if (endpoint.includes("google")) {
                    responseMessage = parsedResponse.candidates[0].content;
                    finish_reason = parsedResponse.candidates[0].finish_reason;
                } else {
                    responseMessage = parsedResponse.choices[0].message;
                    finish_reason = parsedResponse.choices[0].finish_reason;
                }
                if (finish_reason == "length") {
                    console.warn(`GenAI response has been troncated because it was too long. To resolve this issue, you can increase the max_tokens property. max_tokens: ${payload.max_tokens}, prompt_tokens: ${parsedResponse.usage.prompt_tokens}, completion_tokens: ${parsedResponse.usage.completion_tokens}`);
                }
                success = true;
            }
            else if (responseCode === 429) {
                console.warn(`Rate limit reached when calling genAI API, will automatically retry in a few seconds.`);
                // Rate limit reached, wait before retrying.
                let delay = Math.pow(2, retries) * 1000; // Delay in milliseconds, starting at 1 second.
                Utilities.sleep(delay);
                retries++;
            }
            else if (responseCode === 503 || responseCode === 500) {
                // The server is temporarily unavailable, or an issue occured on OpenAI servers. wait before retrying.
                // https://platform.openai.com/docs/guides/error-codes/api-errors
                let delay = Math.pow(2, retries) * 1000; // Delay in milliseconds, starting at 1 second.
                Utilities.sleep(delay);
                retries++;
            }
            else {
                // The request failed for another reason, log the error and exit the loop.
                console.error(`Request to genAI failed with response code ${responseCode} - ${response.getContentText()}`);
                break;
            }
        }

        if (!success) {
            throw new Error(`Failed to call genAI API after ${retries} retries.`);
        }

        if (verbose) {
            if (endpoint.includes("google")) {
                Logger.log({
                    message: `Got response from Gemini API.`,
                    responseMessage: responseMessage
                });
            } else {
                Logger.log({
                    message: `Got response from OpenAI API.`,
                    responseMessage: responseMessage
                });
            }

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
     * @param {Array} webSearchQueries - List to accumulate queries made during web search tool calls.
     * @param {Array} webPagesOpened - List to accumulate URLs accessed during URL fetch tool calls.
     * @returns {Array} - The updated contents array, representing the conversation flow with function calls and responses.
     */
    function _handleGeminiToolCalls(responseMessage, tools, contents, webSearchQueries, webPagesOpened) {
        for (let tool_call in responseMessage.parts) {
            // Call the function
            let functionName = responseMessage.parts[tool_call].functionCall.name;
            let functionArgs = responseMessage.parts[tool_call].functionCall.args;
            contents.push({
                role: "model",
                parts: [{
                    functionCall: {
                        "name": functionName,
                        "args": functionArgs
                    }
                }]
            })

            let argsOrder = [];
            let endWithResult = false;
            let onlyReturnArguments = false;

            for (let t in tools) {
                let currentFunction = tools[t].function._toJson();
                if (currentFunction.name == functionName) {
                    argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order
                    endWithResult = currentFunction.endingFunction;
                    onlyReturnArguments = currentFunction.onlyArgs;
                    break;
                }
            }

            if (endWithResult) {
                // User defined that if this function has been called, then no more actions should be performed with the chat.
                let functionResponse = _callFunction(functionName, functionArgs, argsOrder);
                if (typeof functionResponse != "string") {
                    if (typeof functionResponse == "object") {
                        functionResponse = JSON.stringify(functionResponse);
                    }
                    else {
                        functionResponse = String(functionResponse);
                    }
                }
                contents.push({
                    "role": "user",
                    "parts": {
                        text: functionResponse
                    }
                });
                contents.push({
                    "role": "model",
                    "parts": {
                        text: "endWithResult"
                    }
                });
                return contents;
            }
            else if (onlyReturnArguments) {
                contents.push({
                    "role": "model",
                    "parts": {
                        text: "onlyReturnArguments"
                    }
                });
                return contents;
            }
            else {

                let functionResponse = _callFunction(functionName, functionArgs, argsOrder);
                if (typeof functionResponse != "string") {
                    if (typeof functionResponse == "object") {
                        functionResponse = JSON.stringify(functionResponse);
                    }
                    else {
                        functionResponse = String(functionResponse);
                    }
                }

                if (functionName == "webSearch") {
                    webSearchQueries.push(functionArgs.q);
                }
                else if (functionName == "urlFetch") {
                    webPagesOpened.push(functionArgs.url);
                    if (!functionResponse) {
                        if (verbose) {
                            console.log("The website didn't respond, going back to search results.");
                        }
                        let searchResults = JSON.parse(contents[contents.length - 1].parts[0].text);
                        let updatedSearchResults = searchResults.filter(function (obj) {
                            return obj.link !== functionArgs.url;
                        });
                        contents[contents.length - 1].parts[0].text = JSON.stringify(updatedSearchResults);
                    }
                    if (verbose) {
                        console.log("Web page opened, let model decide what to do next (open another web page or perform another action).");
                    }
                }
                else {
                    if (verbose) {
                        console.log(`function ${functionName}() called by Gemini.`);
                    }
                }
                contents.push({
                    "role": "user",
                    "parts": {
                        text: functionResponse
                    }
                });
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
     * @param {Array} webSearchQueries - Array that accumulates queries made during web search tool calls.
     * @param {Array} webPagesOpened - Array that tracks URLs accessed during URL fetch tool calls.
     * @returns {Array} - The updated messages array, representing the conversation flow with function calls, results, and system responses.
     */
    function _handleOpenAIToolCalls(responseMessage, tools, messages, webSearchQueries, webPagesOpened) {
        messages.push(responseMessage);
        for (let tool_call in responseMessage.tool_calls) {
            if (responseMessage.tool_calls[tool_call].type == "function") {
                // Call the function
                let functionName = responseMessage.tool_calls[tool_call].function.name;
                let functionArgs = _parseResponse(responseMessage.tool_calls[tool_call].function.arguments);

                let argsOrder = [];
                let endWithResult = false;
                let onlyReturnArguments = false;

                for (let t in tools) {
                    let currentFunction = tools[t].function._toJson();
                    if (currentFunction.name == functionName) {
                        argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order
                        endWithResult = currentFunction.endingFunction;
                        onlyReturnArguments = currentFunction.onlyArgs;
                        break;
                    }
                }

                if (endWithResult) {
                    // User defined that if this function has been called, then no more actions should be performed with the chat.
                    let functionResponse = _callFunction(functionName, functionArgs, argsOrder);
                    if (typeof functionResponse != "string") {
                        if (typeof functionResponse == "object") {
                            functionResponse = JSON.stringify(functionResponse);
                        }
                        else {
                            functionResponse = String(functionResponse);
                        }
                    }
                    messages.push({
                        "tool_call_id": responseMessage.tool_calls[tool_call].id,
                        "role": "tool",
                        "name": functionName,
                        "content": functionResponse,
                    });
                    messages.push({
                        "role": "system",
                        "content": functionResponse
                    });
                    messages.push({
                        "role": "system",
                        "content": "endWithResult"
                    });
                    return messages;
                }
                else if (onlyReturnArguments) {
                    messages.push({
                        "tool_call_id": responseMessage.tool_calls[tool_call].id,
                        "role": "tool",
                        "name": functionName,
                        "content": "",
                    });
                    messages.push({
                        "role": "system",
                        "content": "onlyReturnArguments"
                    });
                    return messages;
                }
                else {
                    let functionResponse = _callFunction(functionName, functionArgs, argsOrder);
                    if (typeof functionResponse != "string") {
                        if (typeof functionResponse == "object") {
                            functionResponse = JSON.stringify(functionResponse);
                        }
                        else {
                            functionResponse = String(functionResponse);
                        }
                    }

                    if (functionName == "webSearch") {
                        webSearchQueries.push(functionArgs.q);
                    }
                    else if (functionName == "urlFetch") {
                        webPagesOpened.push(functionArgs.url);
                        if (!functionResponse) {
                            if (verbose) {
                                console.log("The website didn't respond, going back to search results.");
                            }
                            let searchResults = JSON.parse(messages[messages.length - 1].content);
                            let updatedSearchResults = searchResults.filter(function (obj) {
                                return obj.link !== functionArgs.url;
                            });
                            messages[messages.length - 1].content = JSON.stringify(updatedSearchResults);
                        }
                        if (verbose) {
                            console.log("Web page opened, let model decide what to do next (open another web page or perform another action).");
                        }
                    }
                    else {
                        if (verbose) {
                            console.log(`function ${functionName}() called by OpenAI.`);
                        }
                    }
                    messages.push({
                        "tool_call_id": responseMessage.tool_calls[tool_call].id,
                        "role": "tool",
                        "name": functionName,
                        "content": functionResponse,
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
        // Handle internal functions
        if (functionName == "webSearch") {
            return _webSearch(jsonArgs.q);
        }
        if (functionName == "urlFetch") {
            return _urlFetch(jsonArgs.url);
        }
        if (functionName == "getImageDescription") {
            if (jsonArgs.fidelity) {
                return _getImageDescription(jsonArgs.imageUrl, jsonArgs.fidelity);
            } else {
                return _getImageDescription(jsonArgs.imageUrl);
            }
        }
        if (functionName == "runOpenAIAssistant") {
            if (jsonArgs.attachmentId) {
                return _runOpenAIAssistant(jsonArgs.assistantId, jsonArgs.prompt, jsonArgs.attachmentId);
            } else {
                return _runOpenAIAssistant(jsonArgs.assistantId, jsonArgs.prompt);
            }
        }
        // Parse JSON arguments
        var argsObj = jsonArgs;
        let argsArray = argsOrder.map(argName => argsObj[argName]);

        // Call the function dynamically
        if (globalThis[functionName] instanceof Function) {
            let functionResponse = globalThis[functionName].apply(null, argsArray);
            if (functionResponse) {
                return functionResponse;
            }
            else {
                return "The function has been sucessfully executed but has nothing to return";
            }
        }
        else {
            throw Error("Function not found or not a function: " + functionName);
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
            let parsedReponse = JSON.parse(response);
            return parsedReponse;
        }
        catch (e) {
            // Split the response into lines
            let lines = String(response).trim().split('\n');

            if (lines[0] !== '{') {
                return null;
            }
            else if (lines[lines.length - 1] !== '}') {
                lines.push('}');
            }
            for (let i = 1; i < lines.length - 1; i++) {
                let line = lines[i].trim();
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
                let parsedResponse = JSON.parse(response);
                return parsedResponse;
            }
            catch (e) {
                // If parsing still fails, log the error and return null.
                console.warn('Error parsing corrected response: ' + e.message);
                return null;
            }
        }
    }

    /**
     * Creates a new thread and returns the thread ID.
     * 
     * @returns {string} The thread ID.
     */
    function _createThread() {
        var url = 'https://api.openai.com/v1/threads';
        var options = {
            'method': 'post',
            'contentType': 'application/json',
            'headers': {
                'Authorization': 'Bearer ' + openAIKey,
                'OpenAI-Beta': 'assistants=v2'
            },
            'payload': '{}'
        };

        var response = UrlFetchApp.fetch(url, options);
        return JSON.parse(response.getContentText()).id;
    }

    /**
     * Uploads a file to OpenAI and returns the file ID.
     * 
     * @param {string} optionalAttachment - The optional attachment ID from Google Drive.
     * @returns {string} The OpenAI file ID.
     */
    function _uploadFileToOpenAI(optionalAttachment) {
        var file = DriveApp.getFileById(optionalAttachment);
        var mimeType = file.getMimeType();
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

        var token = ScriptApp.getOAuthToken();

        // Fetch the file from Google Drive using the generated URL and OAuth token
        var response = UrlFetchApp.fetch(fileBlobUrl, {
            headers: {
                'Authorization': 'Bearer ' + token
            }
        });

        var fileBlob = response.getBlob();

        const openAIFileEndpoint = 'https://api.openai.com/v1/files';

        var formData = {
            'file': fileBlob,
            'purpose': 'assistants'
        };

        var uploadOptions = {
            'method': 'post',
            'headers': {
                'Authorization': 'Bearer ' + openAIKey
            },
            'payload': formData,
            'muteHttpExceptions': true
        };

        response = UrlFetchApp.fetch(openAIFileEndpoint, uploadOptions);
        var uploadedFileResponse = JSON.parse(response.getContentText());
        if (uploadedFileResponse.error) {
            throw new Error('Error: ' + uploadedFileResponse.error.message);
        }
        return uploadedFileResponse.id;
    }

    /**
     * Adds a message to the thread.
     * 
     * @param {string} threadId - The ID of the thread.
     * @param {string} prompt - The prompt to send to the assistant.
     * @param {string} [optionalAttachment] - The optional attachment ID from Google Drive.
     */
    function _addMessageToThread(threadId, prompt, optionalAttachment) {
        let messagePayload = {
            "role": "user",
            "content": prompt
        };

        if (optionalAttachment) {
            try {
                var openAiFileId = _uploadFileToOpenAI(optionalAttachment);
                messagePayload.attachments = [
                    {
                        "file_id": openAiFileId,
                        "tools": [{ "type": "code_interpreter" }]
                    }
                ];
            } catch (e) {
                Logger.log('Error retrieving the file : ' + e.message);
            }
        }

        var url = `https://api.openai.com/v1/threads/${threadId}/messages`;
        var options = {
            'method': 'post',
            'contentType': 'application/json',
            'headers': {
                'Authorization': 'Bearer ' + openAIKey,
                'OpenAI-Beta': 'assistants=v2'
            },
            'payload': JSON.stringify(messagePayload)
        };

        UrlFetchApp.fetch(url, options);
    }

    /**
     * Runs the assistant and returns the run ID.
     * 
     * @param {string} threadId - The ID of the thread.
     * @param {string} assistantId - The ID of the OpenAI assistant to run.
     * @returns {string} The run ID.
     */
    function _runAssistant(threadId, assistantId) {
        var url = `https://api.openai.com/v1/threads/${threadId}/runs`;
        var assistantPayload = {
            "assistant_id": assistantId
        };
        var options = {
            'method': 'post',
            'contentType': 'application/json',
            'headers': {
                'Authorization': 'Bearer ' + openAIKey,
                'OpenAI-Beta': 'assistants=v2'
            },
            'payload': JSON.stringify(assistantPayload)
        };

        var response = UrlFetchApp.fetch(url, options);
        return JSON.parse(response.getContentText()).id;
    }

    /**
     * Monitors the run status until completion.
     * 
     * @param {string} threadId - The ID of the thread.
     * @param {string} runId - The run ID.
     */
    function _monitorRunStatus(threadId, runId) {
        var url = `https://api.openai.com/v1/threads/${threadId}/runs/${runId}`;
        var status = "queued";
        const sleepTime = 30000; // 30 seconds

        while (status === "queued") {
            Utilities.sleep(sleepTime);

            var statusOptions = {
                'method': 'get',
                'headers': {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + openAIKey,
                    'OpenAI-Beta': 'assistants=v2'
                }
            };

            var response = UrlFetchApp.fetch(url, statusOptions);
            let runStatus = JSON.parse(response.getContentText());
            status = runStatus.status;
        }

        if (status !== "completed") {
            Logger.log('Run did not complete in time.');
            throw new Error('Run did not complete in time.');
        }
    }

    /**
     * Retrieves the assistant's response and references.
     * 
     * @param {string} threadId - The ID of the thread.
     * @param {string} assistantId - The ID of the OpenAI assistant.
     * @returns {string} The assistant's response and references in JSON format.
     */
    function _getAssistantResponse(threadId, assistantId) {
        var url = 'https://api.openai.com/v1/threads/' + threadId + '/messages';
        var options = {
            'method': 'get',
            'headers': {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + openAIKey,
                'OpenAI-Beta': 'assistants=v2'
            }
        };

        try {
            var response = UrlFetchApp.fetch(url, options);
            var json = response.getContentText();
            var data = JSON.parse(json).data[0].content[0].text;

            let references = [];
            JSON.parse(JSON.stringify(data.annotations)).forEach(element => {
                const fileId = element.file_citation.file_id;
                var fileEndpoint = 'https://api.openai.com/v1/files/' + fileId;

                var fileOptions = {
                    'method': 'get',
                    'headers': {
                        'Authorization': 'Bearer ' + openAIKey
                    }
                };

                var response = UrlFetchApp.fetch(fileEndpoint, fileOptions);
                var json = JSON.parse(response.getContentText());

                references.push(json.filename);
            });

            Logger.log({
                message: `Got response from Assistant : ${assistantId}`,
                response: JSON.stringify(data.value),
                references: references
            });

            return JSON.stringify({
                response: JSON.stringify(data.value),
                references: references
            });

        } catch (e) {
            Logger.log('Error retrieving assistant response: ' + e.message);
            return 'Execution failure of the documentation retrieval.';
        }
    }

    /**
     * Runs an OpenAI assistant with the provided prompt and optional attachment.
     * 
     * @param {string} assistantId - The ID of the OpenAI assistant to run.
     * @param {string} prompt - The prompt to send to the assistant.
     * @param {string} [optionalAttachment] - The optional attachment ID from Google Drive.
     * @returns {string} The assistant's response and references in JSON format.
     */
    function _runOpenAIAssistant(assistantId, prompt, optionalAttachment) {
        try {
            var threadId = _createThread();
            _addMessageToThread(threadId, prompt, optionalAttachment);
            var runId = _runAssistant(threadId, assistantId);
            _monitorRunStatus(threadId, runId);
            return _getAssistantResponse(threadId, assistantId);
        } catch (e) {
            Logger.log('Error in runOpenAIAssistant: ' + e.message);
            return 'Execution failure.';
        }
    }

    /**
     * Performs a web search using Google Custom Search API with optional restrictions, such as a specific
     * site or a designated search engine. Fetches and formats up to 10 search results, including titles, links, 
     * and snippets, or returns a predefined message if no results are found.
     *
     * @private
     * @param {string} q - The search query string to be used in the web search.
     * @returns {string} - A JSON string containing an array of search results, each with title, link, and snippet, 
     *                     or a message indicating no results were found.
     */
    function _webSearch(q) {
        // https://programmablesearchengine.google.com/controlpanel/overview?cx=221c662683d054b63
        let searchEngineId = "221c662683d054b63";
        let url = `https://www.googleapis.com/customsearch/v1?key=${googleCustomSearchAPIKey}`;

        // If restrictSearch is defined, check wether to restrict to a specific site or use a specific Search Engine
        if (restrictSearch) {
            if (restrictSearch.includes('.')) {
                // Search restricted to specific site
                if (verbose) {
                    console.log(`Site search on ${restrictSearch}`);
                }
                url += `&siteSearch=${encodeURIComponent(restrictSearch)}&siteSearchFilter=i`;
            }
            else {
                // Use the desired Search Engine
                // https://programmablesearchengine.google.com/controlpanel/all
                searchEngineId = restrictSearch;
            }
        }
        url += `&cx=${searchEngineId}&q=${encodeURIComponent(q)}&num=10`;

        const urlfetchResp = UrlFetchApp.fetch(url);
        const resp = JSON.parse(urlfetchResp.getContentText());

        let searchResults;
        if (!resp.items?.length) {
            searchResults = noResultFromWebSearchMessage;
        }
        else {
            // https://developers.google.com/custom-search/v1/reference/rest/v1/Search?hl=en#Result
            searchResults = JSON.stringify(resp.items.slice(0, 10).map(function (item) {
                return {
                    title: item.title,
                    link: item.link,
                    snippet: item.snippet
                };
                // filter to remove undefined values from the results array
            }).filter(Boolean));
        }

        const nbOfResults = resp.searchInformation.totalResults;
        if (verbose) {
            Logger.log({
                message: `Web search : "${q}" - ${nbOfResults} results`,
                searchResults: searchResults
            });
        }

        return searchResults;
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
            console.log(`Clicked on link : ${url}`);
        }
        let response;
        try {
            response = UrlFetchApp.fetch(url);
        }
        catch (e) {

            console.warn(`Error fetching the URL: ${e.message}`);
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
     * Generates a description of an image using OpenAI's API, based on the provided image URL and fidelity level.
     * Verifies that the image URL has a supported extension (png, jpeg, gif, webp) before proceeding. Adjusts the 
     * description fidelity (detail level) based on the fidelity parameter and constructs an AI prompt to request 
     * a description focused on key elements.
     *
     * @private
     * @param {string} imageUrl - The URL of the image to describe.
     * @param {string} [fidelity="low"] - Optional fidelity level of the description, either "low" or "high". Defaults to "low".
     * @returns {string} - The response message from the OpenAI API, containing the description of the image. 
     *                     Returns an error message if the image format is unsupported.
     */
    function _getImageDescription(imageUrl, fidelity) {
        const extensions = ['png', 'jpeg', 'gif', 'webp'];
        if (!extensions.some(extension => imageUrl.toLowerCase().endsWith(`.${extension}`))) {
            if (verbose) {
                Logger.log({
                    message: `Tried to get description of image, but extension is not supported by OpenAI API. Supported extensions are 'png', 'jpeg', 'gif', 'webp'`,
                    imageUrl: imageUrl
                })
            }
            return "This is not a supported image, no description available."
        }

        if (!fidelity) {
            fidelity = "low";
        } else {
            fidelity = "high";
        }

        let imageMessage = [{
            role: "user",
            content: [
                {
                    type: "text",
                    text: "What is the content of this image ? Focus on important element." // Gives a more human friendly description than the initial example prompt given by OpenAI
                },
                {
                    type: "image_url",
                    image_url: {
                        url: imageUrl,
                        detail: fidelity
                    }
                },
            ]
        }];

        let payload = {
            'messages': imageMessage,
            'model': "gpt-4-vision-preview",
            'max_tokens': 1000,
            'user': Session.getTemporaryActiveUserKey()
        };

        let responseMessage = _callGenAIApi("https://api.openai.com/v1/chat/completions", payload);

        return responseMessage;
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
            let headerRow = content.replace(/<th>(.*?)<\/th>/g, '| $1 ').trim() + '|';
            let separatorRow = headerRow.replace(/[^|]+/g, match => {
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
         * @param {string} gcp_project_region - Your GCP project region (ex: us-central1)
         */
        setGeminiAuth: function (gcp_project_id, gcp_project_region) {
            gcpProjectId = gcp_project_id;
            region = gcp_project_region
        },

        /**
         * If you want to enable browsing
         * @param {string} apiKey - Your Google API key.
         */
        setGoogleSearchAPIKey: function (apiKey) {
            googleCustomSearchAPIKey = apiKey;
        },

        /**
         * If you want to acces what occured during the conversation
         * @param {Chat} chat - your chat instance.
         * @returns {object} - the web search queries, the web pages opened and an historic of all the messages of the chat
         */
        debug: function (chat) {
            return {
                getWebSearchQueries: function () {
                    return chat._toJson().webSearchQueries
                },
                getWebPagesOpened: function () {
                    return chat._toJson().webPagesOpened
                }
            }
        }
    }
})();
