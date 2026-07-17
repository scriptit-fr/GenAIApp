/*
 * Purpose: Demonstrates a single function-calling tool registered on a chat.
 * Use case: Let the model call Apps Script code to retrieve structured app data.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a weather-style answer based on the sampleGetWeather stub result.
 */
function functionCallingBasicsSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const weatherFunction = GenAIApp.newFunction()
    .setName('sampleGetWeather')
    .setDescription('Gets the current weather for a city.')
    .addParameter('city', 'string', 'City name, for example Paris')
    .addParameter('unit', 'string', 'Temperature unit: celsius or fahrenheit', true);

  const chat = GenAIApp.newChat()
    .addMessage('What is the weather in Paris? Use the weather function.')
    .addFunction(weatherFunction);

  const response = chat.run({ model: 'gpt-5.6-terra', function_call: 'sampleGetWeather' });
  Logger.log(response);
}

function sampleGetWeather(city, unit) {
  return {
    city: city,
    unit: unit || 'celsius',
    condition: 'sunny',
    temperature: 22
  };
}
