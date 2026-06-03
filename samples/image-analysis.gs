/*
 * Purpose: Demonstrates image analysis from both a public URL and an Apps Script Blob.
 * Use case: Send screenshots, Drive images, or public images to a vision-capable model.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs a short comparison of the URL image and generated Blob image.
 */
function imageAnalysisSample() {
  GenAIApp.setOpenAIAPIKey(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'));

  const imageUrl = 'https://www.gstatic.com/images/branding/product/2x/apps_script_48dp.png';
  const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob().setName('apps-script-logo.png');

  const chat = GenAIApp.newChat()
    .addMessage('Describe these images and mention whether they appear related.')
    .addImage(imageUrl)
    .addImage(imageBlob);

  const response = chat.run({ model: 'gpt-5.4' });
  Logger.log(response);
}
