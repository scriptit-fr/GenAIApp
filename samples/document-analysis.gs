/*
 * Purpose: Demonstrates document analysis with addFile() using Drive IDs and Blobs.
 * Use case: Summarize PDFs or exported Google Workspace files from Apps Script.
 * Required config: Store an OpenAI API key in Script Properties as OPENAI_API_KEY.
 * Expected output: Logs three concise bullets summarizing the supplied documents.
 */
function documentAnalysisSample() {
  const scriptProperties = PropertiesService.getScriptProperties();
  GenAIApp.setOpenAIAPIKey(scriptProperties.getProperty('OPENAI_API_KEY'));

  const textBlob = Utilities.newBlob(`Ah no! young blade! That was a trifle short! 
    You might have said at least a hundred things 
    By varying the tone. . .like this, suppose,. . . 
    Aggressive: 'Sir, if I had such a nose I'd amputate it!' 
    Friendly: 'When you sup It must annoy you, dipping in your cup; 
    You need a drinking-bowl of special shape!' 
    Descriptive: ''Tis a rock!. . .a peak!. . .a cape! --
    A cape, forsooth! 'Tis a peninsular!' 
    Curious: 'How serves that oblong capsular? 
    For scissor-sheath? Or pot to hold your ink?' 
    Gracious: 'You love the little birds, I think? 
    I see you've managed with a fond research 
    To find their tiny claws a roomy perch!' `, 'text/plain', 'goals.txt');

  const chat = GenAIApp.newChat()
    .addMessage('Summarize the attached file in three bullets.')
    .addFile(textBlob);

  const response = chat.run({ model: 'gpt-5.6-terra' });
  Logger.log(response);
}
