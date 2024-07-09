// License: public domain (https://creativecommons.org/publicdomain/zero/1.0/)
// Kita Toshihiro https://tkita.net 2024
// Google Spreadsheet用 Apps Script

const GPT_API_URL = "https://api.openai.com/v1/chat/completions";
//const DEFAULT_MODEL = "gpt-4o";
const DEFAULT_MODEL = "gpt-3.5-turbo";

const GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/";
//const GEMINI_DEFAULT_MODEL = "gemini-1.5-pro-latest";
const GEMINI_DEFAULT_MODEL = "gemini-1.5-flash-latest";

// OpenAI GPT
function GPT(prompt, model = DEFAULT_MODEL) {
  if (model === ""){
    model = DEFAULT_MODEL;
  }
  const apiKey = getApiKey('A2');
  const json = {
    model: model,
    messages: [{ role: "user", content: prompt }],
    max_tokens: 250,
    temperature: 0.7
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(json)
  };

  const response = UrlFetchApp.fetch(GPT_API_URL, options);
  const responseData = JSON.parse(response.getContentText());
  if (responseData.error) {
    return responseData.error.message;
  } else {
    return responseData.choices[0].message.content.trim();
  }
}

// Google Gemini
function Gemini(prompt, model = GEMINI_DEFAULT_MODEL) {
  if (model === ""){
    model = GEMINI_DEFAULT_MODEL;
  }
  const apiKey = getApiKey('A3');
  const apiURL = `${GEMINI_API_URL}${model}:generateContent?key=${apiKey}`;
  const json = { contents: [{ parts: [{ text: prompt }] }] };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(json)
  };

  const response = UrlFetchApp.fetch(apiURL, options);
  const responseData = JSON.parse(response.getContentText());

  if (responseData.error) {
    return responseData.error.message;
  } else {
    return responseData.candidates[0].content.parts[0].text.trim();
  }
}


function getApiKey(cell) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('api');
  return sheet.getRange(cell).getValue();
}

// Utility functions for Google Sheets
function GPTrange(range, model = DEFAULT_MODEL) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = range;
    //const values = sheet.getRange(range).getValues();
  const prompt = values.map(row => row.join(' | ')).join('\n');
  return GPT(prompt, model);
}

function GPTtranslate(prompt, model = DEFAULT_MODEL, lang = "English") {
  return GPT(`${prompt}\n\n この文章を${lang}に翻訳したもの: `, model);
}

function GPTsummary(prompt, model = DEFAULT_MODEL, length = 200) {
  return GPT(`${prompt}\n\n この文章を${length}文字で要約したもの: `, model);
}


function GeminiRange(range, model = GEMINI_DEFAULT_MODEL) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //const values = sheet.getRange(range).getValues();
  const values = range;
  const prompt = values.map(row => row.join(' | ')).join('\n');
  return Gemini(prompt, model);
}

function GeminiTranslate(prompt, model = GEMINI_DEFAULT_MODEL, lang = "English") {
  return Gemini(`${prompt}\n\n この文章の${lang}に翻訳したもの: `, model);
}

function GeminiSummary(prompt, model = GEMINI_DEFAULT_MODEL, length = 200) {
  return Gemini(`${prompt}\n\n この文章を${length}文字で要約したもの: `, model);
}


