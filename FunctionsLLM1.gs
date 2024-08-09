// License: public domain (https://creativecommons.org/publicdomain/zero/1.0/)
// Kita Toshihiro https://tkita.net 2024
// Google Spreadsheet用 Apps Script

const GPT_API_URL = "https://api.openai.com/v1/chat/completions";
//const DEFAULT_MODEL = "gpt-4o";
const DEFAULT_MODEL = "gpt-4o-mini";
//https://openai.com/api/pricing/

const GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/";
//const GEMINI_DEFAULT_MODEL = "gemini-1.5-pro-latest";
const GEMINI_DEFAULT_MODEL = "gemini-1.5-flash-latest";
//https://ai.google.dev/pricing

// OpenAI GPT
function GPT(prompt, model = DEFAULT_MODEL) {
  if (model === ""){
    model = DEFAULT_MODEL;
  }
  const apiKey = getApiKey('A2');
  prompt = "300文字以内で出力してください。" + prompt
  // max_tokens でも指定できるが、文章の途中で突然切れた出力になることが多い。
  const json = {
    model: model,
    messages: [{ role: "user", content: prompt }],
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
  prompt = "300文字以内で出力してください。" + prompt
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

function GPTtranslate(prompt, lang = "English", model = DEFAULT_MODEL) {
  return GPT(`${prompt}\n\n この文章を${lang}に翻訳したもの: `, model);
}

function GPTsummary(prompt, length = 150, model = DEFAULT_MODEL) {
  return GPT(`${prompt}\n\n この文章を${length}文字で要約したもの: `, model);
}


function GeminiRange(range, model = GEMINI_DEFAULT_MODEL) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //const values = sheet.getRange(range).getValues();
  const values = range;
  const prompt = values.map(row => row.join(' | ')).join('\n');
  return Gemini(prompt, model);
}

function GeminiTranslate(prompt, lang = "English", model = GEMINI_DEFAULT_MODEL) {
  return Gemini(`${prompt}\n\n この文章の${lang}に翻訳したもの: `, model);
}

function GeminiSummary(prompt, length = 150, model = GEMINI_DEFAULT_MODEL) {
  return Gemini(`${prompt}\n\n この文章を${length}文字で要約したもの: `, model);
}
