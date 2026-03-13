const GEMINI_API_KEY = 'AIzaSyBJyaXH5QPSco9Drlj8b79VktH2mlt7HdU'; // 請在此輸入您的 API Key



const GEMINI_MODEL = 'gemini-3-flash-preview';

/**
 * Gemini API 調用函數 (全域共用，具備重試機制)
 */
function callGemini(prompt) {
  if (GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY') {
    return "請先設定 API Key";
  }

  const maxRetries = 3;
  let retryCount = 0;
  let waitTime = 2000;

  while (retryCount <= maxRetries) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
    const payload = {
      "contents": [{
        "parts": [{ "text": prompt }]
      }]
    };
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();
      const result = JSON.parse(responseText);

      if (statusCode === 200 && result.candidates && result.candidates[0].content) {
        return result.candidates[0].content.parts[0].text.trim();
      }

      if ((statusCode === 503 || statusCode === 429) && retryCount < maxRetries) {
        Logger.log(`API 繁忙 (${statusCode})，第 ${retryCount + 1} 次重試...`);
        Utilities.sleep(waitTime);
        retryCount++;
        waitTime *= 2;
        continue;
      }

      return "AI 分析失敗：" + (result.error ? result.error.message : responseText);
    } catch (e) {
      if (retryCount < maxRetries) {
        Utilities.sleep(waitTime);
        retryCount++;
        waitTime *= 2;
        continue;
      }
      return "API 調用出錯：" + e.toString();
    }
  }
}
