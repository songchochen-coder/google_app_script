/**
 * Global_Config.js
 * ================
 * 所有腳本共用的設定與函式。
 * callGemini、parseBatchResponse、renderMarkdownToSheet 只在此定義一次。
 *
 * ⚠️  API Key 儲存方式：
 *   請在 GAS 編輯器中執行一次 setup() 函式，即可安全儲存 Key。
 *   Key 不會出現在程式碼或 GitHub 上。
 */

// ── 模型設定 ────────────────────────────────────────────────────
const GEMINI_MODEL = 'gemini-3-flash-preview';

// ── 初始化：設定 API Key（只需執行一次）────────────────────────
function setup() {
  const key = '請在這裡填入你的 Gemini API Key';
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('✅ API Key 已安全儲存至 Script Properties。');
}

// ── 取得 API Key ────────────────────────────────────────────────
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('❌ 找不到 API Key，請先執行 setup() 函式設定。');
  return key;
}

// ── Gemini API 呼叫（唯一版本，含重試機制）─────────────────────
/**
 * @param {string} prompt - 傳入的提示詞
 * @param {object} [config] - 可覆寫的 generationConfig
 * @returns {string} AI 回傳的文字
 */
function callGemini(prompt, config) {
  const apiKey = getApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const generationConfig = Object.assign(
    { temperature: 0.3, maxOutputTokens: 2000 },
    config || {}
  );

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: generationConfig
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxRetries = 3;
  let retryCount = 0;
  let waitTime = 2000;

  while (retryCount <= maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const json = JSON.parse(response.getContentText());

      if (statusCode === 200 && json.candidates && json.candidates[0].content) {
        return json.candidates[0].content.parts[0].text.trim();
      }

      if ((statusCode === 503 || statusCode === 429) && retryCount < maxRetries) {
        Logger.log(`⚠️  API 繁忙 (${statusCode})，第 ${retryCount + 1} 次重試...`);
        Utilities.sleep(waitTime);
        retryCount++;
        waitTime *= 2;
        continue;
      }

      return 'AI 分析失敗：' + (json.error ? json.error.message : response.getContentText());
    } catch (e) {
      if (retryCount < maxRetries) {
        Utilities.sleep(waitTime);
        retryCount++;
        waitTime *= 2;
        continue;
      }
      return 'API 呼叫出錯：' + e.toString();
    }
  }
}

// ── 解析 AI 批次回傳（唯一版本）───────────────────────────────
/**
 * 解析 "代號:說明" 的多行格式，相容全形/半形冒號、大小寫。
 * @param {string} text - AI 回傳的多行文字
 * @returns {object} { symbol: theme }
 */
function parseBatchResponse(text) {
  const result = {};
  const lines = text.split('\n');
  lines.forEach(line => {
    // 找第一個冒號（半形或全形）
    const idx = line.indexOf(':') !== -1 ? line.indexOf(':') : line.indexOf('：');
    if (idx === -1) return;

    let symbol = line.substring(0, idx).trim().toUpperCase();
    const theme = line.substring(idx + 1).trim();

    // 移除交易所前綴，如 "NASDAQ:NVDA" → "NVDA"
    if (symbol.includes(':')) symbol = symbol.split(':').pop();

    if (symbol && theme) result[symbol] = theme;
  });
  return result;
}

// ── Markdown 渲染器（唯一版本）────────────────────────────────
/**
 * 將 AI 回傳的 Markdown 渲染成 Google Sheets 格式。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 目標工作表
 * @param {string} text - 含 Markdown 的文字
 * @param {object} [colWidths] - 可選自訂欄寬 { c1, c2, c3, c4 }
 * @param {string} [headerColor] - 可選標題背景色（預設深藍）
 */
function renderMarkdownToSheet(sheet, text, colWidths, headerColor) {
  const widths = Object.assign({ c1: 160, c2: 120, c3: 430, c4: 180 }, colWidths || {});
  const hColor = headerColor || '#1a237e';

  sheet.setColumnWidth(1, widths.c1);
  sheet.setColumnWidth(2, widths.c2);
  sheet.setColumnWidth(3, widths.c3);
  sheet.setColumnWidth(4, widths.c4);

  const lines = text.split('\n');
  let row = 1;

  lines.forEach(rawLine => {
    const line = rawLine.trim();
    if (line === '') { row++; return; }

    if (line.startsWith('###')) {
      // 標題列
      const title = line.replace(/^#+\s*/, '');
      sheet.getRange(row, 1, 1, 4).merge()
        .setValue(title)
        .setBackground(hColor)
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(12)
        .setVerticalAlignment('middle');
      sheet.setRowHeight(row, 32);

    } else if (line.startsWith('|')) {
      // 表格列
      if (line.includes('---')) return; // 忽略分隔線
      const cells = line.split('|').filter(c => c.trim() !== '').map(c => c.trim());
      if (!cells.length) return;
      const range = sheet.getRange(row, 1, 1, cells.length);
      range.setValues([cells])
        .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID)
        .setWrap(true);
      // 表格標頭（緊接在 ### 標題後的第一行）
      if (row > 1 && sheet.getRange(row - 1, 1).getBackground() === hColor) {
        range.setBackground('#e8eaf6').setFontWeight('bold');
      }

    } else {
      // 一般文字、清單
      sheet.getRange(row, 1, 1, 4).merge()
        .setValue(line)
        .setWrap(true);
    }

    row++;
  });

  // 全域對齊
  sheet.getRange(1, 1, row, 4).setVerticalAlignment('middle');
}
