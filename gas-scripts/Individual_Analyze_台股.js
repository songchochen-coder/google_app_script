
/**
 * 腳本 1：優化版 - 批次抓取與 AI 分析
 */
function runIndividualAnalyze() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('台股存檔資料') || ss.insertSheet('台股存檔資料');
  sheet.clear();
  sheet.appendRow(['股票代碼', '股票名稱', '月漲幅%', 'AI 個股題材', '抓取時間']);

  Logger.log('🚀 正在從 TradingView 抓取台股高動能資料...');
  const fetchResult = fetchTradingViewData();
  
  if (fetchResult.error) {
    Logger.log('❌ 抓取失敗：' + fetchResult.error);
    return;
  }

  const stocks = fetchResult.data;
  const batchSize = 10; // 👈 關鍵優化：每 10 檔股票呼叫一次 AI
  const now = new Date();

  Logger.log(`✅ 找到 ${stocks.length} 檔股票，開始進行批次 AI 題材分析...`);

  for (let i = 0; i < stocks.length; i += batchSize) {
    const batch = stocks.slice(i, i + batchSize);
    const stockListStr = batch.map(s => `${s.name}(${s.symbol})`).join('\n');

    // --- 優化後的批次 Prompt ---
    const prompt = `
      # Role: 你是一位精通台股市場與產業供應鏈的資深研究員。
      # Task: 分析以下台股近期月漲幅 > 20% 的主要原因（如：營收成長、法說會利多、特定產業需求或政策支持）。
      # Data: 
      ${stockListStr}
      
      # Constraints:
      1. 使用【繁體中文】回答。
      2. 嚴格遵守格式，每行一檔，格式為：代號:分析內容。
      3. 分析內容請精簡在 30 字以內，直擊痛點。
      4. 請勿包含開場白或結尾文字。
    `;

    Logger.log(`正在分析第 ${i + 1} 到 ${Math.min(i + batchSize, stocks.length)} 檔股票...`);
    
    try {
      const response = callGemini(prompt);
      const analysisMap = parseBatchResponse(response);

      batch.forEach(stock => {
        const theme = analysisMap[stock.symbol] || "已完成分析 (請確認資料格式)";
        sheet.appendRow([stock.symbol, stock.name, stock.change, theme, now]);
      });
    } catch (e) {
      Logger.log(`批次處理異常: ${e.message}`);
    }

    // 批次處理呼叫次數少，間隔可縮短至 500ms
    Utilities.sleep(500); 
  }

  Logger.log('✨ 台股個股分析完成！資料已更新至「台股存檔資料」工作表。');
}

/**
 * 輔助函式：精準解析 AI 回傳內容
 */
function parseBatchResponse(text) {
  const lines = text.split('\n');
  const result = {};
  lines.forEach(line => {
    // 相容全形與半形冒號
    const splitIdx = line.indexOf(':') !== -1 ? line.indexOf(':') : line.indexOf('：');
    if (splitIdx !== -1) {
      const symbol = line.substring(0, splitIdx).trim();
      const theme = line.substring(splitIdx + 1).trim();
      result[symbol] = theme;
    }
  });
  return result;
}

/**
 * 呼叫 Gemini API
 */
function callGemini(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "temperature": 0.3, "maxOutputTokens": 800 } 
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const resText = response.getContentText();
  const json = JSON.parse(resText);
  
  if (json.candidates && json.candidates[0].content.parts[0].text) {
    return json.candidates[0].content.parts[0].text.trim();
  }
  return "";
}

/**
 * TradingView 台股篩選器 (保持原篩選邏輯)
 */
function fetchTradingViewData() {
  const url = 'https://scanner.tradingview.com/taiwan/scan';
  const payload = {
    "filter": [
      {"left": "Perf.1M", "operation": "greater", "right": 20},
      {"left": "market_cap_basic", "operation": "greater", "right": 5000000000},
      {"left": "average_volume_30d_calc", "operation": "greater", "right": 5000000},
      {"left": "type", "operation": "in_range", "right": ["stock", "dr", "fund"]}
    ],
    "options": {"lang": "zh_TW"},
    "markets": ["taiwan"],
    "columns": ["name", "description", "Perf.1M"],
    "sort": {"sortBy": "Perf.1M", "sortOrder": "desc"},
    "range": [0, 50] // 限制前 50 檔最強勢股，確保執行穩定
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    if (!data.data) return { error: 'TradingView 無回傳' };

    return { 
      data: data.data.map(item => ({
        symbol: item.s.split(':')[1], 
        name: item.d[1],
        change: item.d[2] ? item.d[2].toFixed(2) : "0.00"
      }))
    };
  } catch (e) {
    return { error: e.toString() };
  }
}