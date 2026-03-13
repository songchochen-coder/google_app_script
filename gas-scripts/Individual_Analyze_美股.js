
/**
 * 主程式：批次抓取美股並分析
 */
function runUSIndividualAnalyze() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('美股存檔資料') || ss.insertSheet('美股存檔資料');
  sheet.clear();
  sheet.appendRow(['股票代碼', '股票名稱', '月漲幅%', 'AI 個股題材', '抓取時間']);

  Logger.log('🚀 正在從 TradingView 抓取美股數據...');
  const fetchResult = fetchTradingViewUSData();
  
  if (fetchResult.error) {
    Logger.log('❌ 抓取失敗：' + fetchResult.error);
    return;
  }

  const stocks = fetchResult.data;
  const batchSize = 10; // 每一組處理 10 檔
  const now = new Date();

  Logger.log(`✅ 找到 ${stocks.length} 檔股票，開始分批分析...`);

  for (let i = 0; i < stocks.length; i += batchSize) {
    const batch = stocks.slice(i, i + batchSize);
    const stockListStr = batch.map(s => `${s.name} (${s.symbol})`).join('\n');

    // --- 優化後的 Prompt ---
    const prompt = `
      # Role: 你是一位精通美股與全球產業鏈的資深投資分析師。
      # Task: 請分析以下美股近期漲幅超過 20% 的核心驅動因素（如：財報、產業政策、技術突破或資金外溢）。
      # Data: 
      ${stockListStr}
      
      # Constraints:
      1. 請用【繁體中文】回答。
      2. 嚴格遵守格式，每行一檔，格式為：代碼:分析內容。
      3. 分析內容請控制在 30 字以內，直擊重點。
      4. 不要任何開場白、不要結尾，直接輸出內容。
      5. 若涉及台股供應鏈（如 AI Server 帶動散熱），請簡短提及。
    `;

    Logger.log(`正在處理第 ${i + 1} 到 ${Math.min(i + batchSize, stocks.length)} 檔...`);
    
    try {
      const response = callGemini(prompt);
      const analysisMap = parseBatchResponse(response);

      // 將分析結果填入試算表
      batch.forEach(stock => {
        const theme = analysisMap[stock.symbol] || "分析完成 (請手動確認資料)";
        sheet.appendRow([stock.symbol, stock.name, stock.change, theme, now]);
      });
    } catch (e) {
      Logger.log(`批次分析錯誤: ${e.message}`);
    }

    // 因為是批次處理，呼叫次數少，間隔可以縮短，避免逾時
    Utilities.sleep(500); 
  }

  Logger.log('✨ 美股個股分析完成！');
}

/**
 * 輔助函式：解析 AI 回傳的批次文字
 * 增加對多種分隔符號的相容性與修剪
 */
function parseBatchResponse(text) {
  const lines = text.split('\n');
  const result = {};
  lines.forEach(line => {
    // 尋找第一個冒號 (相容全形與半形)
    const separatorIndex = line.indexOf(':') !== -1 ? line.indexOf(':') : line.indexOf('：');
    if (separatorIndex !== -1) {
      const symbol = line.substring(0, separatorIndex).trim().toUpperCase();
      const theme = line.substring(separatorIndex + 1).trim();
      // 處理可能帶有前綴或交易所代碼的情況 (如 "NASDAQ:NVDA" -> "NVDA")
      const cleanSymbol = symbol.includes(':') ? symbol.split(':').pop() : symbol;
      result[cleanSymbol] = theme;
    }
  });
  return result;
}

/**
 * 呼叫 Gemini 3 Flash API
 */
function callGemini(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "maxOutputTokens": 1000, "temperature": 0.4 } // 降低隨機性
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  
  if (json.candidates && json.candidates[0].content.parts[0].text) {
    return json.candidates[0].content.parts[0].text.trim();
  }
  return "";
}

/**
 * TradingView 美股數據抓取 (保持不變)
 */
function fetchTradingViewUSData() {
  const url = 'https://scanner.tradingview.com/america/scan';
  const payload = {
    "filter": [
      {"left": "Perf.1M", "operation": "greater", "right": 20},
      {"left": "market_cap_basic", "operation": "greater", "right": 1000000000},
      {"left": "average_volume_10d_calc", "operation": "greater", "right": 1000000}
    ],
    "options": {"lang": "en"},
    "markets": ["america"],
    "columns": ["name", "description", "Perf.1M"],
    "sort": {"sortBy": "Perf.1M", "sortOrder": "desc"},
    "range": [0, 40] 
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
    if (!data.data) return { error: '無資料' };
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