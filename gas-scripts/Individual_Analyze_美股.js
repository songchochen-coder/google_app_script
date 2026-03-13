/**
 * Individual_Analyze_美股.js
 * ==========================
 * 從 TradingView 抓取美股高動能資料，批次呼叫 AI 分析個股題材。
 * 共用函式（callGemini, parseBatchResponse）定義於 Global_Config.js。
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
    sheet.appendRow(['ERROR', fetchResult.error, '', '', new Date()]);
    return;
  }

  const stocks = fetchResult.data;
  const batchSize = 10;
  const now = new Date();
  Logger.log(`✅ 找到 ${stocks.length} 檔股票，開始分批分析...`);

  for (let i = 0; i < stocks.length; i += batchSize) {
    const batch = stocks.slice(i, i + batchSize);
    const stockListStr = batch.map(s => `${s.name} (${s.symbol})`).join('\n');

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
      const response = callGemini(prompt, { temperature: 0.4, maxOutputTokens: 1000 });
      const analysisMap = parseBatchResponse(response);

      batch.forEach(stock => {
        const theme = analysisMap[stock.symbol] || '分析完成 (請手動確認資料)';
        sheet.appendRow([stock.symbol, stock.name, stock.change, theme, now]);
      });
    } catch (e) {
      Logger.log(`批次分析錯誤: ${e.message}`);
    }

    Utilities.sleep(500);
  }

  Logger.log('✨ 美股個股分析完成！');
}

/**
 * TradingView 美股篩選器
 */
function fetchTradingViewUSData() {
  const url = 'https://scanner.tradingview.com/america/scan';
  const payload = {
    filter: [
      { left: 'Perf.1M', operation: 'greater', right: 20 },
      { left: 'market_cap_basic', operation: 'greater', right: 1000000000 },
      { left: 'average_volume_10d_calc', operation: 'greater', right: 1000000 }
    ],
    options: { lang: 'en' },
    markets: ['america'],
    columns: ['name', 'description', 'Perf.1M'],
    sort: { sortBy: 'Perf.1M', sortOrder: 'desc' },
    range: [0, 40]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    if (!data.data) return { error: 'TradingView 無回傳資料' };

    return {
      data: data.data.map(item => ({
        symbol: item.s.split(':')[1],
        name: item.d[1],
        change: item.d[2] ? item.d[2].toFixed(2) : '0.00'
      }))
    };
  } catch (e) {
    return { error: e.toString() };
  }
}