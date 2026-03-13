/**
 * Individual_Analyze_台股.js
 * ==========================
 * 從 TradingView 抓取台股高動能資料，批次呼叫 AI 分析個股題材。
 * 共用函式（callGemini, parseBatchResponse）定義於 Global_Config.js。
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
    sheet.appendRow(['ERROR', fetchResult.error, '', '', new Date()]);
    return;
  }

  const stocks = fetchResult.data;
  const batchSize = 5; // 降低批次大小，減少截斷風險
  const now = new Date();
  Logger.log(`✅ 找到 ${stocks.length} 檔股票，開始批次 AI 題材分析...`);

  for (let i = 0; i < stocks.length; i += batchSize) {
    const batch = stocks.slice(i, i + batchSize);
    const stockListStr = batch.map(s => `${s.name}(${s.symbol})`).join('\n');

    const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy年MM月dd日');
    const prompt = `你是台股市場研究員，今天是 ${today}。
請根據你所知最新的市場資訊（2025年至今），分析以下台股近期月漲幅超過20%的主要驅動題材。

${stockListStr}

輸出規定：
- 每行一檔股票，格式為 股票代號:分析內容（例如：6217:AI伺服器連接器需求強勁，Q1訂單能見度高）
- 代號用純數字，不加括號或前綴
- 分析內容30字以內，聚焦最新事件（法說會、財報、訂單、政策）
- 不要開場白，直接輸出所有 ${batch.length} 檔的分析`;

    Logger.log(`正在分析第 ${i + 1} 到 ${Math.min(i + batchSize, stocks.length)} 檔股票...`);

    try {
      const response = callGemini(prompt, { temperature: 0.2, maxOutputTokens: 3000 });
      Logger.log(`=== 批次 ${i + 1}-${Math.min(i + batchSize, stocks.length)} AI 原始回傳 ===\n${response}\n===`);
      const analysisMap = parseBatchResponse(response, batch); // 傳入 batch 支援名稱反查

      batch.forEach(stock => {
        const theme = analysisMap[stock.symbol] || '已完成分析 (請確認資料格式)';
        sheet.appendRow([stock.symbol, stock.name, stock.change, theme, now]);
      });
    } catch (e) {
      Logger.log(`批次處理異常: ${e.message}`);
    }

    Utilities.sleep(500);
  }

  Logger.log('✨ 台股個股分析完成！資料已更新至「台股存檔資料」工作表。');
}

/**
 * TradingView 台股篩選器
 */
function fetchTradingViewData() {
  const url = 'https://scanner.tradingview.com/taiwan/scan';
  const payload = {
    filter: [
      { left: 'Perf.1M', operation: 'greater', right: 20 },
      { left: 'market_cap_basic', operation: 'greater', right: 5000000000 },
      { left: 'average_volume_30d_calc', operation: 'greater', right: 5000000 },
      { left: 'type', operation: 'in_range', right: ['stock', 'dr', 'fund'] }
    ],
    options: { lang: 'zh_TW' },
    markets: ['taiwan'],
    columns: ['name', 'description', 'Perf.1M'],
    sort: { sortBy: 'Perf.1M', sortOrder: 'desc' },
    range: [0, 50]
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