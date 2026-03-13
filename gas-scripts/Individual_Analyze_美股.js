/**
 * Individual_Analyze_美股.js
 * ==========================
 * 從 TradingView 抓取美股高動能資料，結合 Google News 最新新聞，批次呼叫 AI 分析個股題材。
 * 共用函式（callGemini, parseBatchResponse）定義於 Global_Config.js。
 */

// ── 主程式 ────────────────────────────────────────────────────
function runUSIndividualAnalyze() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('美股存檔資料') || ss.insertSheet('美股存檔資料');
  sheet.clear();
  sheet.appendRow(['股票代碼', '股票名稱', '月漲幅%', 'AI 個股題材', '抓取時間']);

  Logger.log('🚀 正在從 TradingView 抓取美股高動能資料...');
  const fetchResult = fetchTradingViewUSData();

  if (fetchResult.error) {
    Logger.log('❌ 抓取失敗：' + fetchResult.error);
    sheet.appendRow(['ERROR', fetchResult.error, '', '', new Date()]);
    return;
  }

  // 抓取 Google News 最新美股新聞（作為 AI 上下文）
  Logger.log('📰 正在抓取最新美股新聞...');
  const newsContext = fetchUSNews();
  if (newsContext) {
    Logger.log('✅ 已取得最新新聞作為分析依據');
  }

  const stocks = fetchResult.data;
  const batchSize = 5; // 降低批次大小，減少截斷風險
  const now = new Date();
  const today = Utilities.formatDate(now, 'America/New_York', 'yyyy年MM月dd日');
  Logger.log(`✅ 找到 ${stocks.length} 檔股票，開始批次 AI 題材分析...`);

  for (let i = 0; i < stocks.length; i += batchSize) {
    const batch = stocks.slice(i, i + batchSize);
    const stockListStr = batch.map(s => `${s.name}(${s.symbol})`).join('\n');

    const prompt = buildUSPrompt(today, stockListStr, batch.length, newsContext);

    Logger.log(`正在分析第 ${i + 1} 到 ${Math.min(i + batchSize, stocks.length)} 檔股票...`);

    try {
      const response = callGemini(prompt, { temperature: 0.2, maxOutputTokens: 3000 });
      Logger.log(`=== 批次 ${i + 1}-${Math.min(i + batchSize, stocks.length)} AI 回傳 ===\n${response}\n===`);
      const analysisMap = parseBatchResponse(response, batch);

      // ── 寫入工作表，並收集未匹配的股票 ──
      const unmatched = [];
      batch.forEach(stock => {
        const theme = analysisMap[stock.symbol];
        if (theme) {
          const tvUrl = `https://www.tradingview.com/chart/?symbol=${stock.exchange}:${stock.symbol}`;
          const nameLink = `=HYPERLINK("${tvUrl}","${stock.name}")`;
          sheet.appendRow([stock.symbol, nameLink, stock.change, theme, now]);
        } else {
          unmatched.push(stock);
        }
      });

      // ── 對未匹配的股票單獨重試 ──
      if (unmatched.length > 0) {
        Logger.log(`⚠️ ${unmatched.length} 檔未匹配，逐一重試：${unmatched.map(s => s.symbol).join(', ')}`);
        unmatched.forEach(stock => {
          const retryPrompt = `You are a US stock market analyst. Today is ${today}.
Briefly explain in Traditional Chinese (30 chars or less) why ${stock.name}(${stock.symbol}) surged over 20% this month.
Format: ${stock.symbol}:reason`;
          const retryResp = callGemini(retryPrompt, { temperature: 0.2, maxOutputTokens: 200 });
          const retryMap = parseBatchResponse(retryResp, [stock]);
          const theme = retryMap[stock.symbol] || retryResp.trim() || '無法取得分析';
          const tvUrl = `https://www.tradingview.com/chart/?symbol=${stock.exchange}:${stock.symbol}`;
          const nameLink = `=HYPERLINK("${tvUrl}","${stock.name}")`;
          sheet.appendRow([stock.symbol, nameLink, stock.change, theme, now]);
          Utilities.sleep(300);
        });
      }

    } catch (e) {
      Logger.log(`批次處理異常: ${e.message}`);
      // 異常時仍寫入佔位，不漏行
      batch.forEach(stock => sheet.appendRow([stock.symbol, stock.name, stock.change, '分析異常', now]));
    }

    Utilities.sleep(500);
  }

  Logger.log('✨ 美股個股分析完成！資料已更新至「美股存檔資料」工作表。');
}

// ── 建構 Prompt ──────────────────────────────────────────────
function buildUSPrompt(today, stockListStr, count, newsContext) {
  const newsSection = newsContext
    ? `\n【最新美股新聞（請優先參考）】\n${newsContext}\n`
    : '';

  return `你是美股市場研究員，今天是 ${today}。
請根據下方最新新聞與你的知識，分析以下美股近期月漲幅超過20%的主要驅動題材。
${newsSection}
【待分析股票】
${stockListStr}

輸出規定：
- 每行一檔，格式：股票代號:分析內容（例如：NVDA:AI晶片需求爆發，數據中心訂單強勁）
- 代號用原始英文代號，不加括號
- 分析30字以內，聚焦最新事件（財報/法說/政策/產業趨勢）
- 不要開場白，直接輸出全部 ${count} 檔`;
}

// ── Google News 美股新聞抓取 ─────────────────────────────────
/**
 * 抓取 Google News 美股搜尋 RSS，回傳最新 N 則標題作為字串
 * @param {number} [maxItems=15] 最多取幾則
 * @returns {string} 新聞標題字串，失敗則回傳空字串
 */
function fetchUSNews(maxItems) {
  maxItems = maxItems || 15;
  const rssUrl = 'https://news.google.com/rss/search?q=US+stock+market&hl=en-US&gl=US&ceid=US:en';

  try {
    const response = UrlFetchApp.fetch(rssUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('⚠️ Google News 回應碼：' + response.getResponseCode());
      return '';
    }

    const xml = response.getContentText();
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const channel = root.getChild('channel');
    const items = channel.getChildren('item');

    const headlines = [];
    for (let i = 0; i < Math.min(items.length, maxItems); i++) {
      const title = items[i].getChildText('title');
      const pubDate = items[i].getChildText('pubDate');
      if (title) headlines.push(`- ${title}${pubDate ? '（' + pubDate.substring(0, 16) + '）' : ''}`);
    }

    return headlines.join('\n');
  } catch (e) {
    Logger.log('⚠️ 美股新聞抓取失敗：' + e.toString());
    return '';
  }
}

// ── TradingView 美股篩選器 ────────────────────────────────────
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
        exchange: item.s.split(':')[0],  // NASDAQ, NYSE, AMEX 等
        symbol: item.s.split(':')[1],
        name: item.d[1],
        change: item.d[2] ? item.d[2].toFixed(2) : '0.00'
      }))
    };
  } catch (e) {
    return { error: e.toString() };
  }
}