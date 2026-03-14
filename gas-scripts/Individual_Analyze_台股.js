/**
 * Individual_Analyze_台股.js
 * ==========================
 * 從 TradingView 抓取台股高動能資料，結合鉅亨網最新新聞，批次呼叫 AI 分析個股題材。
 * 共用函式（callGemini, parseBatchResponse）定義於 Global_Config.js。
 */

// ── 主程式 ────────────────────────────────────────────────────
function runIndividualAnalyze() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('台股存檔資料') || ss.insertSheet('台股存檔資料');
  sheet.clear();

  // 擴充表頭 15 欄
  const headers = [
    '股票代碼', '股票名稱', '產業板塊', '細分產業',
    '最新收盤價', '單日漲幅%', '五日漲幅%', '月漲幅%',
    '成交量', '成交額(元)', '10日相對成交量', '5日均量',
    '市值(元)', 'RSI', '20日新高', 'AI 個股題材', '抓取時間'
  ];
  sheet.appendRow(headers);

  Logger.log('🚀 正在從 TradingView 抓取台股高動能資料...');
  const fetchResult = fetchTradingViewData();

  if (fetchResult.error) {
    Logger.log('❌ 抓取失敗：' + fetchResult.error);
    sheet.appendRow(['ERROR', fetchResult.error, '', '', new Date()]);
    return;
  }

  // 抓取鉅亨網最新台股新聞（作為 AI 上下文）
  Logger.log('📰 正在抓取鉅亨網最新新聞...');
  const newsContext = fetchCnyesNews();
  if (newsContext) {
    Logger.log('✅ 已取得最新新聞作為分析依據');
  }

  const stocks = fetchResult.data;
  const batchSize = 5;
  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy年MM月dd日');
  Logger.log(`✅ 找到 ${stocks.length} 檔股票，開始批次 AI 題材分析...`);

  for (let i = 0; i < stocks.length; i += batchSize) {
    const batch = stocks.slice(i, i + batchSize);
    const stockListStr = batch.map(s => `${s.name}(${s.symbol})`).join('\n');

    const prompt = buildPrompt(today, stockListStr, batch.length, newsContext);

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

          // 判斷是否紅底漲停 (簡單判斷單日 > 9.9%)
          const finalNameLink = stock.change >= 9.9 ? `="🔥 " & ${nameLink}` : nameLink;

          sheet.appendRow([
            stock.symbol,             // ticker
            finalNameLink,            // name (hyperlink)
            stock.sector,             // sector
            stock.industry,           // industry
            stock.close,              // close
            stock.change,             // change
            stock.perf5d,             // Perf.5D
            stock.perf1m,             // Perf.1M
            stock.volume,             // volume
            stock.valueTraded,        // Value.Traded
            stock.relVol10d,          // relative_volume_10d_calc
            stock.avgVol5d,           // average_volume_5d
            stock.marketCap,          // market_cap_basic
            stock.rsi,                // RSI
            stock.high20,             // High.20
            theme,                    // AI 個股題材
            now                       // 抓取時間
          ]);
        } else {
          unmatched.push(stock);
        }
      });

      // ── 對未匹配的股票單獨重試 ──
      if (unmatched.length > 0) {
        Logger.log(`⚠️ ${unmatched.length} 檔未匹配，逐一重試：${unmatched.map(s => s.symbol).join(', ')}`);
        unmatched.forEach(stock => {
          const retryPrompt = `你是台股研究員，今天 ${today}。
請用一句話（30字以內）說明台股 ${stock.name}(${stock.symbol}) 近期漲幅超過20%的核心原因。
格式：${stock.symbol}:原因`;
          const retryResp = callGemini(retryPrompt, { temperature: 0.2, maxOutputTokens: 200 });
          const retryMap = parseBatchResponse(retryResp, [stock]);
          const theme = retryMap[stock.symbol] || retryResp.trim() || '無法取得分析';
          const tvUrl = `https://www.tradingview.com/chart/?symbol=${stock.exchange}:${stock.symbol}`;
          const nameLink = `=HYPERLINK("${tvUrl}","${stock.name}")`;

          const finalNameLink = stock.change >= 9.9 ? `="🔥 " & ${nameLink}` : nameLink;

          sheet.appendRow([
            stock.symbol, finalNameLink, stock.sector, stock.industry,
            stock.close, stock.change, stock.perf5d, stock.perf1m,
            stock.volume, stock.valueTraded, stock.relVol10d, stock.avgVol5d,
            stock.marketCap, stock.rsi, stock.high20, theme, now
          ]);
          Utilities.sleep(300);
        });
      }

    } catch (e) {
      Logger.log(`批次處理異常: ${e.message}`);
      // 異常時仍寫入佔位，不漏行
      batch.forEach(stock => {
        sheet.appendRow([
          stock.symbol, stock.name, stock.sector, stock.industry,
          stock.close, stock.change, stock.perf5d, stock.perf1m,
          stock.volume, stock.valueTraded, stock.relVol10d, stock.avgVol5d,
          stock.marketCap, stock.rsi, stock.high20, '分析異常', now
        ]);
      });
    }

    Utilities.sleep(500);
  }

  Logger.log('✨ 台股個股分析完成！資料已更新至「台股存檔資料」工作表。');
}

// ── 建構 Prompt ──────────────────────────────────────────────
function buildPrompt(today, stockListStr, count, newsContext) {
  const newsSection = newsContext
    ? `\n【今日最新台股新聞（請優先參考）】\n${newsContext}\n`
    : '';

  return `你是台股市場研究員，今天是 ${today}。
請根據下方最新新聞與你的知識，分析以下台股近期月漲幅超過20%的主要驅動題材。
${newsSection}
【待分析股票】
${stockListStr}

輸出規定：
- 每行一檔，格式：股票代號:分析內容（例如：6217:AI伺服器連接器出貨強勁，Q1訂單能見度高）
- 代號用純數字，不加括號或前綴
- 分析30字以內，聚焦最新事件（法說會/財報/訂單/政策）
- 不要開場白，直接輸出全部 ${count} 檔`;
}

// ── 鉅亨網 RSS 抓取（台股焦點新聞）────────────────────────────
/**
 * 抓取鉅亨網台股頻道 RSS，回傳最新 N 則標題作為字串
 * @param {number} [maxItems=15] 最多取幾則
 * @returns {string} 新聞標題字串，失敗則回傳空字串
 */
function fetchCnyesNews(maxItems) {
  maxItems = maxItems || 15;
  // 鉅亨網已停用 RSS，改用 Google News 台股即時新聞
  const rssUrl = 'https://news.google.com/rss/search?q=%E5%8F%B0%E8%82%A1&hl=zh-TW&gl=TW&ceid=TW:zh-Hant';

  try {
    const response = UrlFetchApp.fetch(rssUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('⚠️ 鉅亨網 RSS 回應碼：' + response.getResponseCode());
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
    Logger.log('⚠️ 鉅亨網新聞抓取失敗：' + e.toString());
    return '';
  }
}

// ── TradingView 台股篩選器 ────────────────────────────────────
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
    columns: [
      'name', 'description', 'sector', 'industry',
      'close', 'change', 'Perf.5D', 'Perf.1M',
      'volume', 'Value.Traded', 'relative_volume_10d_calc', 'average_volume_10d_calc',
      'market_cap_basic', 'RSI', 'High.All'
    ],
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
        exchange: item.s.split(':')[0],
        symbol: item.s.split(':')[1],
        name: item.d[1],
        sector: item.d[2] || '',
        industry: item.d[3] || '',
        close: item.d[4],
        change: item.d[5] ? parseFloat(item.d[5].toFixed(2)) : 0,
        perf5d: item.d[6] ? parseFloat(item.d[6].toFixed(2)) : 0,
        perf1m: item.d[7] ? parseFloat(item.d[7].toFixed(2)) : 0,
        volume: item.d[8] || 0,
        valueTraded: item.d[9] || 0,
        relVol10d: item.d[10] ? parseFloat(item.d[10].toFixed(2)) : 0,
        avgVol5d: item.d[11] || 0,
        marketCap: item.d[12] || 0,
        rsi: item.d[13] ? parseFloat(item.d[13].toFixed(2)) : 0,
        high20: item.d[14] ? parseFloat(item.d[14].toFixed(2)) : 0 // 借用 High.All 代替
      }))
    };
  } catch (e) {
    return { error: e.toString() };
  }
}