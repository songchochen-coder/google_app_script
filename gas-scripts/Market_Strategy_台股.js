/**
 * Market_Strategy_台股.js
 * ========================
 * 專業量化分析儀表板：從「台股存檔資料」讀取個股，進行題材聚類、板塊輪動、
 * 領頭羊與補漲股偵測，最終以專業 Dashboard 格式渲染於 Google Sheets。
 *
 * 共用函式（callGemini）定義於 Global_Config.js。
 */

function runMarketStrategy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('台股存檔資料');

  if (!sourceSheet) {
    Logger.log('⚠️ 找不到「台股存檔資料」工作表，請先執行台股個股分析。');
    try { SpreadsheetApp.getUi().alert('找不到「台股存檔資料」工作表，請先執行台股個股分析。'); } catch (e) { }
    return;
  }

  // 1. 讀取並前處理資料
  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getDisplayValues(); // 使用 DisplayValues 確保拿到文字
  const formulas = dataRange.getFormulas();  // 取得原本儲存的 HYPERLINK 公式
  if (data.length <= 1) {
    Logger.log('⚠️ 「台股存檔資料」中沒有足夠資料。');
    try { SpreadsheetApp.getUi().alert('「台股存檔資料」中沒有足夠資料。'); } catch (e) { }
    return;
  }

  const stocks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      let tvUrl = `https://www.tradingview.com/chart/?symbol=TWSE:${data[i][0]}`; // 預設值
      const formula = formulas[i][1]; // 第二欄是股票名稱（帶有 HYPERLINK）
      if (formula && formula.toUpperCase().includes('HYPERLINK(')) {
        const match = formula.match(/"(https?:\/\/[^"]+)"/);
        if (match) {
          tvUrl = match[1];
        }
      }

      stocks.push({
        symbol: data[i][0],
        name: data[i][1],
        theme: data[i][2],
        close: parseFloat(data[i][5]) || 0,
        change: parseFloat(data[i][6]) || 0,
        high20: parseFloat(data[i][15]) || 0,
        tvUrl: tvUrl
      });
    }
  }

  // 2. 進行 AI 深度分析 (要求一次回傳整合的 JSON，減少 AI 多次呼叫造成的落差)
  Logger.log('🚀 正在進行 AI 盤前市場策略與板塊分析...');
  const strategyJson = analyzeMarketAndSectors(stocks);

  // 3. 渲染 Google Sheets 儀表板
  Logger.log('🎨 正在繪製量化儀表板...');
  buildQuantDashboard(ss, strategyJson, stocks);

  try {
    SpreadsheetApp.getUi().alert('🎯 台股專業量化儀表板已生成！\n請查看「量化儀表板_台股」工作表。');
  } catch (e) {
    Logger.log('🎯 台股專業量化儀表板已生成！請查看「量化儀表板_台股」工作表。');
  }
}

// ==============================================================================
// AI 分析模組
// ==============================================================================

/**
 * Task 1: 題材聚類與板塊輪動分析
 */
// ==============================================================================
// AI 分析模組
// ==============================================================================

/**
 * Task 1: 盤前分析與板塊輪動綜合判斷
 */
function analyzeMarketAndSectors(stocks) {
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  // ── 量化情緒指標：程式自動計算，再供 AI 綜合評分 ─────────────────
  const total = stocks.length || 1;
  const limitUpCount = stocks.filter(s => s.change >= 9.9).length;
  const newHighCount = stocks.filter(s => s.high20 > 0 && s.close >= s.high20).length;
  const highVolCount = stocks.filter(s => s.relVol10d && s.relVol10d >= 1.5).length;
  const highRsiCount = stocks.filter(s => s.rsi && s.rsi >= 70).length;

  const sentimentMetrics =
    `【情緒客觀數據（程式計算，請綜合判斷給出1~10分）】\n` +
    `- 股票總數: ${total} 檔\n` +
    `- 漲停板: ${limitUpCount} 檔 (${((limitUpCount / total) * 100).toFixed(1)}%)\n` +
    `- 創30天以上新高: ${newHighCount} 檔 (${((newHighCount / total) * 100).toFixed(1)}%)\n` +
    `- 量能激增 (相對均量 > 1.5x): ${highVolCount} 檔 (${((highVolCount / total) * 100).toFixed(1)}%)\n` +
    `- RSI 超買區 (>70): ${highRsiCount} 檔 (${((highRsiCount / total) * 100).toFixed(1)}%)`;

  // ── 資料清單 ───────────────────────────────────────────────────────
  const stockListStr = stocks.map(s => {
    const isNewHigh = s.high20 > 0 && s.close >= s.high20;
    return `- ${s.name}(${s.symbol}): 漲幅 ${s.change}%, 題材: ${s.theme}${isNewHigh ? ', (🔥創高)' : ''}`;
  }).join('\n');

  const prompt = `你是頂尖的量化基金經理人，今天是 ${today}。
請根據以下高動能股票清單（月漲幅超過20%），自動進行【盤前交易策略與板塊分析】。

${sentimentMetrics}

【原始資料】
${stockListStr}

【任務要求】
1. 找出3~5個市場主流板塊
2. 每個板塊找出領頭羊
3. 找出補漲潛力股
4. 根據上方客觀數據，綜合判斷市場情緒並給出 1~10 分（1=極度恐懼, 5=中性, 10=極度貪婪）
5. 提出盤前操作策略
6. 根據清單內標記「創高」的股票，整理出一份特別觀察名單

【強制輸出結構】
你必須回傳純 JSON 格式，不要包含 \`\`\`json 等 Markdown 標記，直接依據以下 Schema 輸出：
{
  "sentiment_score": 7,
  "sentiment_label": "樂觀投機情緒",
  "market_view": "市場定調...（50字以內）",
  "risk_warning": "風險提示...（30字以內）",
  "strategies": ["核心策略一", "核心策略二", "核心策略三"],
  "sectors": [
    {
      "sector_name": "板塊名稱 (例如: AI伺服器供應鏈)",
      "momentum_score": 9,
      "analysis": "強勢板塊分析...資金流向判斷",
      "leader": {"symbol": "2330", "name": "台積電", "reason": "領頭羊理由"},
      "laggards": [
        {"symbol": "XXXX", "name": "XXX", "reason": "補漲潛力股理由"}
      ]
    }
  ],
  "new_high_stocks": ["2330 台積電", "3017 奇鋐"],
  "watchlist": ["2330 台積電", "2317 鴻海"]
}`;

  const jsonStr = callGeminiJSON(prompt);
  return parseJSONSafely(jsonStr);
}



// ==============================================================================
// 工具函式
// ==============================================================================

/**
 * 專門給要求 JSON 輸出的 API 呼叫，設定較低的 temperature 確保格式穩定
 */
function callGeminiJSON(prompt) {
  const result = callGemini(prompt, { temperature: 0.1, maxOutputTokens: 8192 });
  return result;
}

/**
 * 安全解析 JSON，過濾掉 AI 可能多嘴加上的 Markdown
 */
function parseJSONSafely(str) {
  try {
    let cleanStr = str.trim();
    if (cleanStr.startsWith('\`\`\`json')) {
      cleanStr = cleanStr.replace(/^\`\`\`json/, '').replace(/\`\`\`$/, '').trim();
    }
    return JSON.parse(cleanStr);
  } catch (e) {
    Logger.log("JSON 解析失敗，原始字串：" + str);
    return null;
  }
}

// ==============================================================================
// 視覺化儀表板渲染模組 (Dashboard Builder)
// ==============================================================================

// Runtime 快取：避免同一 symbol 重複查詢 TradingView
const _twExchangeCache = {};

/**
 * 取得台股 TradingView Chart URL。
 * 優先從 stocks 快取找；找不到時查詢 TradingView scanner 取得正確的
 * exchange 前綴 (TWSE 上市 or TPEX 上櫃)，解決上櫃股票連結失效問題。
 */
function getTvUrl(symbol, stocks) {
  const sym = String(symbol);

  // 1. 優先用已存的 tvUrl（從個股分析存檔取得，含正確 exchange）
  const stock = stocks.find(s => String(s.symbol) === sym);
  if (stock && stock.tvUrl) return stock.tvUrl;

  // 2. 已查詢過的 runtime 快取
  if (_twExchangeCache[sym]) return _twExchangeCache[sym];

  // 3. 查詢 TradingView API 取得正確 exchange (TWSE or TPEX)
  try {
    const resp = UrlFetchApp.fetch('https://scanner.tradingview.com/taiwan/scan', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        filter: [{ left: 'name', operation: 'equal', right: sym }],
        options: { lang: 'zh_TW' },
        markets: ['taiwan'],
        columns: ['name'],
        range: [0, 1]
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    if (data.data && data.data.length > 0) {
      const fullSymbol = data.data[0].s; // e.g. "TPEX:3081" or "TWSE:2330"
      const url = `https://www.tradingview.com/chart/?symbol=${fullSymbol}`;
      _twExchangeCache[sym] = url;
      return url;
    }
  } catch (e) {
    Logger.log(`getTvUrl lookup 失敗 (${sym}): ${e.message}`);
  }

  // 4. 最終 fallback（TWSE 猜測，大多數上市股票適用）
  return `https://www.tradingview.com/chart/?symbol=TWSE:${sym}`;
}

function buildQuantDashboard(ss, strategyJson, stocks) {
  const sheetName = '量化儀表板_台股';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(sheetName);

  // 1. 全域設定 (深色模式風格)
  sheet.getRange(1, 1, 100, 10).setBackground('#1E1E1E').setFontColor('#E0E0E0').setFontFamily('Arial');
  sheet.setColumnWidths(1, 10, 120);

  // 定義顏色常數
  const colors = {
    bg: '#1E1E1E',
    panelBg: '#2A2A2A',
    headerBg: '#0A3D62', // 深藍
    subHeaderBg: '#37474F',
    accentUp: '#4CAF50', // 綠漲
    accentWarn: '#FF9800', // 橘警
    textMain: '#FFFFFF',
    textSub: '#B0BEC5'
  };

  let currentRow = 2; // 從第 2 列開始畫

  // ── 標題 ──────────────────────────────────────────────────────────
  const titleRange = sheet.getRange(currentRow, 2, 1, 5);
  titleRange.merge()
    .setValue('台股量化盤前策略儀表板')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground(colors.headerBg)
    .setFontColor(colors.textMain)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 40);
  currentRow += 2;

  if (strategyJson) {
    // ── 0 情緒計量錶 (儀表板最頂端) ─────────────────────────────────
    const score = typeof strategyJson.sentiment_score === 'number' ? strategyJson.sentiment_score : 5;
    const label = strategyJson.sentiment_label || '中性';

    // 根據分數動態設定顏色與狀態文字
    let scoreBg, statusText;
    if (score >= 8) {
      scoreBg = '#B71C1C'; statusText = '🔥 極度貪婪 / 請等待反轉訊號';
    } else if (score >= 6) {
      scoreBg = '#E65100'; statusText = '🟠 樂觀投機 / 順勢持股';
    } else if (score >= 4) {
      scoreBg = '#37474F'; statusText = '⚪ 中性觀望 / 等待方向選擇';
    } else {
      scoreBg = '#0D47A1'; statusText = '🔵 恐懼情緒 / 逢低佈局機會';
    }

    // 大數字分數格
    sheet.getRange(currentRow, 2, 2, 2).merge()
      .setValue(score)
      .setFontSize(36).setFontWeight('bold')
      .setBackground(scoreBg).setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 35);
    sheet.setRowHeight(currentRow + 1, 35);

    // 標籤欄
    sheet.getRange(currentRow, 4).setValue('市場情緒指數').setFontColor(colors.textSub).setFontSize(10);
    sheet.getRange(currentRow + 1, 4).setValue(label).setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(currentRow, 5, 1, 2).merge().setValue(statusText)
      .setFontColor(scoreBg === '#37474F' ? colors.textMain : scoreBg).setFontWeight('bold');
    sheet.getRange(currentRow + 1, 5, 1, 2).merge().setValue('/ 10 分（1=極度恐懼  10=極度貪婪）')
      .setFontColor(colors.textSub).setFontSize(10);
    currentRow += 3;

    // ── 1 市場定調 & 2 風險提示 ───────────────────────────────────────
    sheet.getRange(currentRow, 2).setValue('大盤情緒：').setFontWeight('bold').setFontColor(colors.textSub);
    sheet.getRange(currentRow, 3, 1, 4).merge().setValue(strategyJson.market_view).setFontColor(colors.textMain);
    currentRow++;

    sheet.getRange(currentRow, 2).setValue('風險提示：').setFontWeight('bold').setFontColor(colors.accentWarn);
    sheet.getRange(currentRow, 3, 1, 4).merge().setValue(strategyJson.risk_warning).setFontColor(colors.accentWarn);
    currentRow += 2;

    // ── 3 三個核心策略 ───────────────────────────────────────────────
    sheet.getRange(currentRow, 2, 1, 5).merge().setValue('盤前核心策略').setFontWeight('bold').setBackground(colors.subHeaderBg);
    currentRow++;
    const plans = strategyJson.strategies || [];
    plans.forEach((plan, idx) => {
      sheet.getRange(currentRow, 2).setValue(`策略 ${idx + 1}`);
      sheet.getRange(currentRow, 3, 1, 4).merge().setValue(plan).setWrap(true);
      currentRow++;
    });
    currentRow++;

    // ── 4 強勢板塊分析 & 5 領頭羊與補漲股 ─────────────────────────────
    sheet.getRange(currentRow, 2, 1, 5).merge()
      .setValue('🔥 主流板塊與輪動資金追蹤')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground(colors.subHeaderBg)
      .setFontColor(colors.textMain);
    currentRow++;

    if (strategyJson.sectors) {
      strategyJson.sectors.sort((a, b) => b.momentum_score - a.momentum_score).forEach(sector => {
        // 板塊標題
        sheet.getRange(currentRow, 2, 1, 5).merge().setValue(`【 ${sector.sector_name} 】 動能: ${sector.momentum_score}/10`).setFontColor('#81D4FA').setFontWeight('bold');
        currentRow++;

        // 板塊分析
        sheet.getRange(currentRow, 2).setValue('分析').setFontColor(colors.textSub);
        sheet.getRange(currentRow, 3, 1, 4).merge().setValue(sector.analysis).setWrap(true);
        currentRow++;

        // 領頭羊
        if (sector.leader) {
          const url = getTvUrl(sector.leader.symbol, stocks);
          const nameLink = `=HYPERLINK("${url}", "${sector.leader.name} (${sector.leader.symbol})")`;
          sheet.getRange(currentRow, 2).setValue('👑 領頭羊').setFontColor('#FFD700');
          sheet.getRange(currentRow, 3).setValue(nameLink).setFontWeight('bold');
          sheet.getRange(currentRow, 4, 1, 3).merge().setValue(sector.leader.reason).setWrap(true);
          currentRow++;
        }

        // 補漲股
        if (sector.laggards && sector.laggards.length > 0) {
          sector.laggards.forEach((laggard, idx) => {
            const url = getTvUrl(laggard.symbol, stocks);
            const nameLink = `=HYPERLINK("${url}", "${laggard.name} (${laggard.symbol})")`;
            sheet.getRange(currentRow, 2).setValue(idx === 0 ? '🚀 補漲/外溢' : '').setFontColor(colors.accentUp);
            sheet.getRange(currentRow, 3).setValue(nameLink);
            sheet.getRange(currentRow, 4, 1, 3).merge().setValue(laggard.reason).setWrap(true).setFontColor(colors.textSub);
            currentRow++;
          });
        }
        currentRow++; // 板塊間留白
      });
    }

    // ── 6 創30天以上新高股票 & 7 今日觀察名單 ─────────────────────────────
    // 新增創高股區塊
    sheet.getRange(currentRow, 2, 1, 5).merge().setValue('📈 創30天以上新高股票').setFontWeight('bold').setBackground(colors.subHeaderBg);
    currentRow++;
    const newHighList = strategyJson.new_high_stocks || [];
    if (newHighList.length === 0) {
      sheet.getRange(currentRow, 2, 1, 5).merge().setValue('無').setFontColor('#FF9800');
      currentRow++;
    } else {
      newHighList.forEach((entry, idx) => {
        const parts = String(entry).trim().split(/\s+/);
        const sym = parts[0];    // e.g. "2330"
        const label = parts.length > 1 ? parts.slice(1).join(' ') : sym;  // e.g. "台積電"
        const url = getTvUrl(sym, stocks);
        const nameLink = `=HYPERLINK("${url}","${sym} ${label}")`;
        sheet.getRange(currentRow, 2).setValue(idx === 0 ? '🔥創高' : '').setFontColor('#FF9800').setFontWeight('bold');
        sheet.getRange(currentRow, 3, 1, 4).merge().setValue(nameLink).setFontColor('#FF9800');
        currentRow++;
      });
    }
    currentRow++;

    sheet.getRange(currentRow, 2, 1, 5).merge().setValue('📌 今日精選觀察名單').setFontWeight('bold').setBackground(colors.subHeaderBg);
    currentRow++;
    const watchList = strategyJson.watchlist || [];
    if (watchList.length === 0) {
      sheet.getRange(currentRow, 2, 1, 5).merge().setValue('無').setFontColor(colors.accentUp);
      currentRow++;
    } else {
      watchList.forEach((entry, idx) => {
        const parts = String(entry).trim().split(/\s+/);
        const sym = parts[0];
        const label = parts.length > 1 ? parts.slice(1).join(' ') : sym;
        const url = getTvUrl(sym, stocks);
        const nameLink = `=HYPERLINK("${url}","${sym} ${label}")`;
        sheet.getRange(currentRow, 2).setValue(idx === 0 ? '📌觀察' : '').setFontColor(colors.accentUp).setFontWeight('bold');
        sheet.getRange(currentRow, 3, 1, 4).merge().setValue(nameLink).setFontColor(colors.accentUp);
        currentRow++;
      });
    }
    currentRow++;
  }

  // ── 加上外框與收尾排版 ──────────────────────────────────────────
  const finalRow = currentRow - 1;
  sheet.getRange(2, 2, finalRow - 1, 5).setBorder(true, true, true, true, false, false, '#444444', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 隱藏不需要的欄與列，使其看起來像乾淨的儀表板
  sheet.hideColumns(1);
  sheet.hideColumns(7, sheet.getMaxColumns() - 6);
  if (sheet.getMaxRows() > finalRow + 5) {
    sheet.hideRows(finalRow + 5, sheet.getMaxRows() - finalRow - 4);
  }
}