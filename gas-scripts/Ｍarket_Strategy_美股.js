/**
 * Market_Strategy_美股.js
 * ========================
 * 從「美股存檔資料」讀取高動能個股，呼叫 AI 產生包含「台美連動」的美股盤前策略，
 * 並渲染成與台股相同的專業六大區塊 Google Sheets 儀表板。
 * 
 * 依賴：
 * 1. Global_Config.js: callGemini
 * 2. Market_Strategy_台股.js: parseJSONSafely, getTvUrl (共用解析與抓取網址邏輯)
 */

function runUSMarketStrategy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('美股存檔資料');

  if (!sourceSheet) {
    Logger.log('⚠️ 找不到「美股存檔資料」工作表，請先執行美股個股分析。');
    try { SpreadsheetApp.getUi().alert('找不到「美股存檔資料」工作表，請先執行美股個股分析。'); } catch (e) { }
    return;
  }

  // 1. 讀取並前處理資料
  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getDisplayValues(); // 使用 DisplayValues 確保拿到文字
  const formulas = dataRange.getFormulas();  // 取得原本儲存的 HYPERLINK 公式

  if (data.length <= 1) {
    Logger.log('⚠️ 「美股存檔資料」中沒有足夠資料。');
    try { SpreadsheetApp.getUi().alert('「美股存檔資料」中沒有足夠資料。'); } catch (e) { }
    return;
  }

  const stocks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      let tvUrl = `https://www.tradingview.com/chart/?symbol=${data[i][0]}`; // 預設值
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

  // 2. 進行 AI 深度分析 (要求一次回傳整合的 JSON，包含台股連動)
  Logger.log('🚀 正在進行 AI 美股盤前市場策略與台美連動分析...');
  const strategyJson = analyzeUSMarketAndSectors(stocks);

  if (!strategyJson) {
    Logger.log('❌ AI 分析回傳空白或解析失敗，無法生成報告。');
    return;
  }

  // 3. 渲染 Google Sheets 儀表板
  Logger.log('🎨 正在繪製美股量化儀表板...');
  buildUSQuantDashboard(ss, strategyJson, stocks);

  // 4. 渲染板塊分類工作表（依 AI 強勢板塊分組、依月漲幅排序）
  Logger.log('📊 正在渲染美股板塊分類表...');
  buildUSSectorBreakdown(ss, strategyJson, stocks);

  try {
    SpreadsheetApp.getUi().alert('🎯 美股專業量化儀表板（含台股連動）已生成！\n請查看「量化儀表板_美股」與「板塊分類_美股」工作表。');
  } catch (e) {
    Logger.log('🎯 美股專業量化儀表板已生成！');
  }
}

// ==============================================================================
// AI 分析模組
// ==============================================================================

/**
 * Task: 美股盤前分析、板塊輪動與台台美股連動綜合判斷
 */
function analyzeUSMarketAndSectors(stocks) {
  const today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');

  // ── 量化情緒指標：程式自動計算，再供 AI 綜合評分 ─────────────────
  const total = stocks.length || 1;
  const bigMovers = stocks.filter(s => s.change >= 9.9).length;
  const newHighCount = stocks.filter(s => s.high20 > 0 && s.close >= s.high20).length;
  const highVolCount = stocks.filter(s => s.relVol10d && s.relVol10d >= 1.5).length;
  const highRsiCount = stocks.filter(s => s.rsi && s.rsi >= 70).length;

  const sentimentMetrics =
    `【情緒客觀數據（程式計算，請綜合判斷給出1~10分）】\n` +
    `- 股票總數: ${total} 檔\n` +
    `- 單日大漲 (>9.9%): ${bigMovers} 檔 (${((bigMovers / total) * 100).toFixed(1)}%)\n` +
    `- 創30天以上新高: ${newHighCount} 檔 (${((newHighCount / total) * 100).toFixed(1)}%)\n` +
    `- 量能激增 (相對均量>1.5x): ${highVolCount} 檔 (${((highVolCount / total) * 100).toFixed(1)}%)\n` +
    `- RSI 超買區 (>70): ${highRsiCount} 檔 (${((highRsiCount / total) * 100).toFixed(1)}%)`;

  // ── 資料清單 ───────────────────────────────────────────────────────
  const stockListStr = stocks.map(s => {
    const isNewHigh = s.high20 > 0 && s.close >= s.high20;
    return `- ${s.name}(${s.symbol}): 漲幅 ${s.change}%, 題材: ${s.theme}${isNewHigh ? ', (🔥創高)' : ''}`;
  }).join('\n');

  const prompt = `你是全球宏觀與跨市場套利基金的投資長，今天是美東時間 ${today}。
請根據以下美股高動能股票清單（月漲幅超過20%），自動進行【盤前交易策略、板塊分析與台股供應鏈連動】。

${sentimentMetrics}

【原始資料 (美股)】
${stockListStr}

【任務要求】
1. 找出3~5個美股市場主流板塊
2. 每個板塊找出領頭羊與補漲潛力股
3. 重點：針對每個美股板塊，找出【台股連動/供應鏈受惠股票群】（自行補充台股標的）
4. 根據上方客觀數據，綜合判斷全球市場情緒並給出 1~10 分（1=極度恐懼, 5=中性, 10=極度貪婪）
5. 提出盤前操作策略
6. 根據清單內標記「創高」的股票，整理出一份特別觀察名單
7. 重要：請將清單上「所有」美股依其主力題材分配到最適合的板塊（'stocks' 陣列填入代號）。板塊名稱請使用真實市場主流題材（例如：AI晶片設計、邊緣AI與設備）。若有不屬於大主流的股票，請額外建立「其他次要板塊」收容，確保沒有任何股票被遺漏。

【強制輸出結構】
你必須回傳純 JSON 格式，不要包含 \`\`\`json 等 Markdown 標記，直接依據以下 Schema 輸出：
{
  "sentiment_score": 7,
  "sentiment_label": "樂觀投機情緒",
  "market_view": "全球與美股市場定調...（50字以內）",
  "risk_warning": "總經與特定風險提示...（30字以內）",
  "strategies": ["核心策略一", "核心策略二", "核心策略三"],
  "sectors": [
    {
      "sector_name": "AI晶片設計",
      "momentum_score": 9,
      "analysis": "強勢板塊分析...資金流向判斷",
      "stocks": ["NVDA", "AMD", "SMCI"],
      "leader": {"symbol": "NVDA", "name": "Nvidia", "reason": "領頭羊理由"},
      "laggards": [
        {"symbol": "AMD", "name": "AMD", "reason": "補漲潛力股理由"}
      ],
      "taiwan_linked_stocks": [
        {"symbol": "2330", "name": "台積電", "reason": "先進製程獨家代工"},
        {"symbol": "3231", "name": "緯創", "reason": "GPU基板主力供應商"}
      ]
    }
  ],
  "new_high_stocks": ["NVDA Nvidia", "PLTR Palantir"],
  "watchlist": ["NVDA", "ARM", "SMCI"]
}`;

  // 共用 Market_Strategy_台股.js 裡定義的 callGeminiJSON 與 parseJSONSafely 
  // (需確保該檔案已在 GAS 專案內可見)
  const jsonStr = callGeminiJSON(prompt);
  return parseJSONSafely(jsonStr);
}

// ==============================================================================
// 視覺化儀表板渲染模組 (Dashboard Builder)
// ==============================================================================

function getTvUrl(symbol, stocks) {
  const stock = stocks.find(s => String(s.symbol) === String(symbol));
  if (stock && stock.tvUrl) {
    return stock.tvUrl;
  }
  // 美股找不到時 fallback 用 NASDAQ
  return `https://www.tradingview.com/chart/?symbol=NASDAQ:${symbol}`;
}

function buildUSQuantDashboard(ss, strategyJson, stocks) {
  const sheetName = '量化儀表板_美股';
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
    headerBg: '#4A148C', // 深紫 (區分美股)
    subHeaderBg: '#37474F',
    accentUp: '#4CAF50', // 綠漲
    accentWarn: '#FF9800', // 橘警
    accentTw: '#00BCD4', // 青綠色凸顯台股連動
    textMain: '#FFFFFF',
    textSub: '#B0BEC5'
  };

  let currentRow = 2; // 從第 2 列開始畫

  // ── 標題 ──────────────────────────────────────────────────────────
  const titleRange = sheet.getRange(currentRow, 2, 1, 5);
  titleRange.merge()
    .setValue('美股量化盤前策略與台台美連動儀表板')
    .setFontSize(16).setFontWeight('bold')
    .setBackground(colors.headerBg).setFontColor(colors.textMain)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 40);
  currentRow += 2;

  if (strategyJson) {
    // ── 0 情緒計量錶 (儀表板最頂端) ─────────────────────────────────
    const score = typeof strategyJson.sentiment_score === 'number' ? strategyJson.sentiment_score : 5;
    const label = strategyJson.sentiment_label || '中性';

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

    sheet.getRange(currentRow, 2, 2, 2).merge()
      .setValue(score)
      .setFontSize(36).setFontWeight('bold')
      .setBackground(scoreBg).setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 35);
    sheet.setRowHeight(currentRow + 1, 35);

    sheet.getRange(currentRow, 4).setValue('全球市場情緒指數').setFontColor(colors.textSub).setFontSize(10);
    sheet.getRange(currentRow + 1, 4).setValue(label).setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(currentRow, 5, 1, 2).merge().setValue(statusText)
      .setFontColor(scoreBg === '#37474F' ? colors.textMain : scoreBg).setFontWeight('bold');
    sheet.getRange(currentRow + 1, 5, 1, 2).merge().setValue('/ 10 分（1=極度恐懼  10=極度貪婪）')
      .setFontColor(colors.textSub).setFontSize(10);
    currentRow += 3;

    // ── 1 市場定調 & 2 風險提示 ───────────────────────────────────────
    sheet.getRange(currentRow, 2).setValue('全球情緒：').setFontWeight('bold').setFontColor(colors.textSub);
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

    // ── 4 強勢板塊分析、5 領頭羊 & 台股連動 ─────────────────────────
    sheet.getRange(currentRow, 2, 1, 5).merge()
      .setValue('🔥 主流板塊與台股連動追蹤')
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground(colors.subHeaderBg)
      .setFontColor(colors.textMain);
    currentRow++;

    if (strategyJson.sectors) {
      strategyJson.sectors.sort((a, b) => b.momentum_score - a.momentum_score).forEach(sector => {
        // 板塊標題
        sheet.getRange(currentRow, 2, 1, 5).merge().setValue(`【 ${sector.sector_name} 】 動能: ${sector.momentum_score}/10`).setFontColor('#E1BEE7').setFontWeight('bold'); // 淡紫
        currentRow++;

        // 板塊分析
        sheet.getRange(currentRow, 2).setValue('分析').setFontColor(colors.textSub);
        sheet.getRange(currentRow, 3, 1, 4).merge().setValue(sector.analysis).setWrap(true);
        currentRow++;

        // 美股領頭羊
        if (sector.leader) {
          const url = getTvUrl(sector.leader.symbol, stocks);
          const nameLink = `=HYPERLINK("${url}", "${sector.leader.name} (${sector.leader.symbol})")`;
          sheet.getRange(currentRow, 2).setValue('👑 美股領頭羊').setFontColor('#FFD700');
          sheet.getRange(currentRow, 3).setValue(nameLink).setFontWeight('bold');
          sheet.getRange(currentRow, 4, 1, 3).merge().setValue(sector.leader.reason).setWrap(true);
          currentRow++;
        }

        // 美股補漲股
        if (sector.laggards && sector.laggards.length > 0) {
          sector.laggards.forEach((laggard, idx) => {
            const url = getTvUrl(laggard.symbol, stocks);
            const nameLink = `=HYPERLINK("${url}", "${laggard.name} (${laggard.symbol})")`;
            sheet.getRange(currentRow, 2).setValue(idx === 0 ? '🚀 美股補漲' : '').setFontColor(colors.accentUp);
            sheet.getRange(currentRow, 3).setValue(nameLink);
            sheet.getRange(currentRow, 4, 1, 3).merge().setValue(laggard.reason).setWrap(true).setFontColor(colors.textSub);
            currentRow++;
          });
        }

        // 台美連動股 (新加入的模組)
        if (sector.taiwan_linked_stocks && sector.taiwan_linked_stocks.length > 0) {
          sector.taiwan_linked_stocks.forEach((twStock, idx) => {
            // 直接組成台股 TradingView URL，假設多數為上市(TWSE)
            const twUrl = `https://www.tradingview.com/chart/?symbol=TWSE:${twStock.symbol}`;
            const nameLink = `=HYPERLINK("${twUrl}", "${twStock.name} (${twStock.symbol})")`;
            sheet.getRange(currentRow, 2).setValue(idx === 0 ? '🇹🇼 台股連動' : '').setFontColor(colors.accentTw).setFontWeight('bold');
            sheet.getRange(currentRow, 3).setValue(nameLink);
            sheet.getRange(currentRow, 4, 1, 3).merge().setValue(twStock.reason).setWrap(true).setFontColor(colors.textSub);
            currentRow++;
          });
        }

        currentRow++; // 板塊間留白
      });
    }

    // ── 6 創30天以上新高股票 & 7 今日觀察名單 ─────────────────────────────
    sheet.getRange(currentRow, 2, 1, 5).merge().setValue('📈 創30天以上新高股票').setFontWeight('bold').setBackground(colors.subHeaderBg);
    currentRow++;
    const newHighList = strategyJson.new_high_stocks || [];
    if (newHighList.length === 0) {
      sheet.getRange(currentRow, 2, 1, 5).merge().setValue('無').setFontColor('#FF9800');
      currentRow++;
    } else {
      newHighList.forEach((entry, idx) => {
        const parts = String(entry).trim().split(/\s+/);
        const sym = parts[0];    // e.g. "NVDA"
        const label = parts.length > 1 ? parts.slice(1).join(' ') : sym;  // e.g. "Nvidia"
        const url = getTvUrl(sym, stocks); // 先從 stocks 清單找初始 URL
        const nameLink = `=HYPERLINK("${url}","${sym} ${label}")`;
        sheet.getRange(currentRow, 2).setValue(idx === 0 ? '🔥創高' : '').setFontColor('#FF9800').setFontWeight('bold');
        sheet.getRange(currentRow, 3, 1, 4).merge().setValue(nameLink).setFontColor('#FF9800');
        currentRow++;
      });
    }
    currentRow++;

    sheet.getRange(currentRow, 2, 1, 5).merge().setValue('📌 今日美股觀察名單').setFontWeight('bold').setBackground(colors.subHeaderBg);
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

// ==============================================================================
// 板塊分類工作表（美股，基於 AI 強勢板塊分組）
// ==============================================================================

/**
 * 以 strategyJson.sectors 為基礎，把存檔資料內的個股依 AI 識別的板塊分組，
 * 板塊依 momentum_score 排序，板塊內個股依月漲幅排序，渲染至「板塊分類_美股」
 */
function buildUSSectorBreakdown(ss, strategyJson, allStocks) {
  if (!strategyJson || !strategyJson.sectors || strategyJson.sectors.length === 0) {
    Logger.log('⚠️ 無 AI sectors 資料，跳過美股板塊分類渲染');
    return;
  }

  // ── 快速查找表 ────────────────────────────────────────────────────
  const stockMap = {};
  allStocks.forEach(s => { stockMap[String(s.symbol)] = s; });

  // ── 板塊依 momentum_score 降冪排序 ──────────────────────────────
  const sectors = strategyJson.sectors.slice()
    .sort((a, b) => (b.momentum_score || 0) - (a.momentum_score || 0));

  // ── 建立工作表 ───────────────────────────────────────────────────
  const sheetName = '板塊分類_美股';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) { sheet.clear(); sheet.clearFormats(); }
  else { sheet = ss.insertSheet(sheetName); }

  const TOTAL_COLS = 8;
  sheet.getRange(1, 1, 600, TOTAL_COLS)
    .setBackground('#1E1E1E').setFontColor('#E0E0E0').setFontFamily('Arial');
  [70, 150, 220, 70, 70, 65, 80, 55].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  let r = 1;

  // ── 大標題 ──
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue('📊 美股強勢板塊分類（AI 識別板塊 × 月漲幅排序）')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#4A148C').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(r, 34);
  r++;

  // 紀錄已被 AI 分配的股票
  const assignedSymbols = new Set();

  // ── 逐板塊渲染 ──
  sectors.forEach(sector => {
    const sectorSymbols = (sector.stocks || []).map(String);
    const leaderSymbol = String(sector.leader ? sector.leader.symbol : '');

    const sectorStocks = sectorSymbols
      .map(sym => stockMap[sym])
      .filter(Boolean)
      .sort((a, b) => b.perf1m - a.perf1m);

    sectorStocks.forEach(s => assignedSymbols.add(String(s.symbol)));

    const score = sector.momentum_score || '-';
    const count = sectorStocks.length;
    const avgPerf = count > 0
      ? (sectorStocks.reduce((s, x) => s + x.perf1m, 0) / count).toFixed(1)
      : '?';

    // 板塊標題
    sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
      .setValue(`【 ${sector.sector_name} 】  動能 ${score}/10  ｜  共 ${count} 檔  均月漲幅 +${avgPerf}%`)
      .setFontWeight('bold').setFontSize(10)
      .setBackground('#2c1263').setFontColor('#CE93D8');
    sheet.setRowHeight(r, 26);
    r++;

    // 板塊分析文字
    if (sector.analysis && sector.analysis.trim() !== "") {
      sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
        .setValue(`📝 ${sector.analysis}`)
        .setFontSize(9).setFontColor('#B0BEC5').setBackground('#212121').setWrap(true);
      sheet.setRowHeight(r, 22);
      r++;
    }

    // 欄位表頭
    const headers = ['代碼', '股票名稱', 'AI 題材', '月漲幅%', '五日漲%', '日漲幅%', '收盤價(US$)', 'RSI'];
    sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([headers])
      .setFontWeight('bold').setBackground('#263238').setFontColor('#B0BEC5').setFontSize(9);
    sheet.setRowHeight(r, 20);
    r++;

    if (sectorStocks.length === 0) {
      sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
        .setValue('（本板塊個股不在當前存檔資料中）')
        .setFontColor('#546E7A').setFontSize(9);
      r++;
    } else {
      sectorStocks.forEach((s, idx) => {
        const isLeader = String(s.symbol) === leaderSymbol;
        const bg = isLeader ? '#1B1B3A' : (idx % 2 === 0 ? '#1E1E1E' : '#252525');
        sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(bg).setFontSize(10);

        const symbolLabel = isLeader ? `👑 ${s.symbol}` : s.symbol;
        sheet.getRange(r, 1).setValue(symbolLabel).setFontColor(isLeader ? '#FFD700' : '#CFD8DC');

        const nc = sheet.getRange(r, 2);
        if (s.tvUrl) {
          const cleanName = String(s.nameText || s.symbol).replace(/^🔥\s*/, '');
          nc.setValue(`=HYPERLINK("${s.tvUrl}","${cleanName}")`).setFontColor(isLeader ? '#FFD700' : '#CE93D8');
        } else {
          nc.setValue(s.nameText || s.symbol).setFontColor('#CE93D8');
        }

        sheet.getRange(r, 3).setValue(s.theme || '').setFontSize(9).setWrap(true).setFontColor('#E0E0E0');

        const clr = v => v >= 0 ? '#4CAF50' : '#EF5350';
        sheet.getRange(r, 4).setValue(s.perf1m).setFontColor(clr(s.perf1m)).setFontWeight('bold').setHorizontalAlignment('center');
        sheet.getRange(r, 5).setValue(s.perf5d).setFontColor(clr(s.perf5d)).setHorizontalAlignment('center');
        sheet.getRange(r, 6).setValue(s.change).setFontColor(clr(s.change)).setHorizontalAlignment('center');
        sheet.getRange(r, 7).setValue(s.close).setFontColor('#ECEFF1').setHorizontalAlignment('center');
        sheet.getRange(r, 8).setValue(s.rsi).setFontColor('#FFF176').setHorizontalAlignment('center');

        sheet.setRowHeight(r, 24);
        r++;
      });
    }

    r++;
  });

  // ── 兜底：未被分類的美股強勢股 ──
  const unassigned = allStocks.filter(s => !assignedSymbols.has(String(s.symbol))).sort((a, b) => b.perf1m - a.perf1m);
  if (unassigned.length > 0) {
    const avgPerf = (unassigned.reduce((s, x) => s + x.perf1m, 0) / unassigned.length).toFixed(1);

    sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
      .setValue(`【 其他強勢股 (未分類) 】  動能 -/10  ｜  共 ${unassigned.length} 檔  均月漲幅 +${avgPerf}%`)
      .setFontWeight('bold').setFontSize(10)
      .setBackground('#1C0A35').setFontColor('#CFD8DC');
    sheet.setRowHeight(r, 26);
    r++;

    const headers = ['代碼', '股票名稱', 'AI 題材', '月漲幅%', '五日漲%', '日漲幅%', '收盤價(US$)', 'RSI'];
    sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([headers])
      .setFontWeight('bold').setBackground('#263238').setFontColor('#B0BEC5').setFontSize(9);
    sheet.setRowHeight(r, 20);
    r++;

    unassigned.forEach((s, idx) => {
      const bg = idx % 2 === 0 ? '#1E1E1E' : '#252525';
      sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(bg).setFontSize(10);
      sheet.getRange(r, 1).setValue(s.symbol).setFontColor('#CFD8DC');

      const nc = sheet.getRange(r, 2);
      if (s.tvUrl) {
        const cleanName = String(s.nameText || s.symbol).replace(/^🔥\s*/, '');
        nc.setValue(`=HYPERLINK("${s.tvUrl}","${cleanName}")`).setFontColor('#CE93D8');
      } else {
        nc.setValue(s.nameText || s.symbol).setFontColor('#CE93D8');
      }

      sheet.getRange(r, 3).setValue(s.theme || '').setFontSize(9).setWrap(true).setFontColor('#E0E0E0');
      const clr = v => v >= 0 ? '#4CAF50' : '#EF5350';
      sheet.getRange(r, 4).setValue(s.perf1m).setFontColor(clr(s.perf1m)).setFontWeight('bold').setHorizontalAlignment('center');
      sheet.getRange(r, 5).setValue(s.perf5d).setFontColor(clr(s.perf5d)).setHorizontalAlignment('center');
      sheet.getRange(r, 6).setValue(s.change).setFontColor(clr(s.change)).setHorizontalAlignment('center');
      sheet.getRange(r, 7).setValue(s.close).setFontColor('#ECEFF1').setHorizontalAlignment('center');
      sheet.getRange(r, 8).setValue(s.rsi).setFontColor('#FFF176').setHorizontalAlignment('center');
      sheet.setRowHeight(r, 24);
      r++;
    });
  }

  Logger.log(`✅ 板塊分類_美股 渲染完成，共 ${sectors.length} 個 AI 板塊`);
}
