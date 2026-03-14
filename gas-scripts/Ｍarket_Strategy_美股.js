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

  try {
    SpreadsheetApp.getUi().alert('🎯 美股專業量化儀表板（含台股連動）已生成！\n請查看「量化儀表板_美股」工作表。');
  } catch (e) {
    Logger.log('🎯 美股專業量化儀表板已生成！請查看「量化儀表板_美股」工作表。');
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

  // 計算是否有創高 (這裡將 close >= high20 視為創 20/30 日以上新高)
  const stockListStr = stocks.map(s => {
    const isNewHigh = s.high20 > 0 && s.close >= s.high20;
    return `- ${s.name}(${s.symbol}): 漲幅 ${s.change}%, 題材: ${s.theme}, 最新價: ${s.close}, 波段高點: ${s.high20} ${isNewHigh ? '(🔥創30天以上新高)' : ''}`;
  }).join('\n');

  const prompt = `你是全球宏觀與跨市場套利基金的投資長，今天是美東時間 ${today}。
請根據以下美股高動能股票清單（月漲幅超過20%），自動進行【盤前交易策略、板塊分析與台股供應鏈連動】。

【原始資料 (美股)】
${stockListStr}

【任務要求】
1. 找出3~5個美股市場主流板塊
2. 每個板塊找出領頭羊與補漲潛力股
3. 重點：針對每個美股板塊，找出【台股連動/供應鏈受惠股票群】（自行補充台股標的）
4. 判斷全球市場情緒
5. 提出盤前操作策略
6. 根據清單內標記「創30天以上新高」的股票，整理出一份特別觀察名單

【強制輸出結構】
你必須回傳純 JSON 格式，不要包含 \`\`\`json 等 Markdown 標記，直接依據以下 Schema 輸出：
{
  "market_view": "全球與美股市場定調...（50字以內）",
  "risk_warning": "總經與特定風險提示...（30字以內）",
  "strategies": ["核心策略一", "核心策略二", "核心策略三"],
  "sectors": [
    {
      "sector_name": "板塊名稱 (例如: AI晶片設計)",
      "momentum_score": 9,
      "analysis": "強勢板塊分析...資金流向判斷",
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
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground(colors.headerBg)
    .setFontColor(colors.textMain)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 40);
  currentRow += 2;

  if (strategyJson) {
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
