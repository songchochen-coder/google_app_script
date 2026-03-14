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

  // 4. 存入板塊輪動歷史，並渲染熱力圖
  Logger.log('🗃️ 寫入板塊輪動歷史...');
  saveSectorHistory(ss, strategyJson);
  Logger.log('🔥 正在繪製板塊輪動熱力圖...');
  buildSectorHeatmap(ss);

  try {
    SpreadsheetApp.getUi().alert('🎯 台股專業量化儀表板已生成！\n請查看「量化儀表板_台股」與「板塊輪動熱力圖」工作表。');
  } catch (e) {
    Logger.log('🎯 台股專業量化儀表板已生成！');
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
  "sector_scores": {
    "AI伺服器/HPC": 8,
    "AI晶片/封裝": 7,
    "散熱/機構件": 6,
    "電源/PCB": 5,
    "網通/資安": 4,
    "電動車/儲能": 3,
    "生技/醫療": 2,
    "金融/壽險": 5,
    "傳產/原物料": 3,
    "消費電子/手機": 4
  },
  "sectors": [
    {
      "sector_name": "板塊名稱 (例如: AI伺服器/HPC)",
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

function getTvUrl(symbol, stocks) {
  const stock = stocks.find(s => String(s.symbol) === String(symbol));
  if (stock && stock.tvUrl) {
    return stock.tvUrl;
  }
  return `https://www.tradingview.com/chart/?symbol=TWSE:${symbol}`;
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

// ==============================================================================
// 板塊輪動歷史 & 熱力圖模組
// ==============================================================================

/**
 * 固定板塊清單（台股，10個標準化板塊）
 * AI 會按此清單逐一評分，保持欄位一致以利跨日比較
 */
const TW_SECTOR_KEYS = [
  'AI伺服器/HPC', 'AI晶片/封裝', '散熱/機構件', '電源/PCB',
  '網通/資安', '電動車/儲能', '生技/醫療', '金融/壽險',
  '傳產/原物料', '消費電子/手機'
];

/**
 * 把當日板塊分數 & 情緒分數追加到歷史工作表
 * 格式：第1欄=日期, 後續欄=各板塊分數, 最後欄=情緒指數
 */
function saveSectorHistory(ss, strategyJson) {
  if (!strategyJson) return;

  const histSheetName = '板塊輪動歷史_台股';
  let histSheet = ss.getSheetByName(histSheetName);

  // 如果歷史工作表不存在，建立並加入表頭
  if (!histSheet) {
    histSheet = ss.insertSheet(histSheetName);
    const headerRow = ['日期', ...TW_SECTOR_KEYS, '情緒指數'];
    histSheet.appendRow(headerRow);
    histSheet.getRange(1, 1, 1, headerRow.length)
      .setFontWeight('bold')
      .setBackground('#263238')
      .setFontColor('#ECEFF1');
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  const sectorScores = strategyJson.sector_scores || {};

  // 組成一行資料，確保板塊順序固定
  const row = [today];
  TW_SECTOR_KEYS.forEach(key => {
    const score = sectorScores[key];
    row.push(typeof score === 'number' ? score : null);
  });
  row.push(strategyJson.sentiment_score || null);

  histSheet.appendRow(row);
  Logger.log(`✅ 已將 ${today} 板塊分數寫入歷史紀錄`);
}

/**
 * 讀取歷史工作表，渲染「板塊輪動熱力圖」
 * - 橫軸: 最近 N 天（最多顯示 20 天）
 * - 縱軸: 10個固定板塊
 * - 格子: 1~10 分動態上色
 */
function buildSectorHeatmap(ss) {
  const histSheetName = '板塊輪動歷史_台股';
  const histSheet = ss.getSheetByName(histSheetName);
  if (!histSheet) {
    Logger.log('⚠️ 找不到歷史工作表，無法渲染熱力圖');
    return;
  }

  // 讀取歷史資料
  const histData = histSheet.getDataRange().getValues();
  if (histData.length <= 1) {
    Logger.log('⚠️ 歷史資料不足，至少需要1行資料');
    return;
  }

  // 建立/清空熱力圖工作表
  const hmSheetName = '板塊輪動熱力圖';
  let hmSheet = ss.getSheetByName(hmSheetName);
  if (hmSheet) {
    hmSheet.clear();
    hmSheet.clearFormats();
  } else {
    hmSheet = ss.insertSheet(hmSheetName);
  }

  // 取最近 20 天數據（排除表頭）
  const START_COL_OFFSET = 20;  // 欄寬 px
  const maxDays = 20;
  const rows = histData.slice(1); // 去掉表頭
  const recentRows = rows.slice(-maxDays); // 最多取最後20天

  const dates = recentRows.map(r => r[0]);
  const numDates = dates.length;
  const numSectors = TW_SECTOR_KEYS.length;

  // 全域深色背景
  hmSheet.getRange(1, 1, numSectors + 5, numDates + 3).setBackground('#1E1E1E').setFontColor('#E0E0E0').setFontFamily('Arial');

  // ── 大標題 ──
  hmSheet.getRange(1, 2, 1, numDates + 1).merge()
    .setValue('📊 台股板塊輪動熱力圖（近期動能追蹤）')
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#0A3D62').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  hmSheet.setRowHeight(1, 35);

  // ── 凡例 ──
  const legendRow = 2;
  const legendData = [
    ['9-10', '#B71C1C'], ['7-8', '#E64A19'], ['5-6', '#F57F17'],
    ['3-4', '#546E7A'], ['1-2', '#0D47A1'], ['無資料', '#2A2A2A']
  ];
  hmSheet.getRange(legendRow, 2).setValue('圖例:').setFontColor('#B0BEC5').setFontSize(9);
  legendData.forEach((item, idx) => {
    const cell = hmSheet.getRange(legendRow, 3 + idx);
    cell.setValue(item[0]).setBackground(item[1]).setFontColor('#FFFFFF')
      .setFontSize(8).setHorizontalAlignment('center');
  });
  hmSheet.setRowHeight(legendRow, 20);

  // ── 日期表頭（橫軸）──
  const dateHeaderRow = 3;
  hmSheet.getRange(dateHeaderRow, 2).setValue('板塊 \\ 日期').setFontWeight('bold').setFontSize(10).setBackground('#263238');
  dates.forEach((d, colIdx) => {
    const cell = hmSheet.getRange(dateHeaderRow, 3 + colIdx);
    const displayDate = typeof d === 'string' ? d.substring(5) : Utilities.formatDate(d, 'Asia/Taipei', 'MM/dd'); // 只顯示 MM/dd
    cell.setValue(displayDate).setFontWeight('bold').setBackground('#263238').setHorizontalAlignment('center').setFontSize(9);
    hmSheet.setColumnWidth(3 + colIdx, 50);
  });
  hmSheet.setRowHeight(dateHeaderRow, 28);
  hmSheet.setColumnWidth(2, 110); // 板塊名稱欄較寬

  // ── 板塊資料格子（縱軸）──
  TW_SECTOR_KEYS.forEach((sectorKey, rowIdx) => {
    const dataRow = dateHeaderRow + 1 + rowIdx;
    hmSheet.setRowHeight(dataRow, 30);

    // 板塊名稱欄
    hmSheet.getRange(dataRow, 2).setValue(sectorKey)
      .setFontSize(9).setFontColor('#E0E0E0').setBackground('#263238');

    // 各日期分數格
    recentRows.forEach((histRow, colIdx) => {
      const cellScore = histRow[1 + rowIdx]; // +1 跳過日期欄
      const cell = hmSheet.getRange(dataRow, 3 + colIdx);

      if (cellScore === null || cellScore === '' || typeof cellScore !== 'number') {
        cell.setValue('').setBackground('#2A2A2A');
      } else {
        const score = Math.max(1, Math.min(10, Math.round(cellScore)));
        cell.setValue(score).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(11);

        // 動態上色
        let bg, fg;
        if (score >= 9) { bg = '#B71C1C'; fg = '#FFFFFF'; }
        else if (score >= 7) { bg = '#E64A19'; fg = '#FFFFFF'; }
        else if (score >= 5) { bg = '#F57F17'; fg = '#212121'; }
        else if (score >= 3) { bg = '#546E7A'; fg = '#FFFFFF'; }
        else { bg = '#0D47A1'; fg = '#FFFFFF'; }
        cell.setBackground(bg).setFontColor(fg);
      }
    });
  });

  // ── 情緒指數尾行 ──
  const sentimentRow = dateHeaderRow + numSectors + 1;
  hmSheet.setRowHeight(sentimentRow, 28);
  hmSheet.getRange(sentimentRow, 2).setValue('🌡️ 情緒指數').setFontColor('#FFD700').setFontWeight('bold').setBackground('#263238').setFontSize(9);
  recentRows.forEach((histRow, colIdx) => {
    const sentScore = histRow[1 + numSectors]; // 最後一欄是情緒分數
    const cell = hmSheet.getRange(sentimentRow, 3 + colIdx);
    if (sentScore === null || sentScore === '') {
      cell.setValue('').setBackground('#2A2A2A');
    } else {
      const s = Math.max(1, Math.min(10, Math.round(sentScore)));
      let bg = s >= 8 ? '#B71C1C' : s >= 6 ? '#E65100' : s >= 4 ? '#546E7A' : '#0D47A1';
      cell.setValue(s).setBackground(bg).setFontColor('#FFFFFF')
        .setHorizontalAlignment('center').setFontWeight('bold').setFontSize(11);
    }
  });

  Logger.log(`✅ 板塊輪動熱力圖已渲染，共 ${numDates} 天 × ${numSectors} 板塊`);
}