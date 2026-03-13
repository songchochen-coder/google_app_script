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
  const data = sourceSheet.getDataRange().getDisplayValues(); // 使用 DisplayValues 確保拿到文字
  if (data.length <= 1) {
    Logger.log('⚠️ 「台股存檔資料」中沒有足夠資料。');
    try { SpreadsheetApp.getUi().alert('「台股存檔資料」中沒有足夠資料。'); } catch (e) { }
    return;
  }

  const stocks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      // 假設遇到有 HYPERLINK 公式的儲存格，使用 Text 取值（或簡單取字串）
      // 因為 getDisplayValues() 會拿到純文字，所以直接可用
      stocks.push({
        symbol: data[i][0],
        name: data[i][1],
        change: parseFloat(data[i][2]) || 0,
        theme: data[i][3]
      });
    }
  }

  // 2. 進行 AI 深度分析 (要求回傳 JSON)
  Logger.log('🚀 正在進行 AI 題材聚類與板塊分析...');
  const clusteringJson = analyzeThematicClustering(stocks);

  Logger.log('🚀 正在挖掘領頭羊與補漲股...');
  const leadershipJson = detectLeadersAndLaggards(clusteringJson, stocks);

  Logger.log('🚀 正在生成總體市場策略...');
  const strategyJson = generateMarketStrategy(clusteringJson, leadershipJson);

  // 3. 渲染 Google Sheets 儀表板
  Logger.log('🎨 正在繪製量化儀表板...');
  buildQuantDashboard(ss, clusteringJson, leadershipJson, strategyJson);

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
function analyzeThematicClustering(stocks) {
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const stockList = stocks.map(s => `- ${s.name}(${s.symbol}): 漲幅 ${s.change}%, 題材: ${s.theme}`).join('\n');

  const prompt = `你是頂尖的量化基金經理人，今天是 ${today}。
請根據以下高動能股票清單（月漲幅超過20%），自動進行【題材聚類與板塊資金流向分析】。

【原始資料】
${stockList}

【任務要求】
1. 將這些股票依據「真實產業/題材」聚類為 3~5 大強勢板塊。
2. 評估該板塊的資金動能強度（1-10分）。
3. 給出對該板塊輪動趨勢的簡短看法。

【強制輸出】
你必須回傳純 JSON 格式，不要包含 \`\`\`json 等 Markdown 標記，直接輸出：
{
  "date": "202X-XX-XX",
  "sectors": [
    {
      "sector_name": "AI伺服器供應鏈",
      "momentum_score": 9,
      "trend_analysis": "資金持續湧入散熱與滑軌，有擴散至電源跡象。",
      "key_stocks": ["2330", "3017", "3324"]
    }
  ]
}`;

  const jsonStr = callGeminiJSON(prompt);
  return parseJSONSafely(jsonStr);
}

/**
 * Task 2: 領頭羊與補漲股偵測
 */
function detectLeadersAndLaggards(clusteringJson, stocks) {
  if (!clusteringJson || !clusteringJson.sectors) return {};

  const prompt = `你是專精於「板塊輪動與籌碼擴散」的台股操盤手。
以下是目前的強勢板塊聚類結果：
${JSON.stringify(clusteringJson)}

請找出每個板塊中的「絕對領頭羊（Leader）」以及「具備潛力的補漲股或外溢受惠股（Laggard/Spillover）」。
補漲股可以是清單內的落後者，也可以是你憑專業知識找出「尚未出現在清單中，但同屬該產業且基期較低」的台股標的。

【強制輸出】
必須回傳純 JSON 格式，不可有 Markdown 標記：
{
  "leadership": [
    {
      "sector_name": "AI伺服器供應鏈",
      "leader": {"symbol": "2330", "name": "台積電", "reason": "先進封裝技術護城河"},
      "laggards": [
        {"symbol": "XXXX", "name": "XXX", "reason": "同屬散熱板塊，本益比偏低"}
      ]
    }
  ]
}`;

  const jsonStr = callGeminiJSON(prompt);
  return parseJSONSafely(jsonStr);
}

/**
 * Task 3: 總體市場策略
 */
function generateMarketStrategy(clusteringJson, leadershipJson) {
  if (!clusteringJson || !leadershipJson) return {};

  const prompt = `你是管理百億資金的台股投資長。
目前的盤面板塊輪動如下：
${JSON.stringify(clusteringJson)}
目前的領頭羊與補漲輪廓如下：
${JSON.stringify(leadershipJson)}

請結合當前總體經濟環境與上述數據，給出今日/本週的具體交易策略。

【強制輸出】
必須回傳純 JSON 格式，不可有 Markdown 標記：
{
  "market_view": "強勢股續強，資金呈現...（50字以內）",
  "action_plan": ["策略一...", "策略二...", "策略三..."],
  "risk_warning": "當前需注意的風險...（30字以內）"
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

function buildQuantDashboard(ss, clusteringJson, leadershipJson, strategyJson) {
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

  // ── 區塊 A：主標題與市場總覽 ────────────────────────────────────
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

  if (strategyJson && strategyJson.market_view) {
    sheet.getRange(currentRow, 2).setValue('大盤定調：').setFontWeight('bold').setFontColor(colors.textSub);
    sheet.getRange(currentRow, 3, 1, 4).merge().setValue(strategyJson.market_view).setFontColor(colors.textMain);
    currentRow++;

    sheet.getRange(currentRow, 2).setValue('風險提示：').setFontWeight('bold').setFontColor(colors.accentWarn);
    sheet.getRange(currentRow, 3, 1, 4).merge().setValue(strategyJson.risk_warning).setFontColor(colors.accentWarn);
    currentRow += 2;

    // 核心操作策略 (三個 bullet points)
    sheet.getRange(currentRow, 2, 1, 5).merge().setValue('核心操作計畫').setFontWeight('bold').setBackground(colors.subHeaderBg);
    currentRow++;
    const plans = strategyJson.action_plan || [];
    plans.forEach((plan, idx) => {
      sheet.getRange(currentRow, 2).setValue(`策略 ${idx + 1}`);
      sheet.getRange(currentRow, 3, 1, 4).merge().setValue(plan).setWrap(true);
      currentRow++;
    });
    currentRow++;
  }

  // ── 區塊 B：強勢板塊與資金輪動 ──────────────────────────────────
  sheet.getRange(currentRow, 2, 1, 5).merge()
    .setValue('🔥 強勢板塊資金輪動聚類')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground(colors.subHeaderBg)
    .setFontColor(colors.textMain);
  currentRow++;

  // 標題列
  sheet.getRange(currentRow, 2).setValue('板塊名稱').setFontWeight('bold').setBorder(null, null, true, null, null, null, '#444444', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(currentRow, 3).setValue('動能分數').setFontWeight('bold').setBorder(null, null, true, null, null, null, '#444444', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(currentRow, 4, 1, 3).merge().setValue('資金流向研判').setFontWeight('bold').setBorder(null, null, true, null, null, null, '#444444', SpreadsheetApp.BorderStyle.SOLID);
  currentRow++;

  if (clusteringJson && clusteringJson.sectors) {
    clusteringJson.sectors.sort((a, b) => b.momentum_score - a.momentum_score).forEach(sector => {
      sheet.getRange(currentRow, 2).setValue(sector.sector_name);

      const scoreCell = sheet.getRange(currentRow, 3);
      scoreCell.setValue(sector.momentum_score + ' / 10').setHorizontalAlignment('center');
      if (sector.momentum_score >= 8) scoreCell.setFontColor(colors.accentUp).setFontWeight('bold');

      sheet.getRange(currentRow, 4, 1, 3).merge().setValue(sector.trend_analysis).setWrap(true).setFontColor(colors.textSub);
      currentRow++;
    });
  }
  currentRow++;

  // ── 區塊 C：領頭羊與補漲股挖掘清單 ─────────────────────────────
  sheet.getRange(currentRow, 2, 1, 5).merge()
    .setValue('🎯 領頭羊與補漲潛力雷達')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground(colors.subHeaderBg)
    .setFontColor(colors.textMain);
  currentRow++;

  if (leadershipJson && leadershipJson.leadership) {
    leadershipJson.leadership.forEach(group => {
      // 繪製板塊名稱橫幅
      sheet.getRange(currentRow, 2, 1, 5).merge().setValue(`【 ${group.sector_name} 】`).setFontColor('#81D4FA').setFontWeight('bold');
      currentRow++;

      // 領頭羊
      if (group.leader) {
        sheet.getRange(currentRow, 2).setValue('👑 領頭羊').setFontColor('#FFD700');
        sheet.getRange(currentRow, 3).setValue(`${group.leader.name} (${group.leader.symbol})`).setFontWeight('bold');
        sheet.getRange(currentRow, 4, 1, 3).merge().setValue(group.leader.reason).setWrap(true);
        currentRow++;
      }

      // 補漲股 (可能有多支)
      if (group.laggards && group.laggards.length > 0) {
        group.laggards.forEach((laggard, idx) => {
          sheet.getRange(currentRow, 2).setValue(idx === 0 ? '🚀 補漲/外溢' : '').setFontColor(colors.accentUp);
          sheet.getRange(currentRow, 3).setValue(`${laggard.name} (${laggard.symbol})`);
          sheet.getRange(currentRow, 4, 1, 3).merge().setValue(laggard.reason).setWrap(true).setFontColor(colors.textSub);
          currentRow++;
        });
      }
      currentRow++; // 板塊間留空行
    });
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