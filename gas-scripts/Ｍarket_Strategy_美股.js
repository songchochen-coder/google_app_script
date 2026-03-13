/**
 * Market_Strategy_美股.js
 * ========================
 * 從「美股存檔資料」讀取數據，呼叫 AI 產生美股板塊與台股聯動報告，
 * 並渲染成專業格式的 Google Sheets 工作表。
 * 共用函式（callGemini, renderMarkdownToSheet）定義於 Global_Config.js。
 */

function runUSMarketStrategy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('美股存檔資料');

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('找不到「美股存檔資料」工作表，請先執行美股個股分析。');
    return;
  }

  // 初始化報告工作表
  let reportSheet = ss.getSheetByName('AI 分析報告_美股與台股聯動');
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet('AI 分析報告_美股與台股聯動');
  }

  // 提取原始資料
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('「美股存檔資料」中沒有足夠資料。');
    return;
  }

  const analysisResults = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '') {
      analysisResults.push({
        symbol: data[i][0],
        name: data[i][1],
        change: data[i][2],
        theme: data[i][3]
      });
    }
  }

  // 呼叫 AI 產生美股板塊與台股聯動報告
  Logger.log('正在產生美股板塊與台股聯動報告...');
  const themesSummary = summarizeUSTopThemes(analysisResults);
  const finalStrategy = findUSLeadersAndTaiwanSupplyChain(themesSummary);

  const fullContent =
    `執行時間：${new Date().toLocaleString()}\n\n` +
    themesSummary + '\n\n' +
    finalStrategy;

  // 渲染到工作表（使用共用 renderMarkdownToSheet）
  renderMarkdownToSheet(reportSheet, fullContent, { c1: 180, c2: 120, c3: 450, c4: 150 }, '#0d47a1');

  Logger.log('美股-台股聯動分析完畢！');
  SpreadsheetApp.getUi().alert('美股-台股聯動專業報告已生成！');
}

/**
 * 提示詞 1：美股板塊歸納
 */
function summarizeUSTopThemes(results) {
  const context = results.map(r => `- ${r.symbol} ${r.name}, 漲幅:${r.change}%, 題材:${r.theme}`).join('\n');
  const prompt =
    `你現在是資深美股研究員。請根據資料撰寫報告，名稱需轉為中文。\n` +
    `[資料清單]\n${context}\n` +
    `[要求格式]\n### 一、 美股強勢板塊佔比統計\n使用表格包含：板塊名稱 | 佔比(檔數) | 成員名單(代碼+名稱)\n### 二、 板塊上漲核心理由分析\n條列式簡述基本面與消息面。`;
  return callGemini(prompt, { temperature: 0.5, maxOutputTokens: 2000 });
}

/**
 * 提示詞 2：對接台股供應鏈
 */
function findUSLeadersAndTaiwanSupplyChain(themesSummary) {
  const prompt =
    `你現在是專業基金經理人。根據以下美股分析，找出聯動的台股：\n${themesSummary}\n` +
    `[要求格式]\n### 三、 台股供應鏈聯動策略\n請條列：\n` +
    `1. 該板塊美股領頭羊\n` +
    `2. 【對應受惠台股】：列出台股名稱與代號（例如：台積電 2330）。\n` +
    `3. 聯動邏輯簡述。`;
  return callGemini(prompt, { temperature: 0.5, maxOutputTokens: 2000 });
}
