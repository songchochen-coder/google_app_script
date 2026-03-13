/**
 * Market_Strategy_台股.js
 * ========================
 * 從「台股存檔資料」讀取個股，呼叫 AI 產生市場策略報告，
 * 並渲染成專業格式的 Google Sheets 工作表。
 * 共用函式（callGemini, renderMarkdownToSheet）定義於 Global_Config.js。
 */

function runMarketStrategy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('台股存檔資料');

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('找不到「台股存檔資料」工作表，請先執行台股個股分析。');
    return;
  }

  // 初始化報告工作表
  let reportSheet = ss.getSheetByName('AI 分析報告_台股');
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet('AI 分析報告_台股');
  }

  // 提取原始資料
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('「台股存檔資料」中沒有足夠資料。');
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

  // 呼叫 AI 產生報告
  Logger.log('正在產生 AI 台股市場報告...');
  const themesSummary = summarizeTopThemes(analysisResults);
  const strategyReport = findLeadersAndSpillovers(themesSummary);

  const fullContent =
    `執行時間：${new Date().toLocaleString()}\n\n` +
    themesSummary + '\n\n' +
    strategyReport;

  // 渲染到工作表（使用共用 renderMarkdownToSheet）
  renderMarkdownToSheet(reportSheet, fullContent, { c1: 150, c2: 120, c3: 400, c4: 200 }, '#1a237e');

  Logger.log('台股市場分析完畢！');
  SpreadsheetApp.getUi().alert('AI 台股專業格式報告已生成！');
}

/**
 * 提示詞：台股題材族群分析
 */
function summarizeTopThemes(results) {
  const context = results.map(r => `- ${r.symbol} ${r.name}, 漲幅:${r.change}%, 題材:${r.theme}`).join('\n');
  const prompt =
    `你現在是專業台股市場分析師。根據資料撰寫報告：\n${context}\n` +
    `格式要求：\n### 一、 數據篩選\n使用表格列出題材、佔比、名單。\n### 二、 前三大熱門題材\n條列基本面與消息面。`;
  return callGemini(prompt, { temperature: 0.5, maxOutputTokens: 2000 });
}

/**
 * 提示詞：領頭羊與外溢效應分析
 */
function findLeadersAndSpillovers(themesSummary) {
  const prompt =
    `你現在是專業基金經理人。根據以下內容撰寫報告：\n${themesSummary}\n` +
    `格式要求：\n### 三、 資金外溢與價值低估挖掘\n包含領頭羊、低估股、外溢效應、補漲股。`;
  return callGemini(prompt, { temperature: 0.5, maxOutputTokens: 2000 });
}