/**
 * 腳本：台股市場策略分析 (AI 2026 專業視覺化版)
 * 功能：從存檔資料提取個股，並將 AI 的 Markdown 轉化為試算表格式（表格、標題、顏色）。
 */

function runMarketStrategy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('台股存檔資料');
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('找不到「台股存檔資料」工作表，請確認資料已存檔。');
    return;
  }

  // 1. 初始化或建立報告工作表
  let reportSheet = ss.getSheetByName('AI 分析報告_台股');
  if (reportSheet) {
    reportSheet.clear(); // 清除舊資料
  } else {
    reportSheet = ss.insertSheet('AI 分析報告_台股');
  }

  // 2. 提取原始資料
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('「台股存檔資料」中沒有足夠資料。');
    return;
  }

  const analysisResults = [];
  for (let i = 1; i < data.length; i++) {
    analysisResults.push({
      symbol: data[i][0],
      name: data[i][1],
      change: data[i][2],
      theme: data[i][3]
    });
  }

  // 3. 呼叫 AI 產生分析內容
  Logger.log('正在產生 AI 分析報告...');
  const themesSummary = summarizeTopThemes(analysisResults);
  const strategyReport = findLeadersAndSpillovers(themesSummary);
  
  // 合併內容準備渲染
  const fullContent = `執行時間：${new Date().toLocaleString()}\n\n` + themesSummary + "\n\n" + strategyReport;

  // 4. 執行 Markdown 轉 試算表格式渲染
  renderMarkdownToSheet(reportSheet, fullContent);

  Logger.log('市場分析執行完畢！');
  SpreadsheetApp.getUi().alert('AI 專業格式報告已生成！');
}

/**
 * 渲染器：將 Markdown 文本轉換為 Google Sheets 的儲存格樣式
 */
function renderMarkdownToSheet(sheet, text) {
  const lines = text.split('\n');
  let currentRow = 1;

  // 初始化欄寬
  sheet.setColumnWidth(1, 150); // 題材分類 / 項目
  sheet.setColumnWidth(2, 120); // 佔比 / 代碼
  sheet.setColumnWidth(3, 400); // 名單 / 內容
  sheet.setColumnWidth(4, 200); // 備註

  lines.forEach(line => {
    line = line.trim();
    if (line === "") {
      currentRow++;
      return;
    }

    const range = sheet.getRange(currentRow, 1);

    // --- 處理標題 (###) ---
    if (line.startsWith('###')) {
      const title = line.replace(/###/g, '').trim();
      sheet.getRange(currentRow, 1, 1, 4).merge()
           .setValue(title)
           .setBackground('#1a237e') // 深藍色背景
           .setFontColor('#ffffff')   // 白色字體
           .setFontWeight('bold')
           .setFontSize(12)
           .setVerticalAlignment('middle');
      sheet.setRowHeight(currentRow, 30);
    } 
    
    // --- 處理表格 (| ... |) ---
    else if (line.startsWith('|')) {
      if (line.includes('---')) return; // 忽略 Markdown 分隔線行
      
      const cells = line.split('|').filter(c => c.trim() !== "").map(c => c.trim());
      if (cells.length > 0) {
        sheet.getRange(currentRow, 1, 1, cells.length).setValues([cells])
             .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
        
        // 如果是表格標頭（偵測前一行是否為標題）
        if (sheet.getRange(currentRow - 1, 1).getBackground() === '#1a237e') {
          sheet.getRange(currentRow, 1, 1, cells.length).setBackground('#eeeeee').setFontWeight('bold');
        }
      }
    } 
    
    // --- 處理清單 (* 或 1.) ---
    else if (line.startsWith('*') || /^\d+\./.test(line)) {
      sheet.getRange(currentRow, 1, 1, 4).mergeAcross()
           .setValue(line)
           .setWrap(true);
    } 
    
    // --- 處理一般文字 ---
    else {
      sheet.getRange(currentRow, 1, 1, 4).mergeAcross()
           .setValue(line)
           .setWrap(true);
    }

    currentRow++;
  });

  // 全域樣式微調
  sheet.getRange(1, 1, currentRow, 4).setVerticalAlignment('middle');
}

/**
 * 提示詞：族群分析
 */
function summarizeTopThemes(results) {
  const context = results.map(r => `- ${r.symbol} ${r.name}, 漲幅:${r.change}%, 題材:${r.theme}`).join('\n');
  const prompt = `你現在是專業台股市場分析師。根據資料撰寫報告：\n${context}\n格式要求：### 一、 數據篩選\n使用表格列出題材、佔比、名單。### 二、 前三大熱門題材\n條列基本面與消息面。`;
  return callGemini(prompt);
}

/**
 * 提示詞：策略分析
 */
function findLeadersAndSpillovers(themesSummary) {
  const prompt = `你現在是專業經理人。根據以下內容撰寫報告：\n${themesSummary}\n格式要求：### 三、 資金外溢與價值低估挖掘\n包含領頭羊、低估股、外溢效應、補漲股。`;
  return callGemini(prompt);
}

/**
 * Gemini API 呼叫
 */
function callGemini(prompt) {
  const apiKey = "你的_API_KEY_請填於此"; 
  const apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;

  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "temperature": 0.5, "maxOutputTokens": 2000 }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());
    return json.candidates[0].content.parts[0].text;
  } catch (e) {
    return "連線失敗：" + e.toString();
  }
}