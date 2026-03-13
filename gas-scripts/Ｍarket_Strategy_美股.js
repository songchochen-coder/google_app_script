/**
 * 腳本：美股板塊趨勢與台股聯動分析 (視覺化渲染版)
 * 功能：從存檔資料提取美股數據，並將分析結果渲染為具備專業格式的試算表。
 */

function runUSMarketStrategy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('美股存檔資料');
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('找不到「美股存檔資料」工作表，請確認資料已存檔。');
    return;
  }

  // 1. 初始化報告工作表
  let reportSheet = ss.getSheetByName('AI 分析報告_美股與台股聯動');
  if (reportSheet) {
    reportSheet.clear(); 
  } else {
    reportSheet = ss.insertSheet('AI 分析報告_美股與台股聯動');
  }

  // 2. 提取原始資料
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('「美股存檔資料」中沒有足夠資料。');
    return;
  }

  // 整理美股數據
  const analysisResults = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== "") {
      analysisResults.push({
        symbol: data[i][0],
        name: data[i][1],
        change: data[i][2],
        theme: data[i][3]
      });
    }
  }

  // 3. 呼叫 AI 產生美股與台股聯動分析
  Logger.log('正在產生美股板塊與台股聯動報告...');
  const themesSummary = summarizeUSTopThemes(analysisResults);
  const finalStrategy = findUSLeadersAndTaiwanSupplyChain(themesSummary);
  
  // 合併內容準備渲染
  const fullContent = `執行時間：${new Date().toLocaleString()}\n\n` + 
                      themesSummary + "\n\n" + 
                      finalStrategy;

  // 4. 執行 Markdown 轉 試算表格式渲染
  renderMarkdownToSheet(reportSheet, fullContent);

  Logger.log('市場聯動分析執行完畢！');
  SpreadsheetApp.getUi().alert('美股-台股聯動專業報告已生成！');
}

/**
 * 渲染器：將 Markdown 文本轉換為 Google Sheets 樣式
 */
function renderMarkdownToSheet(sheet, text) {
  const lines = text.split('\n');
  let currentRow = 1;

  // 設定欄寬 (針對美股-台股報告優化)
  sheet.setColumnWidth(1, 180); // 板塊 / 類別
  sheet.setColumnWidth(2, 120); // 佔比 / 代碼
  sheet.setColumnWidth(3, 450); // 名單 / 台股供應鏈
  sheet.setColumnWidth(4, 150); // 備註

  lines.forEach(line => {
    line = line.trim();
    if (line === "") {
      currentRow++;
      return;
    }

    // --- 處理標題 (###) ---
    if (line.startsWith('###')) {
      const title = line.replace(/###/g, '').trim();
      sheet.getRange(currentRow, 1, 1, 4).merge()
           .setValue(title)
           .setBackground('#0d47a1') // 深藍色
           .setFontColor('#ffffff')   
           .setFontWeight('bold')
           .setFontSize(12)
           .setVerticalAlignment('middle');
      sheet.setRowHeight(currentRow, 30);
    } 
    
    // --- 處理表格 (| ... |) ---
    else if (line.startsWith('|')) {
      if (line.includes('---')) return; 
      
      const cells = line.split('|').filter(c => c.trim() !== "").map(c => c.trim());
      if (cells.length > 0) {
        sheet.getRange(currentRow, 1, 1, cells.length).setValues([cells])
             .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
        
        // 如果是表格標頭（偵測前一行是否為標題）
        if (currentRow > 1 && sheet.getRange(currentRow - 1, 1).getBackground() === '#0d47a1') {
          sheet.getRange(currentRow, 1, 1, cells.length).setBackground('#e3f2fd').setFontWeight('bold');
        }
      }
    } 
    
    // --- 處理清單或重點分析 ---
    else if (line.startsWith('*') || /^\d+\./.test(line) || line.includes(':')) {
      sheet.getRange(currentRow, 1, 1, 4).mergeAcross()
           .setValue(line)
           .setWrap(true);
    } 
    
    // --- 一般文字 ---
    else {
      sheet.getRange(currentRow, 1, 1, 4).mergeAcross()
           .setValue(line)
           .setWrap(true);
    }

    currentRow++;
  });

  // 全域樣式微調
  sheet.getRange(1, 1, currentRow, 4).setVerticalAlignment('top');
}

/**
 * 提示詞 1：美股板塊歸納
 */
function summarizeUSTopThemes(results) {
  const context = results.map(r => `- ${r.symbol} ${r.name}, 漲幅:${r.change}%, 題材:${r.theme}`).join('\n');
  const prompt = `你現在是資深美股研究員。請根據資料撰寫報告，名稱需轉為中文。
    [資料清單]
    ${context}
    [要求格式]
    ### 一、 美股強勢板塊佔比統計
    使用表格包含：板塊名稱 | 佔比(檔數) | 成員名單(代碼+名稱)
    ### 二、 板塊上漲核心理由分析
    條列式簡述基本面與消息面。`;
  return callGemini(prompt);
}

/**
 * 提示詞 2：對接台股供應鏈
 */
function findUSLeadersAndTaiwanSupplyChain(themesSummary) {
  const prompt = `你現在是專業基金經理人。根據以下美股分析，找出聯動的台股：
    ${themesSummary}
    [要求格式]
    ### 三、 台股供應鏈聯動策略
    請條列：
    1. 該板塊美股領頭羊
    2. 【對應受惠台股】：列出台股名稱與代號（例如：台積電 2330）。
    3. 聯動邏輯簡述。`;
  return callGemini(prompt);
}

/**
 * Gemini API 呼叫 (請確保 API Key 已填寫)
 */
