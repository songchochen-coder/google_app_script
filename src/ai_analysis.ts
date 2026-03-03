import { Stock } from './tradingview';

export interface Stage1Result {
    code: string;
    name: string;
    theme: string;
    reason: string;
}

export async function searchStockNews(ticker: string, name: string, apiKey: string): Promise<string> {
    if (!apiKey) return "No Search API Key provided. AI will rely on internal knowledge.";

    const query = `${ticker} ${name} 台股 上漲 題材 新聞 2024 2025`;
    try {
        const response = await fetch("https://api.tavily.com/search", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                api_key: apiKey,
                query: query,
                search_depth: "advanced",
                include_domains: ["money.udn.com", "chinatimes.com", "cnee.com.tw", "ctee.com.tw", "technews.tw"]
            })
        });
        const data = await response.json() as any;
        return data.results.map((r: any) => r.content).join("\n\n").slice(0, 3000);
    } catch (e) {
        return "Search failed. AI will rely on internal knowledge.";
    }
}

export async function runStage1(stocks: Stock[], ai: any, searchApiKey: string): Promise<Stage1Result[]> {
    const stockListStr = stocks.map(s => `${s.ticker} ${s.name}`).join(", ");

    // To optimize for speed, we'll do one big prompt for Stage 1 if the list is short,
    // or batch them if long. Since we limited to 15, one prompt might be too big for context.
    // But let's try a single comprehensive prompt for "The list" first as requested by user.

    const prompt = `你現在是一位專業的台股分析師。請分析以下這份月漲幅超過 20% 的股票名單：\n[${stockListStr}]\n\n任務：\n請針對每支股票，總結其近一個月上漲的核心題材（如：法說會利多、特定產品獲認證、產業趨勢如矽光子/CPO、政策標案、集團作帳等）。\n請排除純技術面（如：跌深反彈）的描述，專注於「產業與基本面消息」。\n\n輸出格式（JSON）：\n[\n{"code": "股票代號", "name": "名稱", "theme": "核心題材關鍵字", "reason": "簡短原理解析"}\n]`;

    const response = await ai.run('@cf/meta/llama-3.1-8b-instruct', {
        messages: [
            { role: "system", content: "你是一個專業的台灣股市分析助手。請僅以 JSON 格式回應。" },
            { role: "user", content: prompt }
        ]
    });

    try {
        // Basic JSON extraction from potential markdown
        const jsonStr = response.response.match(/\[.*\]/s)?.[0] || response.response;
        return JSON.parse(jsonStr);
    } catch (e) {
        console.error("Stage 1 Parse Error:", e);
        return [];
    }
}

export async function runStage2(stage1Results: Stage1Result[], ai: any): Promise<string> {
    const inputData = JSON.stringify(stage1Results, null, 2);
    const prompt = `根據以下分析後的個股題材資料，請進行大數據彙整：\n${inputData}\n\n任務：\n1. 統計出現頻率最高、漲幅力道最強的前三大上漲題材。\n2. 為每個題材命名一個清晰的標籤（例如：AI 散熱族群、ASIC IC 設計、重電外銷）。\n3. 簡述該題材目前在台股市場的擴散程度（是剛起步還是已接近高點）。\n\n輸出格式：\n第一名題材： [標籤名稱] | 權重：[個股數量] | 描述：[為何強勢]\n第二名題材： ...\n第三名題材： ...`;

    const response = await ai.run('@cf/meta/llama-3.1-8b-instruct', {
        messages: [
            { role: "system", content: "你是一個專業的市場趨勢分析師。" },
            { role: "user", content: prompt }
        ]
    });

    return response.response;
}

export async function runStage3(top3Themes: string, ai: any): Promise<string> {
    const prompt = `針對目前台股最強勢的三大題材：\n${top3Themes}\n\n請執行以下推論：\n\n任務：\n1. 識別領頭羊： 在已知清單中，哪幾支股票是該題材的「純度最高」指標股？\n2. 挖掘外溢股 (Spillover)： 請運用你的資料庫，找出具備相同題材、且屬於「同產業鏈上下游」或「同族群」但目前漲幅尚未明顯落後、或位階較低的潛力股。\n3. 邏輯說明： 說明為什麼這幾支股票具備補漲潛力。\n\n輸出格式：\n題材 A： [領頭羊] -> [建議關注的外溢股] | 推薦理由：[...] \n題材 B： ...`;

    const response = await ai.run('@cf/meta/llama-3.1-8b-instruct', {
        messages: [
            { role: "system", content: "你是一個深耕台股產業鏈的投資專家。" },
            { role: "user", content: prompt }
        ]
    });

    return response.response;
}
