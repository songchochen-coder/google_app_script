export interface Stock {
    ticker: string;
    name: string;
    industry: string;
    monthly_change: number;
}

export async function fetchTopGainers(): Promise<Stock[]> {
    const url = "https://scanner.tradingview.com/taiwan/scan";

    const body = {
        filter: [
            { left: "change|30", operation: "greater", right: 20 },
            { left: "type", operation: "in_range", right: ["stock", "dr", "fund"] },
            { left: "subtype", operation: "in_range", right: ["common", "foreign-issuers", ""] },
            { left: "is_primary", operation: "equal", right: true }
        ],
        options: { lang: "zh" },
        markets: ["taiwan"],
        symbols: { query: { types: [] }, tickers: [] },
        columns: ["name", "description", "industry", "change|30"],
        sort: { sortBy: "change|30", sortOrder: "desc" },
        range: [0, 15] // Limit to top 15 to stay within worker time limits
    };

    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        throw new Error(`TradingView API failed: ${response.statusText}`);
    }

    const data = await response.json() as any;

    return data.data.map((item: any) => ({
        ticker: item.s.split(":")[1],
        name: item.d[1],
        industry: item.d[2] || "未知",
        monthly_change: item.d[3]
    }));
}
