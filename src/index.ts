import { fetchTopGainers } from './tradingview';
import { runStage1, runStage2, runStage3 } from './ai_analysis';

export interface Env {
    AI: any;
    TAVILY_API_KEY?: string;
}

export default {
    async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
        const url = new URL(request.url);
        if (url.pathname !== "/") {
            return new Response("Not Found", { status: 404 });
        }

        try {
            console.log("Fetching top gainers from TradingView...");
            const stocks = await fetchTopGainers();

            if (stocks.length === 0) {
                return new Response(JSON.stringify({ error: "No stocks found with > 20% growth" }), {
                    status: 404,
                    headers: { "Content-Type": "application/json" }
                });
            }

            console.log(`Analyzing ${stocks.length} stocks. Running Stage 1...`);
            const stage1 = await runStage1(stocks, env.AI, env.TAVILY_API_KEY || "");

            console.log("Running Stage 2...");
            const stage2 = await runStage2(stage1, env.AI);

            console.log("Running Stage 3...");
            const stage3 = await runStage3(stage2, env.AI);

            const result = {
                timestamp: new Date().toISOString(),
                stocks_analyzed: stocks.map(s => `${s.ticker} ${s.name} (${s.monthly_change.toFixed(2)}%)`),
                stage1_analysis: stage1,
                stage2_summary: stage2,
                stage3_spillover: stage3
            };

            return new Response(JSON.stringify(result, null, 2), {
                headers: { "Content-Type": "application/json; charset=utf-8" }
            });

        } catch (error: any) {
            console.error("Workflow Error:", error);
            return new Response(JSON.stringify({ error: error.message }), {
                status: 500,
                headers: { "Content-Type": "application/json" }
            });
        }
    },
};
