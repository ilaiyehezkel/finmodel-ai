import { NextResponse } from "next/server";


const SYSTEM_PROMPT = `You are a world-class financial analyst and modeling expert. You specialize in:

- DCF (Discounted Cash Flow) analysis
- LBO (Leveraged Buyout) modeling
- Comparable company analysis (Comps)
- Revenue forecasting and scenario modeling
- Cap table modeling and dilution analysis
- Unit economics (LTV, CAC, payback period, contribution margin)
- Financial statement modeling (3-statement models)
- Sensitivity analysis and data tables
- Startup valuation (pre-revenue, seed, Series A+)
- Real estate financial modeling (NOI, cap rates, IRR)

Your responses are:
- Precise and numbers-driven. Always show formulas and calculations explicitly.
- Structured with clear sections (Assumptions, Model, Output, Interpretation).
- Formatted using markdown tables when presenting financial data.
- Practical, explaining what inputs the user needs and what the outputs mean.
- Direct, with no fluff or unnecessary disclaimers.

When a user asks for a model, you:
1. State key assumptions clearly
2. Show the model/calculations step by step
3. Present results in a clean table or structured format
4. Offer 2-3 sensitivity scenarios (bear/base/bull or low/mid/high)
5. Give a 1-2 sentence interpretation of what the numbers mean

Always ask clarifying questions if critical inputs are missing rather than making silent assumptions.`;

export async function POST(req) {
  try {
    const { messages, simpleMode } = await req.json();

    if (!process.env.ANTHROPIC_API_KEY) {
      return NextResponse.json(
        { error: "Server is missing ANTHROPIC_API_KEY." },
        { status: 500 }
      );
    }

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 2000,
        system: simpleMode
          ? SYSTEM_PROMPT + "\n\nIMPORTANT: Explain everything in simple terms a teenager can understand. Avoid financial jargon. Use relatable everyday analogies and concrete examples. Keep language conversational and clear."
          : SYSTEM_PROMPT,
        messages,
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      return NextResponse.json(
        { error: `Anthropic API error: ${errText}` },
        { status: response.status }
      );
    }

    const data = await response.json();
    const reply = data.content?.[0]?.text || "No response generated.";

    return NextResponse.json({ reply });
  } catch (err) {
    return NextResponse.json(
      { error: "Internal server error: " + err.message },
      { status: 500 }
    );
  }
}
