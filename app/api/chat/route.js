import { NextResponse } from "next/server";
import { validateAssumptions, computeDCF } from "../../../lib/dcf";

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

// Used only for DCF requests. Claude extracts raw assumptions from the
// conversation — it never calculates anything; lib/dcf.js does all the math.
const DCF_ASSUMPTIONS_PROMPT = `You are a financial data extraction engine. Read the user's request and return ONLY a single raw JSON object (no markdown, no code fences, no commentary, no explanation) with exactly these fields:

{
  "companyName": string,
  "baseRevenue": number,          // most recent full-year revenue in dollars
  "revenueGrowth": [number, number, number, number, number], // YoY growth rate for each of the next 5 years, as decimals (0.6 means 60%)
  "ebitdaMargin": number,         // stabilized EBITDA margin, decimal
  "taxRate": number,              // decimal
  "daPct": number,                // D&A as a % of revenue, decimal
  "capexPct": number,             // CapEx as a % of revenue, decimal
  "nwcPct": number,               // change in net working capital as a % of revenue, decimal
  "wacc": number,                 // decimal
  "terminalGrowth": number,       // decimal
  "netDebt": number,              // dollars, 0 if none or unknown
  "sharesOutstanding": number     // if unknown, estimate a reasonable figure for a company of this size
}

Do not perform any calculations and do not include any derived figures — no revenue projections, EBITDA, free cash flow, enterprise value, terminal value, or share price. Only the raw input assumptions above, inferred or reasonably estimated from the user's request. If the user didn't specify a value, make a sensible estimate rather than omitting the field.`;

const INTERP_SYSTEM_PROMPT = `You are a financial analyst writing the interpretation section of a DCF model that has already been fully calculated by deterministic code. You will be given the exact computed outputs (enterprise value, equity value, implied share price, and key year-5 figures). Write a short interpretation — 2 to 4 sentences — of what these results mean for the business.

Use ONLY the numbers provided to you. Do not recalculate, round differently, estimate, or invent any figure that isn't given to you. Do not restate the raw assumptions. Be direct, with no fluff or disclaimers.`;

const INTERP_SYSTEM_PROMPT_SIMPLE = INTERP_SYSTEM_PROMPT +
  "\n\nIMPORTANT: Explain it in simple terms a teenager could understand. Avoid financial jargon, use a relatable analogy, and keep it conversational.";

async function callClaude(system, messages, maxTokens) {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: maxTokens,
      system,
      messages,
    }),
  });

  if (!response.ok) {
    throw new Error(`Claude API request failed (${response.status})`);
  }

  const data = await response.json();
  return data.content?.[0]?.text || "";
}

function isDCFRequest(messages) {
  const lastUser = [...messages].reverse().find((m) => m.role === "user");
  const t = (lastUser?.content || "").toLowerCase();
  return /\bdcf\b/.test(t) || t.includes("discounted cash flow");
}

function extractJSON(text) {
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch {
    return null;
  }
}

// Small, focused summary so the interpretation prompt only ever sees the
// final calculated numbers — never the raw assumptions it could "recompute".
function summarizeForInterpretation(assumptions, computed) {
  const lastYear = computed.years[computed.years.length - 1];
  return {
    companyName: assumptions.companyName,
    year5Revenue: round2(lastYear.revenue),
    year5Ebitda: round2(lastYear.ebitda),
    year5UnleveredFCF: round2(lastYear.ufcf),
    sumOfPVYear1to5: round2(computed.sumPV),
    terminalValue: computed.terminalValid ? round2(computed.terminalValue) : null,
    enterpriseValue: computed.terminalValid ? round2(computed.enterpriseValue) : null,
    netDebt: round2(assumptions.netDebt),
    equityValue: computed.terminalValid ? round2(computed.equityValue) : null,
    sharesOutstanding: assumptions.sharesOutstanding,
    impliedSharePrice: computed.terminalValid ? round2(computed.impliedSharePrice) : null,
    valuationNote: computed.terminalValid
      ? null
      : "WACC was not greater than the terminal growth rate, so enterprise/equity value and implied share price could not be calculated.",
  };
}

function round2(n) {
  return typeof n === "number" ? Math.round(n * 100) / 100 : n;
}

export async function POST(req) {
  try {
    const { messages, simpleMode } = await req.json();

    if (!process.env.ANTHROPIC_API_KEY) {
      return NextResponse.json(
        { error: "Server is missing ANTHROPIC_API_KEY." },
        { status: 500 }
      );
    }

    if (isDCFRequest(messages)) {
      // 1. Claude extracts raw assumptions only — no math.
      const assumptionsText = await callClaude(DCF_ASSUMPTIONS_PROMPT, messages, 700);
      const assumptions = validateAssumptions(extractJSON(assumptionsText));

      // 2. lib/dcf.js performs every calculation deterministically.
      const computed = computeDCF(assumptions);

      // 3. Claude writes a short interpretation of the computed numbers only.
      const summary = summarizeForInterpretation(assumptions, computed);
      const interpretation = await callClaude(
        simpleMode ? INTERP_SYSTEM_PROMPT_SIMPLE : INTERP_SYSTEM_PROMPT,
        [{
          role: "user",
          content: `Here are the fully calculated DCF outputs. Write the interpretation using only these numbers:\n\n${JSON.stringify(summary, null, 2)}`,
        }],
        400
      );

      return NextResponse.json({
        kind: "dcf",
        assumptions,
        interpretation: interpretation.trim() || "Model calculated — see the panel for the full breakdown.",
      });
    }

    // ── Every other model type: unchanged single-call behavior ──
    const reply = await callClaude(
      simpleMode
        ? SYSTEM_PROMPT + "\n\nIMPORTANT: Explain everything in simple terms a teenager can understand. Avoid financial jargon. Use relatable everyday analogies and concrete examples. Keep language conversational and clear."
        : SYSTEM_PROMPT,
      messages,
      2000
    );

    return NextResponse.json({ reply: reply || "No response generated." });
  } catch (err) {
    return NextResponse.json(
      { error: "Something went wrong. Please try again." },
      { status: 500 }
    );
  }
}
