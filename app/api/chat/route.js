import { NextResponse } from "next/server";
import { validateAssumptions, computeDCF } from "../../../lib/dcf";
import { validateAssumptions as validateLBOAssumptions, computeLBO } from "../../../lib/lbo";
import { validateAssumptions as validateUEAssumptions, computeUnitEconomics } from "../../../lib/unitEconomics";
import { validateAssumptions as validateCapTableAssumptions, computeCapTable } from "../../../lib/capTable";

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

// Used only for LBO requests. Claude extracts raw assumptions from the
// conversation — it never calculates anything; lib/lbo.js does all the math.
const LBO_ASSUMPTIONS_PROMPT = `You are a financial data extraction engine. Read the user's request and return ONLY a single raw JSON object (no markdown, no code fences, no commentary, no explanation) with exactly these fields:

{
  "companyName": string,
  "entryRevenue": number,          // most recent full-year revenue in dollars, at entry
  "revenueGrowth": [number, number, number, number, number], // YoY growth rate for each of the next 5 years, as decimals (0.05 means 5%)
  "ebitdaMargin": number,          // stabilized EBITDA margin, decimal
  "daPct": number,                 // D&A as a % of revenue, decimal
  "capexPct": number,              // CapEx as a % of revenue, decimal
  "nwcPct": number,                // change in net working capital as a % of revenue, decimal
  "taxRate": number,               // decimal
  "entryMultiple": number,         // entry EV / EBITDA multiple
  "exitMultiple": number,          // exit EV / EBITDA multiple
  "holdingPeriod": number,         // years, 1-5, default 5 if not specified
  "feePct": number,                // transaction fees as a % of enterprise value, decimal
  "minCash": number,               // minimum cash balance held on the balance sheet, dollars
  "seniorDebt": { "multiple": number, "rate": number, "amortPct": number }, // x EBITDA, interest rate decimal, mandatory annual amortization % of original principal (decimal)
  "subDebt": { "multiple": number, "rate": number },  // x EBITDA, interest rate decimal — no amortization, bullet at exit
  "sweepPct": number               // % of excess free cash flow swept to pay down senior debt, decimal
}

Do not perform any calculations and do not include any derived figures — no sources & uses, debt schedule, free cash flow, enterprise value, MOIC, or IRR. Only the raw input assumptions above, inferred or reasonably estimated from the user's request. If the user didn't specify a value, make a sensible estimate rather than omitting the field.`;

const LBO_INTERP_SYSTEM_PROMPT = `You are a financial analyst writing the interpretation section of an LBO model that has already been fully calculated by deterministic code. You will be given the exact computed outputs (sources & uses, exit valuation, MOIC, and IRR). Write a short interpretation — 2 to 4 sentences — of what these results mean for the deal.

Use ONLY the numbers provided to you. Do not recalculate, round differently, estimate, or invent any figure that isn't given to you. Do not restate the raw assumptions. Be direct, with no fluff or disclaimers.`;

const LBO_INTERP_SYSTEM_PROMPT_SIMPLE = LBO_INTERP_SYSTEM_PROMPT +
  "\n\nIMPORTANT: Explain it in simple terms a teenager could understand. Avoid financial jargon, use a relatable analogy, and keep it conversational.";

// Used only for unit economics requests. Claude extracts raw assumptions
// from the conversation — it never calculates anything; lib/unitEconomics.js
// does all the math.
const UE_ASSUMPTIONS_PROMPT = `You are a financial data extraction engine. Read the user's request and return ONLY a single raw JSON object (no markdown, no code fences, no commentary, no explanation) with exactly these fields:

{
  "companyName": string,
  "arpu": number,                  // average monthly revenue per customer, dollars
  "grossMargin": number,           // decimal (0.75 means 75%)
  "monthlyChurn": number,          // decimal (0.03 means 3% of customers lost per month)
  "cac": number,                   // fully-loaded customer acquisition cost, dollars
  "monthlyDiscountRate": number,   // decimal, default 0.01 if not specified
  "projectionMonths": number,      // default 36 if not specified
  "notes": string                  // explain any unit conversion you made, otherwise ""
}

If the user gives annual churn instead of monthly, convert it: monthlyChurn = 1 - (1 - annualChurn)^(1/12). If they give retention instead of churn, churn = 1 - retention (converting to monthly first if the retention figure is annual). If they give annual revenue instead of monthly ARPU, divide by 12. Whenever you convert a number, describe exactly what you converted and the resulting value in the "notes" field (e.g. "Converted 30% annual churn to 2.96% monthly churn"). If nothing needed converting, set "notes" to "".

Do not perform any calculations beyond those unit conversions — no LTV, CAC payback, or LTV/CAC ratio. Only the raw input assumptions above, inferred or reasonably estimated from the user's request. If the user didn't specify a value, make a sensible estimate rather than omitting the field.`;

const UE_INTERP_SYSTEM_PROMPT = `You are a financial analyst writing the interpretation section of a unit economics model that has already been fully calculated by deterministic code. You will be given the exact computed outputs (LTV, LTV/CAC, CAC payback period). Write a short interpretation — 2 to 4 sentences — of what these results mean for the business.

Use ONLY the numbers provided to you. Do not recalculate, round differently, estimate, or invent any figure that isn't given to you. Do not restate the raw assumptions. If a "notes" field describing a unit conversion is present, mention it briefly. Be direct, with no fluff or disclaimers.`;

const UE_INTERP_SYSTEM_PROMPT_SIMPLE = UE_INTERP_SYSTEM_PROMPT +
  "\n\nIMPORTANT: Explain it in simple terms a teenager could understand. Avoid financial jargon, use a relatable analogy, and keep it conversational.";

// Used only for cap table requests. Claude extracts raw assumptions from the
// conversation — it never calculates anything; lib/capTable.js does all the math.
const CAPTABLE_ASSUMPTIONS_PROMPT = `You are a financial data extraction engine. Read the user's request and return ONLY a single raw JSON object (no markdown, no code fences, no commentary, no explanation) with exactly these fields:

{
  "companyName": string,
  "founders": [ { "name": string, "shares": number } ],
  "optionPoolShares": number,   // existing option pool share count, 0 if none
  "rounds": [
    {
      "name": string,                     // e.g. "Seed", "Series A"
      "preMoney": number,                 // dollars
      "investment": number,               // dollars
      "targetOptionPoolPct": number,      // decimal, target % of post-money fully-diluted cap the option pool should represent after this round; 0 if no pool refresh
      "liquidationPreference": number     // multiple, default 1 (1x non-participating) if not specified
    }
  ],
  "safes": [
    { "name": string, "amount": number, "cap": number, "discount": number }  // cap is a pre-money valuation cap in dollars, discount is a decimal (0.20 means 20%); omit entirely (empty array) if the user didn't mention any SAFEs
  ]
}

List rounds in chronological order. All SAFEs are assumed to convert in the first priced round. Do not perform any calculations and do not include any derived figures — no price per share, dilution, ownership percentages, or exit values. Only the raw input assumptions above, inferred or reasonably estimated from the user's request. If the user didn't specify a value, make a sensible estimate rather than omitting the field.`;

const CAPTABLE_INTERP_SYSTEM_PROMPT = `You are a financial analyst writing the interpretation section of a cap table model that has already been fully calculated by deterministic code. You will be given the exact computed outputs (price per share, dilution, and final ownership percentages). Write a short interpretation — 2 to 4 sentences — of what these results mean for the founders and investors.

Use ONLY the numbers provided to you. Do not recalculate, round differently, estimate, or invent any figure that isn't given to you. Do not restate the raw assumptions. Be direct, with no fluff or disclaimers.`;

const CAPTABLE_INTERP_SYSTEM_PROMPT_SIMPLE = CAPTABLE_INTERP_SYSTEM_PROMPT +
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

function isLBORequest(messages) {
  const lastUser = [...messages].reverse().find((m) => m.role === "user");
  const t = (lastUser?.content || "").toLowerCase();
  return /\blbo\b/.test(t) || t.includes("leveraged buyout");
}

function isUnitEconomicsRequest(messages) {
  const lastUser = [...messages].reverse().find((m) => m.role === "user");
  const t = (lastUser?.content || "").toLowerCase();
  return t.includes("unit economics") || (/\bltv\b/.test(t) && /\bcac\b/.test(t)) || t.includes("cac payback");
}

function isCapTableRequest(messages) {
  const lastUser = [...messages].reverse().find((m) => m.role === "user");
  const t = (lastUser?.content || "").toLowerCase();
  return t.includes("cap table") || t.includes("capitalization table");
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

// Small, focused summary so the interpretation prompt only ever sees the
// final calculated LBO outputs — never the raw assumptions it could "recompute".
function summarizeLBOForInterpretation(assumptions, computed) {
  const exitYear = computed.years[assumptions.holdingPeriod - 1];
  return {
    companyName: assumptions.companyName,
    holdingPeriodYears: assumptions.holdingPeriod,
    entryEnterpriseValue: round2(computed.sourcesUses.entryEV),
    sponsorEquityAtEntry: round2(computed.sourcesUses.sponsorEquity),
    seniorDebtAtEntry: round2(computed.sourcesUses.seniorDebt),
    subDebtAtEntry: round2(computed.sourcesUses.subDebt),
    exitEbitda: round2(exitYear.ebitda),
    exitEnterpriseValue: round2(computed.exit.exitEV),
    exitNetDebt: round2(computed.exit.exitNetDebt),
    exitEquityValue: round2(computed.exit.exitEquityValue),
    sponsorProceeds: round2(computed.exit.sponsorProceeds),
    moic: computed.valid ? round2(computed.returns.moic) : null,
    irrPct: computed.valid ? round2(computed.returns.irr * 100) : null,
    dealNote: computed.valid
      ? null
      : "Sponsor equity was zero or negative at entry (the debt tranches plus fees exceeded the purchase price), so this capital structure doesn't work as specified.",
  };
}

// Small, focused summary so the interpretation prompt only ever sees the
// final calculated unit economics outputs — never the raw assumptions.
function summarizeUEForInterpretation(assumptions, computed) {
  return {
    companyName: assumptions.companyName,
    grossProfitPerMonth: round2(computed.grossProfitPerMonth),
    avgLifetimeMonths: computed.avgLifetimeMonths !== null ? round2(computed.avgLifetimeMonths) : null,
    ltv: round2(computed.simpleLTV),
    discountedLtv: round2(computed.discountedLTV),
    ltvToCac: computed.ltvToCac !== null ? round2(computed.ltvToCac) : null,
    cacPaybackMonths: computed.cacPaybackMonths !== null ? round2(computed.cacPaybackMonths) : null,
    paybackNote: computed.cacPaybackMonths === null
      ? `CAC is not recovered within the ${assumptions.projectionMonths}-month projection window.`
      : null,
    notes: assumptions.notes || null,
  };
}

// Small, focused summary so the interpretation prompt only ever sees the
// final calculated cap table outputs — never the raw assumptions.
function summarizeCapTableForInterpretation(assumptions, computed) {
  const lastRound = computed.rounds[computed.rounds.length - 1];
  return {
    companyName: assumptions.companyName,
    totalRaised: round2(computed.rounds.reduce((s, r) => s + r.investment, 0)),
    lastRoundName: lastRound?.name ?? null,
    lastRoundPricePerShare: lastRound ? round2(lastRound.pricePerShare) : null,
    lastRoundPostMoney: lastRound ? round2(lastRound.postMoney) : null,
    safeConversions: computed.rounds[0]?.safeConversions?.map((s) => ({ name: s.name, shares: Math.round(s.shares) })) ?? [],
    finalOwnershipPct: computed.finalOwnership.map((h) => ({ name: h.name, pct: round2(h.pct * 100) })),
  };
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

    if (isLBORequest(messages)) {
      // 1. Claude extracts raw assumptions only — no math.
      const assumptionsText = await callClaude(LBO_ASSUMPTIONS_PROMPT, messages, 800);
      const assumptions = validateLBOAssumptions(extractJSON(assumptionsText));

      // 2. lib/lbo.js performs every calculation deterministically.
      const computed = computeLBO(assumptions);

      // 3. Claude writes a short interpretation of the computed numbers only.
      const summary = summarizeLBOForInterpretation(assumptions, computed);
      const interpretation = await callClaude(
        simpleMode ? LBO_INTERP_SYSTEM_PROMPT_SIMPLE : LBO_INTERP_SYSTEM_PROMPT,
        [{
          role: "user",
          content: `Here are the fully calculated LBO outputs. Write the interpretation using only these numbers:\n\n${JSON.stringify(summary, null, 2)}`,
        }],
        400
      );

      return NextResponse.json({
        kind: "lbo",
        assumptions,
        interpretation: interpretation.trim() || "Model calculated — see the panel for the full breakdown.",
      });
    }

    if (isUnitEconomicsRequest(messages)) {
      // 1. Claude extracts raw assumptions only — no math.
      const assumptionsText = await callClaude(UE_ASSUMPTIONS_PROMPT, messages, 600);
      const assumptions = validateUEAssumptions(extractJSON(assumptionsText));

      // 2. lib/unitEconomics.js performs every calculation deterministically.
      const computed = computeUnitEconomics(assumptions);

      // 3. Claude writes a short interpretation of the computed numbers only.
      const summary = summarizeUEForInterpretation(assumptions, computed);
      const interpretation = await callClaude(
        simpleMode ? UE_INTERP_SYSTEM_PROMPT_SIMPLE : UE_INTERP_SYSTEM_PROMPT,
        [{
          role: "user",
          content: `Here are the fully calculated unit economics outputs. Write the interpretation using only these numbers:\n\n${JSON.stringify(summary, null, 2)}`,
        }],
        400
      );

      return NextResponse.json({
        kind: "unitEconomics",
        assumptions,
        interpretation: interpretation.trim() || "Model calculated — see the panel for the full breakdown.",
      });
    }

    if (isCapTableRequest(messages)) {
      // 1. Claude extracts raw assumptions only — no math.
      const assumptionsText = await callClaude(CAPTABLE_ASSUMPTIONS_PROMPT, messages, 900);
      const assumptions = validateCapTableAssumptions(extractJSON(assumptionsText));

      // 2. lib/capTable.js performs every calculation deterministically.
      const computed = computeCapTable(assumptions);

      // 3. Claude writes a short interpretation of the computed numbers only.
      const summary = summarizeCapTableForInterpretation(assumptions, computed);
      const interpretation = await callClaude(
        simpleMode ? CAPTABLE_INTERP_SYSTEM_PROMPT_SIMPLE : CAPTABLE_INTERP_SYSTEM_PROMPT,
        [{
          role: "user",
          content: `Here are the fully calculated cap table outputs. Write the interpretation using only these numbers:\n\n${JSON.stringify(summary, null, 2)}`,
        }],
        400
      );

      return NextResponse.json({
        kind: "capTable",
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
