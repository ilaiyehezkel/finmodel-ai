"use client";

import { useState, useRef, useEffect } from "react";
import { formatMessage } from "../lib/format";
import * as XLSXStyle from "xlsx-js-style";
import { computeDCF, computeSensitivity } from "../lib/dcf";

// ── Templates ───────────────────────────────────────────────────

const TEMPLATES = [
  {
    label: "DCF",
    prompt: `Build a complete DCF model with these assumptions:
- Company type: SaaS
- Year 1 Revenue: $5M, growing 60% / 45% / 30% / 20% / 15% over 5 years
- Gross margin: 75% | EBITDA margin (stabilized): 20%
- CapEx: 3% of revenue | NWC change: 2% of revenue delta
- WACC: 12% | Terminal growth rate: 3%
- Net debt: $0

Show: 5-year DCF, terminal value, enterprise value, equity value, and a sensitivity table for WACC (10–14%) vs. terminal growth rate (2–4%).`,
  },
  {
    label: "Cap Table",
    prompt: `Model a startup cap table through Series A:
- 2 founders, 2,000,000 shares each at founding
- Employee option pool: 10% of fully diluted (pre-Series A)
- Seed: $2M raised at $8M pre-money
- Series A: $10M raised at $40M pre-money + 10% post-money option top-up

Show share counts, ownership %, and dilution at each round for all stakeholders.`,
  },
  {
    label: "LBO",
    prompt: `Build a leveraged buyout model:
- Target EBITDA: $20M | Purchase price: 8x EBITDA
- Debt: 60% of purchase price, Term Loan B at SOFR+400bps, 1% annual amortization
- Revenue growth: 5%/yr | EBITDA margin: +1%/yr improvement
- Management fees: $1M/yr | Tax rate: 25%
- Exit: 5 years at 7x, 8x, and 9x EBITDA

Show: sources & uses, debt schedule, P&L, cash flow, MoM and IRR at each exit multiple.`,
  },
  {
    label: "Unit Economics",
    prompt: `Calculate unit economics for a B2B SaaS company:
- MRR per customer: $500 | Gross margin: 80%
- Monthly churn: 2% | Sales cycle: 45 days
- Fully loaded CAC: $3,000
- Marketing: $20K/month generating 50 MQLs

Calculate LTV, LTV/CAC, CAC payback period, contribution margin, and Rule of 40. Show sensitivity on churn (1%, 2%, 3%, 4%).`,
  },
  {
    label: "3-Statement",
    prompt: `Build a 3-statement model for a retail company:
- Year 1 Revenue: $10M, growing 15%/yr for 3 years
- COGS: 60% | OpEx: $2.5M fixed + 5% variable | D&A: $300K/yr
- CapEx: $500K/yr | AR: 30 days | Inventory: 45 days | AP: 30 days
- Opening cash: $1M | Existing debt: $2M at 6% | Tax rate: 25%

Produce a linked Income Statement, Balance Sheet, and Cash Flow Statement for 3 years. Ensure they tie.`,
  },
  {
    label: "Startup Valuation",
    prompt: `Value a pre-revenue B2B SaaS / HR Tech startup:
- Stage: Seed | Team: 3 ex-FAANG/McKinsey founders + 2 engineers
- Product: MVP live, 5 design partners with LOIs | TAM: $8B
- Comps: valued at 10–15x ARR | Raising: $2M seed

Apply: (1) Scorecard method, (2) Berkus method, (3) First Chicago method with 3 scenarios, (4) VC method at 10x target return / 5-yr exit. Recommend a valuation range.`,
  },
];

// ── Suggestions ─────────────────────────────────────────────────

const PRO_SUGGESTIONS = [
  "Build a DCF model for a SaaS company with $2M ARR growing 80% YoY",
  "Model a cap table for a startup raising a $3M seed at $12M pre-money",
  "Calculate LTV/CAC ratio given $50 monthly revenue, 85% retention, $200 CAC",
];

const SIMPLE_SUGGESTIONS = [
  "Walk me through a DCF model step by step",
  "Explain what a cap table is and build me a simple one",
  "What is LTV and CAC? Show me an example",
];

// ── Table utilities ─────────────────────────────────────────────

function parseMarkdownTables(text) {
  const tables = [];
  const lines = text.split("\n");
  let i = 0;
  while (i < lines.length) {
    if (!lines[i].trim().startsWith("|")) { i++; continue; }
    const block = [];
    while (i < lines.length && lines[i].trim().startsWith("|")) {
      block.push(lines[i].trim());
      i++;
    }
    if (block.length < 3) continue;
    if (!/^\|[\s\-:|]+\|$/.test(block[1])) continue;
    const parseRow = (l) =>
      l.split("|").filter((_, idx, a) => idx > 0 && idx < a.length - 1).map((c) => c.trim());
    const headers = parseRow(block[0]);
    const rows = block.slice(2).map(parseRow);
    if (headers.length && rows.length) tables.push({ headers, rows });
  }
  return tables;
}

function stripTables(text) {
  return text.replace(/(\|.+\|\n?)+/g, "").replace(/\n{3,}/g, "\n\n").trim();
}

// ── Row / cell classification for the spreadsheet view ──────────

const SECTION_KEYWORDS = [
  "revenue", "ebitda", "ebit", "gross profit", "operating income",
  "free cash flow", "net income",
];

const TOTAL_KEYWORDS = [
  "total", "subtotal", "sum", "enterprise value", "equity value",
  "terminal value", "purchase price", "irr", "moic", "sources", "uses",
  "valuation", "implied value",
];

function rowClass(row) {
  const text = row.join(" ").toLowerCase().replace(/\*/g, "");
  if (TOTAL_KEYWORDS.some((kw) => text.includes(kw))) return "row-total";
  if (SECTION_KEYWORDS.some((kw) => text.includes(kw))) return "row-section";
  return "";
}

function displayCell(raw) {
  const s = String(raw ?? "").trim().replace(/\*\*/g, "");
  if (!s || s === "-" || s === "—" || s.toLowerCase() === "n/a") return { text: s, numeric: false };

  // Already parenthesized, e.g. "(1,234)"
  if (/^\(.+\)$/.test(s) && /\d/.test(s)) return { text: s, numeric: true };

  const isNumericLike = /^-?\$?-?[\d,.]+%?[xX]?$/.test(s.replace(/\s/g, ""));
  if (!isNumericLike) return { text: s, numeric: false };

  const isNegative = /^-/.test(s) || /\$-/.test(s);
  if (isNegative) {
    const positive = s.replace("-", "");
    return { text: `(${positive})`, numeric: true };
  }
  return { text: s, numeric: true };
}

// ── Excel export colours & constants ───────────────────────────

const CLR = {
  teal:      "14B8A6",
  tealSoft:  "CFF5EF",
  tealDark:  "0B5E54",
  black:     "000000",
  white:     "FFFFFF",
  rowAlt:    "F2F2F5",
  dark:      "1A1A2A",
  mid:       "60606E",
  subBg:     "111113",
};

const SUMMARY_KEYWORDS = [
  "total", "subtotal", "sum", "net", "gross",
  "ebitda", "ebit", "ebita",
  "enterprise value", "equity value",
  "irr", "moic", "roi",
  "valuation", "implied value", "purchase price",
  "sources", "uses", "terminal value",
];

function detectModelType(text) {
  const t = text.toLowerCase();
  if (/\bdcf\b/.test(t) || t.includes("discounted cash flow")) return "DCF Model";
  if (/\blbo\b/.test(t) || t.includes("leveraged buyout"))     return "LBO Model";
  if (t.includes("cap table") || t.includes("capitalization table")) return "Cap Table";
  if (t.includes("unit economics") || (t.includes("ltv") && t.includes("cac"))) return "Unit Economics";
  if (t.includes("3-statement") || t.includes("three-statement") ||
    (t.includes("income statement") && t.includes("balance sheet")))  return "3-Statement Model";
  if (t.includes("startup valuation") || t.includes("berkus") ||
    t.includes("scorecard method") || t.includes("pre-revenue"))      return "Startup Valuation";
  return "Financial Model";
}

function isSummaryRow(row) {
  const text = row.join(" ").toLowerCase().replace(/\*/g, "");
  return SUMMARY_KEYWORDS.some((kw) => text.includes(kw));
}

function parseCell(raw) {
  const s = String(raw ?? "").trim().replace(/\*\*/g, "");
  if (!s || s === "-" || s === "—" || s.toLowerCase() === "n/a") return { v: s, t: "s" };

  // Percentage: "12.5%", "-5.0%"
  if (/^-?[\d,.]+\s*%$/.test(s)) {
    const n = parseFloat(s.replace(/[,%\s]/g, ""));
    if (!isNaN(n)) return { v: n / 100, t: "n", numFmt: "0.0%" };
  }

  // Currency with $
  if (s.includes("$")) {
    const neg = s.includes("(") || s.startsWith("-");
    const raw2 = parseFloat(s.replace(/[$,()\s]/g, "").replace(/,/g, ""));
    if (!isNaN(raw2)) return { v: neg && raw2 > 0 ? -raw2 : raw2, t: "n", numFmt: '"$"#,##0;("$"#,##0)' };
  }

  // Multiple: "2.5x"
  if (/^-?[\d,.]+[xX]$/.test(s)) {
    const n = parseFloat(s.slice(0, -1).replace(/,/g, ""));
    if (!isNaN(n)) return { v: n, t: "n", numFmt: '0.0"x"' };
  }

  // Negative in parens without $: "(1,234)"
  if (/^\([\d,.]+\)$/.test(s)) {
    const n = parseFloat(s.replace(/[(),]/g, "").replace(/,/g, ""));
    if (!isNaN(n)) return { v: -n, t: "n", numFmt: "#,##0" };
  }

  // Plain number: "1,234,567" or "123.45" or "-500"
  const numStr = s.replace(/,/g, "");
  if (/^-?\d+\.?\d*$/.test(numStr)) {
    const n = parseFloat(numStr);
    return { v: n, t: "n", numFmt: numStr.includes(".") ? "#,##0.00" : "#,##0" };
  }

  return { v: s, t: "s" };
}

function exportToExcel(text) {
  const tables = parseMarkdownTables(text);
  if (!tables.length) { alert("No tables found in this response."); return; }

  const modelType = detectModelType(text);
  const wb = XLSXStyle.utils.book_new();
  const ws = {};
  ws["!merges"] = [];

  const maxCols = Math.max(...tables.map((t) => t.headers.length));
  const colMax = new Array(maxCols).fill(8);
  let R = 0;

  function setCell(r, c, cell) {
    ws[XLSXStyle.utils.encode_cell({ r, c })] = cell;
    const len = String(cell.v ?? "").length;
    if (c < colMax.length) colMax[c] = Math.max(colMax[c], len);
  }

  // ── Title row ─────────────────────────────────────────────────
  for (let c = 0; c < maxCols; c++) {
    setCell(R, c, {
      v: c === 0 ? modelType : "",
      t: "s",
      s: {
        fill: { fgColor: { rgb: CLR.teal } },
        font: { bold: true, sz: 16, color: { rgb: CLR.black } },
        alignment: { horizontal: "left", vertical: "center" },
      },
    });
  }
  ws["!merges"].push({ s: { r: 0, c: 0 }, e: { r: 0, c: maxCols - 1 } });
  R++;

  // ── Subtitle row ──────────────────────────────────────────────
  const dateStr = new Date().toLocaleDateString("en-US", {
    year: "numeric", month: "long", day: "numeric",
  });
  for (let c = 0; c < maxCols; c++) {
    setCell(R, c, {
      v: c === 0 ? `Generated by FinModel AI  ·  ${dateStr}` : "",
      t: "s",
      s: {
        fill: { fgColor: { rgb: CLR.subBg } },
        font: { italic: true, sz: 9, color: { rgb: CLR.mid } },
        alignment: { horizontal: "left", vertical: "center" },
      },
    });
  }
  ws["!merges"].push({ s: { r: 1, c: 0 }, e: { r: 1, c: maxCols - 1 } });
  R++;

  // ── Empty gap row ─────────────────────────────────────────────
  R++;

  // ── Tables ────────────────────────────────────────────────────
  tables.forEach((table, tIdx) => {
    const cols = table.headers.length;

    // Column header row
    for (let c = 0; c < cols; c++) {
      setCell(R, c, {
        v: table.headers[c],
        t: "s",
        s: {
          fill: { fgColor: { rgb: CLR.teal } },
          font: { bold: true, sz: 10, color: { rgb: CLR.black } },
          alignment: { horizontal: c === 0 ? "left" : "center", vertical: "center" },
          border: { bottom: { style: "medium", color: { rgb: "0E8C7F" } } },
        },
      });
    }
    R++;

    // Data rows
    table.rows.forEach((row, rIdx) => {
      const isSummary = isSummaryRow(row);
      const bg = isSummary ? CLR.tealSoft : rIdx % 2 === 0 ? CLR.white : CLR.rowAlt;

      for (let c = 0; c < cols; c++) {
        const cellVal = row[c] ?? "";
        const isFirst = c === 0;
        const parsed = parseCell(cellVal);

        const style = {
          fill: { fgColor: { rgb: bg } },
          font: {
            bold: isSummary || isFirst,
            sz: 10,
            color: { rgb: isSummary ? CLR.tealDark : CLR.dark },
          },
          alignment: {
            horizontal: isFirst ? "left" : parsed.t === "n" ? "right" : "center",
            vertical: "center",
          },
        };

        if (isSummary) {
          style.border = {
            top: { style: "thin", color: { rgb: "10A696" } },
            bottom: { style: "thin", color: { rgb: "10A696" } },
          };
        }

        const wsCell = { v: parsed.v, t: parsed.t, s: style };
        if (parsed.numFmt) wsCell.z = parsed.numFmt;
        setCell(R, c, wsCell);
      }
      R++;
    });

    // 2-row gap between tables
    if (tIdx < tables.length - 1) R += 2;
  });

  // ── Sheet metadata ─────────────────────────────────────────────
  ws["!ref"] = XLSXStyle.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: R - 1, c: maxCols - 1 },
  });

  ws["!cols"] = colMax.map((w) => ({ wch: Math.min(Math.max(w + 3, 12), 45) }));

  ws["!rows"] = [
    { hpt: 30 },  // title
    { hpt: 15 },  // subtitle
  ];

  // Freeze rows 0-2 (title, subtitle, gap) so column headers scroll with data
  ws["!freeze"] = { r: 3, c: 0 };

  const sheetName = modelType.slice(0, 31);
  XLSXStyle.utils.book_append_sheet(wb, ws, sheetName);
  XLSXStyle.writeFile(wb, `FinModel_${modelType.replace(/\s+/g, "_")}.xlsx`);
}

function exportToPDF(htmlContent) {
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
body{font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:1.7;color:#111;padding:40px 56px;max-width:860px;margin:0 auto}
h1{font-size:22px;margin:16px 0 6px}h2{font-size:18px;margin:14px 0 6px}h3{font-size:15px;color:#0f766e;margin:12px 0 6px}
p{margin:0 0 10px}
table{width:100%;border-collapse:collapse;margin:14px 0;font-size:12px;page-break-inside:avoid;font-family:"Courier New",monospace}
th{background:#f2f2f2;padding:8px 12px;border:1px solid #ccc;text-align:left;font-weight:600;font-size:11px;letter-spacing:.04em}
td{padding:7px 12px;border:1px solid #ddd}tr:nth-child(even) td{background:#fafafa}
code{background:#f4f4f4;padding:2px 5px;border-radius:3px;font-size:12px}
pre{background:#f4f4f4;padding:14px;border-radius:6px;margin:10px 0}pre code{background:none;padding:0}
strong{color:#000}ul{padding-left:18px;margin:8px 0}li{margin:3px 0}
.ph{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #0f766e;padding-bottom:12px;margin-bottom:24px}
.ph-title{color:#0f766e;font-size:18px;font-weight:600;margin:0}.ph-sub{color:#666;font-size:11px}
</style></head><body>
<div class="ph"><span class="ph-title">FinModel AI</span><span class="ph-sub">Financial Modeling Assistant</span></div>
${htmlContent}
</body></html>`;
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  const win = window.open(url);
  if (win) win.addEventListener("load", () => setTimeout(() => win.print(), 400));
  setTimeout(() => URL.revokeObjectURL(url), 15000);
}

// ── DCF interactive panel helpers ───────────────────────────────

function formatMoney(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  const str = "$" + abs.toLocaleString("en-US", { maximumFractionDigits: 0 });
  return n < 0 ? `(${str})` : str;
}

function formatPrice(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  const str = "$" + abs.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return n < 0 ? `(${str})` : str;
}

function formatPct(n, decimals = 1) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return (n * 100).toFixed(decimals) + "%";
}

// Converts canonical numeric assumptions (decimals/dollars) into the raw
// strings the editable inputs display and edit.
function fieldsFromAssumptions(a) {
  return {
    companyName: a.companyName,
    baseRevenue: String(a.baseRevenue),
    growth: a.revenueGrowth.map((g) => (g * 100).toString()),
    ebitdaMargin: (a.ebitdaMargin * 100).toString(),
    taxRate: (a.taxRate * 100).toString(),
    daPct: (a.daPct * 100).toString(),
    capexPct: (a.capexPct * 100).toString(),
    nwcPct: (a.nwcPct * 100).toString(),
    wacc: (a.wacc * 100).toString(),
    terminalGrowth: (a.terminalGrowth * 100).toString(),
    netDebt: String(a.netDebt),
    sharesOutstanding: String(a.sharesOutstanding),
  };
}

// Converts the editable field strings back into canonical numeric assumptions
// for lib/dcf.js. Unparseable input falls back to 0 rather than crashing.
function assumptionsFromFields(f) {
  const n = (s) => {
    const v = parseFloat(s);
    return Number.isFinite(v) ? v : 0;
  };
  return {
    companyName: f.companyName || "Company",
    baseRevenue: n(f.baseRevenue),
    revenueGrowth: f.growth.map((g) => n(g) / 100),
    ebitdaMargin: n(f.ebitdaMargin) / 100,
    taxRate: n(f.taxRate) / 100,
    daPct: n(f.daPct) / 100,
    capexPct: n(f.capexPct) / 100,
    nwcPct: n(f.nwcPct) / 100,
    wacc: n(f.wacc) / 100,
    terminalGrowth: n(f.terminalGrowth) / 100,
    netDebt: n(f.netDebt),
    sharesOutstanding: n(f.sharesOutstanding) || 1,
  };
}

function dcfSummaryText(a, computed) {
  const lines = [
    `${a.companyName} — DCF Model`,
    `Enterprise value: ${formatMoney(computed.enterpriseValue)}`,
    `Equity value: ${formatMoney(computed.equityValue)}`,
    `Implied share price: ${formatPrice(computed.impliedSharePrice)}`,
  ];
  if (!computed.terminalValid) {
    lines.push("(WACC must exceed the terminal growth rate to calculate a valuation.)");
  }
  return lines.join("\n");
}

// Builds a DCF workbook with real Excel formulas — every line item below the
// hard-coded (teal) assumption cells is a live formula, so editing an
// assumption in Excel recalculates the entire model.
function exportDCFToExcel(assumptions, sensitivity) {
  const a = assumptions;
  const wb = XLSXStyle.utils.book_new();
  const ws = {};
  ws["!merges"] = [];
  const col = (c) => XLSXStyle.utils.encode_col(c);
  const abs = (c, r) => `$${col(c)}$${r + 1}`;
  const rel = (c, r) => `${col(c)}${r + 1}`;
  const colMax = [28, 14, 14, 14, 14, 14];

  function setCell(r, c, cell) {
    ws[XLSXStyle.utils.encode_cell({ r, c })] = cell;
  }

  const inputStyle = (numFmt) => ({
    fill: { fgColor: { rgb: CLR.tealSoft } },
    font: { bold: true, sz: 10, color: { rgb: CLR.tealDark } },
    alignment: { horizontal: "right", vertical: "center" },
    ...(numFmt ? { z: numFmt } : {}),
  });
  const labelStyle = (bold) => ({
    font: { bold: !!bold, sz: 10, color: { rgb: CLR.dark } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  const formulaStyle = (bold, numFmt, totalBorder) => ({
    font: { bold: !!bold, sz: 10, color: { rgb: CLR.dark } },
    alignment: { horizontal: "right", vertical: "center" },
    ...(numFmt ? { z: numFmt } : {}),
    ...(totalBorder ? { border: { top: { style: "thin", color: { rgb: "10A696" } } } } : {}),
  });

  const setInput = (r, c, value, numFmt) =>
    setCell(r, c, { v: value, t: "n", s: inputStyle(numFmt) });
  const setLabel = (r, c, text, bold) =>
    setCell(r, c, { v: text, t: "s", s: labelStyle(bold) });
  const setFormula = (r, c, formula, numFmt, bold, totalBorder) =>
    setCell(r, c, { f: formula, t: "n", s: formulaStyle(bold, numFmt, totalBorder) });

  const PCT = "0.0%";
  const USD = '"$"#,##0';
  const USD2 = '"$"#,##0.00';
  const NUM0 = "#,##0";
  const DF = "0.000";

  let R = 0;

  // Title + subtitle
  for (let c = 0; c <= 5; c++) {
    setCell(R, c, {
      v: c === 0 ? `DCF Model — ${a.companyName}` : "",
      t: "s",
      s: { fill: { fgColor: { rgb: CLR.teal } }, font: { bold: true, sz: 16, color: { rgb: CLR.black } }, alignment: { horizontal: "left", vertical: "center" } },
    });
  }
  ws["!merges"].push({ s: { r: R, c: 0 }, e: { r: R, c: 5 } });
  R++;
  const dateStr = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  for (let c = 0; c <= 5; c++) {
    setCell(R, c, {
      v: c === 0 ? `Generated by FinModel AI  ·  ${dateStr}  ·  teal cells are AI-sourced inputs, everything else is a live formula` : "",
      t: "s",
      s: { fill: { fgColor: { rgb: CLR.subBg } }, font: { italic: true, sz: 9, color: { rgb: CLR.mid } }, alignment: { horizontal: "left", vertical: "center" } },
    });
  }
  ws["!merges"].push({ s: { r: R, c: 0 }, e: { r: R, c: 5 } });
  R += 2;

  // ── Assumptions block ──
  setLabel(R, 0, "Assumptions", true); R++;
  const rowBaseRevenue = R; setLabel(R, 0, "Base revenue ($)"); setInput(R, 1, a.baseRevenue, USD); R++;
  const rowEbitdaMargin = R; setLabel(R, 0, "EBITDA margin"); setInput(R, 1, a.ebitdaMargin, PCT); R++;
  const rowTax = R; setLabel(R, 0, "Tax rate"); setInput(R, 1, a.taxRate, PCT); R++;
  const rowDA = R; setLabel(R, 0, "D&A (% of revenue)"); setInput(R, 1, a.daPct, PCT); R++;
  const rowCapex = R; setLabel(R, 0, "CapEx (% of revenue)"); setInput(R, 1, a.capexPct, PCT); R++;
  const rowNWC = R; setLabel(R, 0, "Change in NWC (% of revenue)"); setInput(R, 1, a.nwcPct, PCT); R++;
  const rowWacc = R; setLabel(R, 0, "WACC"); setInput(R, 1, a.wacc, PCT); R++;
  const rowTermGrowth = R; setLabel(R, 0, "Terminal growth rate"); setInput(R, 1, a.terminalGrowth, PCT); R++;
  const rowNetDebt = R; setLabel(R, 0, "Net debt ($)"); setInput(R, 1, a.netDebt, USD); R++;
  const rowShares = R; setLabel(R, 0, "Shares outstanding"); setInput(R, 1, a.sharesOutstanding, NUM0); R++;
  R++;

  // ── Projection table ──
  setLabel(R, 0, "", true);
  for (let yc = 1; yc <= 5; yc++) setCell(R, yc, { v: `Year ${yc}`, t: "s", s: { font: { bold: true, sz: 10, color: { rgb: CLR.black } }, fill: { fgColor: { rgb: CLR.teal } }, alignment: { horizontal: "right", vertical: "center" } } });
  R++;

  const rowGrowth = R; setLabel(R, 0, "Revenue growth");
  for (let yc = 1; yc <= 5; yc++) setInput(R, yc, a.revenueGrowth[yc - 1], PCT);
  R++;

  const rowRevenue = R; setLabel(R, 0, "Revenue", true);
  for (let yc = 1; yc <= 5; yc++) {
    const f = yc === 1
      ? `${abs(1, rowBaseRevenue)}*(1+${rel(yc, rowGrowth)})`
      : `${rel(yc - 1, rowRevenue)}*(1+${rel(yc, rowGrowth)})`;
    setFormula(R, yc, f, USD, true);
  }
  R++;

  const rowEbitda = R; setLabel(R, 0, "EBITDA", true);
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowRevenue)}*${abs(1, rowEbitdaMargin)}`, USD, true);
  R++;

  const rowDAline = R; setLabel(R, 0, "D&A");
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowRevenue)}*${abs(1, rowDA)}`, USD);
  R++;

  const rowEbit = R; setLabel(R, 0, "EBIT");
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowEbitda)}-${rel(yc, rowDAline)}`, USD);
  R++;

  const rowTaxes = R; setLabel(R, 0, "Taxes");
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `MAX(${rel(yc, rowEbit)},0)*${abs(1, rowTax)}`, USD);
  R++;

  const rowNopat = R; setLabel(R, 0, "NOPAT");
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowEbit)}-${rel(yc, rowTaxes)}`, USD);
  R++;

  const rowCapexLine = R; setLabel(R, 0, "CapEx");
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowRevenue)}*${abs(1, rowCapex)}`, USD);
  R++;

  const rowNwcLine = R; setLabel(R, 0, "Change in NWC");
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowRevenue)}*${abs(1, rowNWC)}`, USD);
  R++;

  const rowUfcf = R; setLabel(R, 0, "Unlevered free cash flow", true);
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowNopat)}+${rel(yc, rowDAline)}-${rel(yc, rowCapexLine)}-${rel(yc, rowNwcLine)}`, USD, true, true);
  R++;

  const rowDF = R; setLabel(R, 0, "Discount factor");
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `1/(1+${abs(1, rowWacc)})^${yc}`, DF);
  R++;

  const rowPV = R; setLabel(R, 0, "Present value of FCF", true);
  for (let yc = 1; yc <= 5; yc++) setFormula(R, yc, `${rel(yc, rowUfcf)}*${rel(yc, rowDF)}`, USD, true);
  R++;
  R++;

  // ── Valuation ──
  const rowSumPV = R; setLabel(R, 0, "Sum of PV (Years 1–5)"); setFormula(R, 1, `SUM(${rel(1, rowPV)}:${rel(5, rowPV)})`, USD); R++;
  const rowTV = R; setLabel(R, 0, "Terminal value"); setFormula(R, 1, `${rel(5, rowUfcf)}*(1+${abs(1, rowTermGrowth)})/(${abs(1, rowWacc)}-${abs(1, rowTermGrowth)})`, USD); R++;
  const rowPVTV = R; setLabel(R, 0, "PV of terminal value"); setFormula(R, 1, `${rel(1, rowTV)}*${rel(5, rowDF)}`, USD); R++;
  const rowEV = R; setLabel(R, 0, "Enterprise value", true); setFormula(R, 1, `${rel(1, rowSumPV)}+${rel(1, rowPVTV)}`, USD, true, true); R++;
  const rowLessDebt = R; setLabel(R, 0, "Less: net debt"); setFormula(R, 1, `-${abs(1, rowNetDebt)}`, USD); R++;
  const rowEQ = R; setLabel(R, 0, "Equity value", true); setFormula(R, 1, `${rel(1, rowEV)}+${rel(1, rowLessDebt)}`, USD, true, true); R++;
  const rowSharesRef = R; setLabel(R, 0, "Shares outstanding"); setFormula(R, 1, `${abs(1, rowShares)}`, NUM0); R++;
  setLabel(R, 0, "Implied share price", true); setFormula(R, 1, `${rel(1, rowEQ)}/${rel(1, rowSharesRef)}`, USD2, true, true); R++;
  R++;

  // ── Sensitivity (computed snapshot — a native 2-variable data table isn't portable via this writer) ──
  setLabel(R, 0, "Sensitivity — implied share price (snapshot)", true); R++;
  setLabel(R, 0, "WACC ↓ / Terminal growth →");
  sensitivity.growthSteps.forEach((g, gi) => {
    setCell(R, gi + 1, { v: g, t: "n", z: PCT, s: { font: { bold: true, sz: 10, color: { rgb: CLR.black } }, fill: { fgColor: { rgb: CLR.teal } }, alignment: { horizontal: "right", vertical: "center" } } });
  });
  R++;
  sensitivity.grid.forEach((row, wi) => {
    setLabel(R, 0, `${(sensitivity.waccSteps[wi] * 100).toFixed(1)}%`, true);
    row.forEach((cell, gi) => {
      if (cell.error) {
        setCell(R, gi + 1, { v: "n/a", t: "s", s: { font: { sz: 10, color: { rgb: "B00020" } }, alignment: { horizontal: "right", vertical: "center" } } });
      } else {
        setCell(R, gi + 1, { v: cell.price, t: "n", z: USD2, s: { font: { sz: 10, color: { rgb: CLR.dark } }, alignment: { horizontal: "right", vertical: "center" } } });
      }
    });
    R++;
  });
  R++;

  setLabel(R, 0, "Assumptions were generated by AI and should be verified; every calculation above is a live Excel formula.");
  R++;

  ws["!ref"] = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: R, c: 5 } });
  ws["!cols"] = colMax.map((w) => ({ wch: w }));
  XLSXStyle.utils.book_append_sheet(wb, ws, "DCF Model");
  XLSXStyle.writeFile(wb, `FinModel_DCF_${a.companyName.replace(/\s+/g, "_")}.xlsx`);
}

// ── Main Component ──────────────────────────────────────────────

export default function Home() {
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [simpleMode, setSimpleMode] = useState(false);
  const [copiedIdx, setCopiedIdx] = useState(null);
  const [modal, setModal] = useState(null);
  const [mobileTab, setMobileTab] = useState("chat");
  const [templateSel, setTemplateSel] = useState("");
  const [dcfFields, setDcfFields] = useState(null);
  const [dcfMsgIdx, setDcfMsgIdx] = useState(null);

  const bottomRef = useRef(null);
  const msgRefs = useRef({});
  const textareaRef = useRef(null);
  const dcfPanelRef = useRef(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, loading]);

  const modelHistory = messages.reduce((acc, m, i) => {
    if (m.role === "user" && messages[i + 1]?.role === "assistant") {
      acc.push({ title: m.content.split("\n")[0].slice(0, 55), idx: i + 1 });
    }
    return acc;
  }, []);

  // Latest assistant message that contains a model — either a deterministic
  // DCF (kind: "dcf") or a legacy markdown-table model (any other model type).
  let latestModel = null;
  for (let i = messages.length - 1; i >= 0; i--) {
    const m = messages[i];
    if (m.role !== "assistant") continue;
    if (m.kind === "dcf") {
      latestModel = { type: "dcf", idx: i, assumptions: m.assumptions, modelType: "DCF Model" };
      break;
    }
    const tables = parseMarkdownTables(m.content);
    if (tables.length) {
      latestModel = { type: "table", idx: i, content: m.content, tables, modelType: detectModelType(m.content) };
      break;
    }
  }

  useEffect(() => {
    if (latestModel?.type === "dcf" && latestModel.idx !== dcfMsgIdx) {
      setDcfFields(fieldsFromAssumptions(latestModel.assumptions));
      setDcfMsgIdx(latestModel.idx);
    }
  }, [latestModel?.idx, latestModel?.type]);

  const liveAssumptions = dcfFields ? assumptionsFromFields(dcfFields) : null;
  const liveComputed = liveAssumptions ? computeDCF(liveAssumptions) : null;
  const liveSensitivity = liveAssumptions ? computeSensitivity(liveAssumptions) : null;

  const updateDcfField = (key, value) => {
    setDcfFields((f) => ({ ...f, [key]: value }));
  };
  const updateDcfGrowth = (yearIdx, value) => {
    setDcfFields((f) => {
      const growth = [...f.growth];
      growth[yearIdx] = value;
      return { ...f, growth };
    });
  };

  const sendMessage = async (text) => {
    const userMsg = text !== undefined ? text : input.trim();
    if (!userMsg || loading) return;
    setInput("");
    setError(null);
    if (textareaRef.current) textareaRef.current.style.height = "auto";

    const next = [...messages, { role: "user", content: userMsg }];
    setMessages(next);
    setLoading(true);
    setMobileTab("chat");

    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: next, simpleMode }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Request failed");
      if (data.kind === "dcf") {
        setMessages([...next, {
          role: "assistant",
          kind: "dcf",
          assumptions: data.assumptions,
          content: data.interpretation,
        }]);
      } else {
        setMessages([...next, { role: "assistant", content: data.reply }]);
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(); }
  };

  const copyMsg = async (content, i) => {
    await navigator.clipboard.writeText(content);
    setCopiedIdx(i);
    setTimeout(() => setCopiedIdx(null), 2000);
  };

  const openModal = (content) => {
    const tables = parseMarkdownTables(content);
    if (!tables.length) return;
    setModal({ table: JSON.parse(JSON.stringify(tables[0])) });
  };

  const regenerateWithEdits = () => {
    if (!modal) return;
    const { headers, rows } = modal.table;
    const md = [
      "| " + headers.join(" | ") + " |",
      "| " + headers.map(() => "---").join(" | ") + " |",
      ...rows.map((r) => "| " + r.join(" | ") + " |"),
    ].join("\n");
    setModal(null);
    sendMessage(
      `Please regenerate the financial model using these updated assumptions:\n\n${md}\n\nRecalculate all outputs based on the new inputs and keep the same structure and format.`
    );
  };

  const isEmpty = messages.length === 0;
  const suggestions = simpleMode ? SIMPLE_SUGGESTIONS : PRO_SUGGESTIONS;

  return (
    <>
      {/* Sidebar */}
      <aside className={`sidebar ${sidebarOpen ? "open" : ""}`}>
        <div className="sidebar-head">
          <span className="sidebar-title">Model history</span>
          <button className="sidebar-close" onClick={() => setSidebarOpen(false)}>×</button>
        </div>
        <div className="sidebar-list">
          {modelHistory.length === 0 ? (
            <p className="sidebar-empty">No models yet. Ask a question to get started.</p>
          ) : (
            modelHistory.map((item, n) => (
              <button
                key={n}
                className="sidebar-item"
                onClick={() => {
                  setSidebarOpen(false);
                  setMobileTab("chat");
                  msgRefs.current[item.idx]?.scrollIntoView({ behavior: "smooth", block: "start" });
                }}
              >
                <span className="sidebar-num">{String(n + 1).padStart(2, "0")}</span>
                <span className="sidebar-label">{item.title}</span>
              </button>
            ))
          )}
        </div>
      </aside>
      {sidebarOpen && <div className="sidebar-backdrop" onClick={() => setSidebarOpen(false)} />}

      {/* Main app */}
      <div className="app-shell">
        {/* Top bar */}
        <div className="topbar">
          <span className="wordmark">FinModel</span>
          <div className="topbar-right">
            <div className="mode-toggle">
              <button
                className={`mode-btn ${!simpleMode ? "active" : ""}`}
                onClick={() => setSimpleMode(false)}
              >
                Pro
              </button>
              <button
                className={`mode-btn ${simpleMode ? "active" : ""}`}
                onClick={() => setSimpleMode(true)}
              >
                Simple
              </button>
            </div>
            <button className="history-btn" onClick={() => setSidebarOpen((o) => !o)}>
              <svg viewBox="0 0 24 24" fill="currentColor"><path d="M3 18h18v-2H3v2zm0-5h18v-2H3v2zm0-7v2h18V6H3z" /></svg>
              History
            </button>
          </div>
        </div>

        {/* Mobile tab bar */}
        <div className="mobile-tabbar">
          <button className={`mobile-tab ${mobileTab === "chat" ? "active" : ""}`} onClick={() => setMobileTab("chat")}>
            Chat
          </button>
          <button className={`mobile-tab ${mobileTab === "model" ? "active" : ""}`} onClick={() => setMobileTab("model")}>
            Model
          </button>
        </div>

        {/* Workspace */}
        <div className="workspace">
          {/* Left pane — chat */}
          <section className={`pane pane-chat ${mobileTab === "chat" ? "mobile-active" : ""}`}>
            <div className="chat-scroll">
              {isEmpty ? (
                <div className="chat-empty">
                  <div className="chat-empty-title">What do you want to model?</div>
                  <div className="chat-empty-examples">
                    {suggestions.map((s, i) => (
                      <button key={i} className="example-link" onClick={() => sendMessage(s)}>
                        {s}
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                <>
                  {messages.length > 2 && (
                    <button className="clear-btn" onClick={() => setMessages([])}>Clear chat</button>
                  )}
                  {messages.map((m, i) => {
                    const displayText = m.role === "assistant"
                      ? (stripTables(m.content) || "Model updated — see the Model panel.")
                      : m.content;
                    return (
                      <div
                        key={i}
                        className={`message ${m.role}`}
                        ref={(el) => { if (el) msgRefs.current[i] = el; }}
                      >
                        <div className="message-label">{m.role === "user" ? "You" : "Analyst"}</div>
                        <div
                          className="message-content"
                          dangerouslySetInnerHTML={{ __html: formatMessage(displayText) }}
                        />
                        {m.kind === "dcf" && (
                          <p className="chat-model-note">The full interactive model is in the Model panel →</p>
                        )}
                        {m.role === "assistant" && (
                          <div className="msg-foot">
                            <button className="copy-link" onClick={() => copyMsg(m.content, i)}>
                              {copiedIdx === i ? "Copied" : "Copy"}
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                  {loading && (
                    <div className="message assistant">
                      <div className="message-label">Analyst</div>
                      <div className="typing-indicator">
                        <span /><span /><span />
                      </div>
                    </div>
                  )}
                </>
              )}
              <div ref={bottomRef} />
            </div>

            {/* Composer */}
            <div className="composer">
              <div className="composer-toolbar">
                <select
                  className="template-select"
                  value={templateSel}
                  onChange={(e) => {
                    const t = TEMPLATES.find((tpl) => tpl.label === e.target.value);
                    if (t) { setInput(t.prompt); textareaRef.current?.focus(); }
                    setTemplateSel("");
                  }}
                >
                  <option value="" disabled>Model type…</option>
                  {TEMPLATES.map((t) => (
                    <option key={t.label} value={t.label}>{t.label}</option>
                  ))}
                </select>
              </div>
              <div className="composer-input">
                <textarea
                  ref={textareaRef}
                  value={input}
                  onChange={(e) => {
                    setInput(e.target.value);
                    e.target.style.height = "auto";
                    e.target.style.height = e.target.scrollHeight + "px";
                  }}
                  onKeyDown={handleKeyDown}
                  placeholder="Describe a company or scenario to model..."
                  rows={1}
                />
                <button
                  className="send-btn"
                  onClick={() => sendMessage()}
                  disabled={!input.trim() || loading}
                >
                  <svg viewBox="0 0 24 24"><path d="M2 21l21-9L2 3v7l15 2-15 2v7z" /></svg>
                </button>
              </div>
              {error && <div className="error-msg">{error}</div>}
              <div className="composer-hint">Shift+Enter for new line · Enter to send</div>
            </div>
          </section>

          {/* Right pane — model */}
          <section className={`pane pane-model ${mobileTab === "model" ? "mobile-active" : ""}`}>
            <div className="model-panel-head">
              <span className="model-panel-title">{latestModel ? latestModel.modelType : "Model"}</span>
              {latestModel?.type === "table" && (
                <div className="model-panel-actions">
                  <button className="panel-btn" onClick={() => copyMsg(latestModel.content, latestModel.idx)}>
                    {copiedIdx === latestModel.idx ? "Copied" : "Copy"}
                  </button>
                  <button className="panel-btn" onClick={() => exportToExcel(latestModel.content)}>Excel</button>
                  <button className="panel-btn" onClick={() => exportToPDF(formatMessage(latestModel.content))}>PDF</button>
                  <button className="panel-btn panel-btn--primary" onClick={() => openModal(latestModel.content)}>
                    Edit assumptions
                  </button>
                </div>
              )}
              {latestModel?.type === "dcf" && liveComputed && (
                <div className="model-panel-actions">
                  <button
                    className="panel-btn"
                    onClick={() => copyMsg(dcfSummaryText(liveAssumptions, liveComputed), latestModel.idx)}
                  >
                    {copiedIdx === latestModel.idx ? "Copied" : "Copy"}
                  </button>
                  <button className="panel-btn" onClick={() => exportDCFToExcel(liveAssumptions, liveSensitivity)}>Excel</button>
                  <button className="panel-btn" onClick={() => exportToPDF(dcfPanelRef.current?.innerHTML || "")}>PDF</button>
                </div>
              )}
            </div>
            <div className="model-panel-body">
              {!latestModel && (
                <div className="model-empty">
                  <div className="model-empty-grid" />
                  <p className="model-empty-text">Your model will appear here.</p>
                </div>
              )}

              {latestModel?.type === "table" && (
                latestModel.tables.map((table, ti) => (
                  <div className="model-table-wrap" key={ti}>
                    <table className="model-table">
                      <thead>
                        <tr>
                          {table.headers.map((h, hi) => (
                            <th key={hi}>{h.replace(/\*\*/g, "")}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {table.rows.map((row, ri) => (
                          <tr key={ri} className={rowClass(row)}>
                            {row.map((cell, ci) => {
                              const { text, numeric } = displayCell(cell);
                              return (
                                <td key={ci} className={numeric ? "cell-num" : ""}>{text}</td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ))
              )}

              {latestModel?.type === "dcf" && dcfFields && liveComputed && (
                <div className="dcf-panel" ref={dcfPanelRef}>
                  <div className="dcf-block">
                    <div className="dcf-section-title">Assumptions</div>
                    <div className="dcf-assumptions-grid">
                      <label className="dcf-field dcf-field--wide">
                        <span>Company name</span>
                        <input
                          className="dcf-input dcf-input-text"
                          value={dcfFields.companyName}
                          onChange={(e) => updateDcfField("companyName", e.target.value)}
                        />
                      </label>
                      {[
                        ["baseRevenue", "Base revenue ($)"],
                        ["ebitdaMargin", "EBITDA margin (%)"],
                        ["taxRate", "Tax rate (%)"],
                        ["daPct", "D&A (% of revenue)"],
                        ["capexPct", "CapEx (% of revenue)"],
                        ["nwcPct", "Δ NWC (% of revenue)"],
                        ["wacc", "WACC (%)"],
                        ["terminalGrowth", "Terminal growth (%)"],
                        ["netDebt", "Net debt ($)"],
                        ["sharesOutstanding", "Shares outstanding"],
                      ].map(([key, label]) => (
                        <label className="dcf-field" key={key}>
                          <span>{label}</span>
                          <input
                            className="dcf-input"
                            value={dcfFields[key]}
                            onChange={(e) => updateDcfField(key, e.target.value)}
                          />
                        </label>
                      ))}
                    </div>
                  </div>

                  <div className="dcf-block">
                    <div className="dcf-section-title">5-year projection</div>
                    <div className="model-table-wrap">
                      <table className="model-table">
                        <thead>
                          <tr>
                            <th></th>
                            {liveComputed.years.map((y) => <th key={y.year}>Year {y.year}</th>)}
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td>Revenue growth</td>
                            {dcfFields.growth.map((g, gi) => (
                              <td key={gi} className="cell-num dcf-cell-editable">
                                <input
                                  className="dcf-input dcf-input--cell"
                                  value={g}
                                  onChange={(e) => updateDcfGrowth(gi, e.target.value)}
                                />
                              </td>
                            ))}
                          </tr>
                          <tr className="row-section">
                            <td>Revenue</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.revenue)}</td>)}
                          </tr>
                          <tr className="row-section">
                            <td>EBITDA</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.ebitda)}</td>)}
                          </tr>
                          <tr>
                            <td>D&A</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.da)}</td>)}
                          </tr>
                          <tr>
                            <td>EBIT</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.ebit)}</td>)}
                          </tr>
                          <tr>
                            <td>Taxes</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.taxes)}</td>)}
                          </tr>
                          <tr>
                            <td>NOPAT</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.nopat)}</td>)}
                          </tr>
                          <tr>
                            <td>CapEx</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.capex)}</td>)}
                          </tr>
                          <tr>
                            <td>Change in NWC</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.nwcChange)}</td>)}
                          </tr>
                          <tr className="row-total">
                            <td>Unlevered free cash flow</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.ufcf)}</td>)}
                          </tr>
                          <tr>
                            <td>Discount factor</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{y.discountFactor.toFixed(3)}</td>)}
                          </tr>
                          <tr className="row-total">
                            <td>Present value of FCF</td>
                            {liveComputed.years.map((y) => <td key={y.year} className="cell-num">{formatMoney(y.pv)}</td>)}
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div className="dcf-block">
                    <div className="dcf-section-title">Valuation</div>
                    <div className="model-table-wrap">
                      <table className="model-table">
                        <tbody>
                          <tr>
                            <td>Sum of PV (Years 1–5)</td>
                            <td className="cell-num">{formatMoney(liveComputed.sumPV)}</td>
                          </tr>
                          <tr>
                            <td>Terminal value</td>
                            <td className="cell-num">
                              {liveComputed.terminalValid
                                ? formatMoney(liveComputed.terminalValue)
                                : <span className="dcf-error">WACC ≤ terminal growth</span>}
                            </td>
                          </tr>
                          <tr>
                            <td>PV of terminal value</td>
                            <td className="cell-num">{liveComputed.terminalValid ? formatMoney(liveComputed.pvTerminalValue) : "—"}</td>
                          </tr>
                          <tr className="row-total">
                            <td>Enterprise value</td>
                            <td className="cell-num">{liveComputed.terminalValid ? formatMoney(liveComputed.enterpriseValue) : "—"}</td>
                          </tr>
                          <tr>
                            <td>Less: net debt</td>
                            <td className="cell-num">{formatMoney(-liveAssumptions.netDebt)}</td>
                          </tr>
                          <tr className="row-total">
                            <td>Equity value</td>
                            <td className="cell-num">{liveComputed.terminalValid ? formatMoney(liveComputed.equityValue) : "—"}</td>
                          </tr>
                          <tr>
                            <td>Shares outstanding</td>
                            <td className="cell-num">{liveAssumptions.sharesOutstanding.toLocaleString("en-US")}</td>
                          </tr>
                          <tr className="row-total">
                            <td>Implied share price</td>
                            <td className="cell-num dcf-price-highlight">
                              {liveComputed.terminalValid
                                ? formatPrice(liveComputed.impliedSharePrice)
                                : <span className="dcf-error">n/a</span>}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div className="dcf-block">
                    <div className="dcf-section-title">Sensitivity — implied share price</div>
                    <div className="model-table-wrap">
                      <table className="model-table">
                        <thead>
                          <tr>
                            <th>WACC ↓ / growth →</th>
                            {liveSensitivity.growthSteps.map((g, gi) => <th key={gi}>{formatPct(g)}</th>)}
                          </tr>
                        </thead>
                        <tbody>
                          {liveSensitivity.grid.map((row, wi) => (
                            <tr key={wi}>
                              <td>{formatPct(liveSensitivity.waccSteps[wi])}</td>
                              {row.map((cell, gi) => (
                                <td key={gi} className="cell-num">
                                  {cell.error ? <span className="dcf-error">n/a</span> : formatPrice(cell.price)}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <p className="dcf-note">
                    Assumptions were generated by AI and should be checked — the math above is calculated deterministically, not by the AI.
                  </p>
                </div>
              )}
            </div>
          </section>
        </div>
      </div>

      {/* Edit Assumptions Modal */}
      {modal && (
        <div className="modal-overlay" onClick={() => setModal(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <span>Edit assumptions</span>
              <button className="modal-close" onClick={() => setModal(null)}>×</button>
            </div>
            <div className="modal-body">
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      {modal.table.headers.map((h, hi) => <th key={hi}>{h}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {modal.table.rows.map((row, ri) => (
                      <tr key={ri}>
                        {row.map((cell, ci) => (
                          <td key={ci}>
                            <input
                              className="cell-input"
                              value={cell}
                              onChange={(e) => {
                                const t = JSON.parse(JSON.stringify(modal.table));
                                t.rows[ri][ci] = e.target.value;
                                setModal({ table: t });
                              }}
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="modal-hint">Edit values above, then click Regenerate to recalculate the model with new inputs.</p>
            </div>
            <div className="modal-foot">
              <button className="modal-btn" onClick={() => setModal(null)}>Cancel</button>
              <button className="modal-btn modal-btn--go" onClick={regenerateWithEdits}>
                Regenerate ↗
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
