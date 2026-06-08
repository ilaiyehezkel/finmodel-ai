"use client";

import { useState, useRef, useEffect } from "react";
import { formatMessage } from "../lib/format";
import * as XLSXStyle from "xlsx-js-style";

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

const SUGGESTIONS = [
  "Build a DCF model for a SaaS company with $2M ARR growing 80% YoY",
  "Model a cap table for a startup raising a $3M seed at $12M pre-money",
  "Calculate LTV/CAC ratio given $50 monthly revenue, 85% retention, $200 CAC",
  "3-statement model for a business with $500K revenue and 60% gross margin",
  "LBO model for acquiring a company at 8x EBITDA with 60% debt financing",
  "Revenue forecast for a marketplace with 1,000 users and 5% monthly growth",
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

// ── Excel export colours & constants ───────────────────────────

const CLR = {
  green:     "00D084",
  greenSoft: "D6F5EA",
  greenDark: "004D2E",
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

function exportToExcel(text, msgIdx) {
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
        fill: { fgColor: { rgb: CLR.green } },
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
          fill: { fgColor: { rgb: CLR.green } },
          font: { bold: true, sz: 10, color: { rgb: CLR.black } },
          alignment: { horizontal: c === 0 ? "left" : "center", vertical: "center" },
          border: { bottom: { style: "medium", color: { rgb: "008050" } } },
        },
      });
    }
    R++;

    // Data rows
    table.rows.forEach((row, rIdx) => {
      const isSummary = isSummaryRow(row);
      const bg = isSummary ? CLR.greenSoft : rIdx % 2 === 0 ? CLR.white : CLR.rowAlt;

      for (let c = 0; c < cols; c++) {
        const cellVal = row[c] ?? "";
        const isFirst = c === 0;
        const parsed = parseCell(cellVal);

        const style = {
          fill: { fgColor: { rgb: bg } },
          font: {
            bold: isSummary || isFirst,
            sz: 10,
            color: { rgb: isSummary ? CLR.greenDark : CLR.dark },
          },
          alignment: {
            horizontal: isFirst ? "left" : parsed.t === "n" ? "right" : "center",
            vertical: "center",
          },
        };

        if (isSummary) {
          style.border = {
            top: { style: "thin", color: { rgb: "00A068" } },
            bottom: { style: "thin", color: { rgb: "00A068" } },
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
body{font-family:monospace;font-size:13px;line-height:1.7;color:#111;padding:40px 56px;max-width:860px;margin:0 auto}
h1{font-size:22px;margin:16px 0 6px}h2{font-size:18px;margin:14px 0 6px}h3{font-size:15px;color:#006b3c;margin:12px 0 6px}
p{margin:0 0 10px}
table{width:100%;border-collapse:collapse;margin:14px 0;font-size:12px;page-break-inside:avoid}
th{background:#f2f2f2;padding:8px 12px;border:1px solid #ccc;text-align:left;font-weight:600;font-size:11px;letter-spacing:.04em}
td{padding:7px 12px;border:1px solid #ddd}tr:nth-child(even) td{background:#fafafa}
code{background:#f4f4f4;padding:2px 5px;border-radius:3px;font-size:12px}
pre{background:#f4f4f4;padding:14px;border-radius:6px;margin:10px 0}pre code{background:none;padding:0}
strong{color:#000}ul{padding-left:18px;margin:8px 0}li{margin:3px 0}
.ph{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #006b3c;padding-bottom:12px;margin-bottom:24px}
.ph-title{color:#006b3c;font-size:18px;font-weight:600;margin:0}.ph-sub{color:#666;font-size:11px}
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

  const bottomRef = useRef(null);
  const msgRefs = useRef({});
  const textareaRef = useRef(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, loading]);

  const modelHistory = messages.reduce((acc, m, i) => {
    if (m.role === "user" && messages[i + 1]?.role === "assistant") {
      acc.push({ title: m.content.split("\n")[0].slice(0, 55), idx: i + 1 });
    }
    return acc;
  }, []);

  const sendMessage = async (text) => {
    const userMsg = text !== undefined ? text : input.trim();
    if (!userMsg || loading) return;
    setInput("");
    setError(null);
    if (textareaRef.current) textareaRef.current.style.height = "auto";

    const next = [...messages, { role: "user", content: userMsg }];
    setMessages(next);
    setLoading(true);

    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: next, simpleMode }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Request failed");
      setMessages([...next, { role: "assistant", content: data.reply }]);
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

  return (
    <>
      {/* Sidebar */}
      <aside className={`sidebar ${sidebarOpen ? "open" : ""}`}>
        <div className="sidebar-head">
          <span className="sidebar-title">Session Models</span>
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
      <div className="app">
        {/* Header */}
        <div className="header">
          <button
            className="sidebar-toggle"
            onClick={() => setSidebarOpen((o) => !o)}
            title="Toggle model history"
          >
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor">
              <path d="M3 18h18v-2H3v2zm0-5h18v-2H3v2zm0-7v2h18V6H3z" />
            </svg>
          </button>
          <div className="logo">$</div>
          <div className="header-text">
            <h1>FinModel AI</h1>
            <p>Financial Modeling Assistant</p>
          </div>
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
          <div className="header-badge">● Live</div>
        </div>

        {/* Template bar */}
        <div className="template-bar">
          {TEMPLATES.map((t) => (
            <button
              key={t.label}
              className="template-btn"
              onClick={() => { setInput(t.prompt); textareaRef.current?.focus(); }}
            >
              {t.label}
            </button>
          ))}
        </div>

        {/* Chat area */}
        <div className="chat-area">
          {isEmpty ? (
            <div className="empty-state">
              <div className="empty-icon">📊</div>
              <div className="empty-title">Build any financial model</div>
              <div className="empty-sub">DCF · LBO · Cap Tables · Unit Economics · Forecasts</div>
              <div className="suggestions-grid">
                {SUGGESTIONS.map((s, i) => (
                  <button key={i} className="suggestion-btn" onClick={() => sendMessage(s)}>
                    {s}
                  </button>
                ))}
              </div>
            </div>
          ) : (
            <>
              {messages.length > 2 && (
                <button className="clear-btn" onClick={() => setMessages([])}>clear chat</button>
              )}
              {messages.map((m, i) => {
                const hasTables = m.role === "assistant" && parseMarkdownTables(m.content).length > 0;
                return (
                  <div
                    key={i}
                    className={`message ${m.role}`}
                    ref={(el) => { if (el) msgRefs.current[i] = el; }}
                  >
                    <div className="message-label">{m.role === "user" ? "YOU" : "ANALYST"}</div>
                    <div
                      className="message-content"
                      dangerouslySetInnerHTML={{ __html: formatMessage(m.content) }}
                    />
                    {m.role === "assistant" && (
                      <div className="msg-actions">
                        <button className="act-btn" onClick={() => copyMsg(m.content, i)}>
                          {copiedIdx === i ? "Copied!" : "Copy"}
                        </button>
                        {hasTables && (
                          <button className="act-btn" onClick={() => exportToExcel(m.content, i)}>
                            ↓ Excel
                          </button>
                        )}
                        <button
                          className="act-btn"
                          onClick={() => {
                            const el = msgRefs.current[i]?.querySelector(".message-content");
                            exportToPDF(el?.innerHTML || "");
                          }}
                        >
                          ↓ PDF
                        </button>
                        {hasTables && (
                          <button className="act-btn act-btn--hi" onClick={() => openModal(m.content)}>
                            Edit Assumptions
                          </button>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
              {loading && (
                <div className="message assistant">
                  <div className="message-label">ANALYST</div>
                  <div className="typing-indicator">
                    <span /><span /><span />
                  </div>
                </div>
              )}
            </>
          )}
          <div ref={bottomRef} />
        </div>

        {/* Input area */}
        <div className="input-area">
          <div className="input-wrap">
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
          <div className="input-hint">
            SHIFT+ENTER for new line · ENTER to send
            {simpleMode && <span className="simple-badge"> · SIMPLE MODE ON</span>}
          </div>
        </div>
      </div>

      {/* Edit Assumptions Modal */}
      {modal && (
        <div className="modal-overlay" onClick={() => setModal(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <span>Edit Assumptions</span>
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
