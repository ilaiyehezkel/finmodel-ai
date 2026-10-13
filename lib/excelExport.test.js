import test from "node:test";
import assert from "node:assert/strict";
import { HyperFormula } from "hyperformula";
import XLSXStyle from "xlsx-js-style";
import { buildDCFWorkbook, buildLBOWorkbook } from "./excelExport.js";
import { computeDCF, validateAssumptions as validateDCFAssumptions, computeSensitivity as computeDCFSensitivity } from "./dcf.js";
import { computeLBO, validateAssumptions as validateLBOAssumptions, computeSensitivity as computeLBOSensitivity } from "./lbo.js";

const close = (actual, expected, tolerance = 0.01, label = "") => {
  assert.ok(
    Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance,
    `${label ? label + ": " : ""}expected ${actual} to be within ${tolerance} of ${expected}`
  );
};

// Converts a SheetJS worksheet (bounded by its own declared !ref) into the
// 2D array HyperFormula expects — formula cells become "=..." strings,
// everything else passes through as a raw value. Bounding strictly by !ref
// (rather than scanning every key in `ws`) deliberately mirrors how a strict
// reader would see the sheet, so a dimension bug (like the one that caused
// the original #NUM! bug) would reproduce here too.
function sheetToArray(ws) {
  const range = XLSXStyle.utils.decode_range(ws["!ref"]);
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSXStyle.utils.encode_cell({ r, c })];
      if (!cell) { row.push(null); continue; }
      row.push(cell.f ? `=${cell.f}` : (cell.v ?? null));
    }
    rows.push(row);
  }
  return rows;
}

function evaluate(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const hf = HyperFormula.buildFromArray(sheetToArray(ws), { licenseKey: "gpl-v3" });
  return (row, col = 1) => hf.getCellValue({ sheet: 0, row, col });
}

// ── LBO ──────────────────────────────────────────────────────────

function checkLBOWorkbook(rawAssumptions, label) {
  const assumptions = validateLBOAssumptions(rawAssumptions);
  const ground = computeLBO(assumptions);
  const sensitivity = computeLBOSensitivity(assumptions);
  const { wb, refs } = buildLBOWorkbook(assumptions, sensitivity);
  const cell = evaluate(wb, "LBO Model");

  assert.ok(ground.valid, `${label}: fixture must be a valid deal (sponsor equity > 0)`);

  close(cell(refs.rows.sponsorEquity), ground.sourcesUses.sponsorEquity, 1, `${label} sponsor equity`);
  close(cell(refs.rows.exitEquity), ground.exit.exitEquityValue, 1, `${label} exit equity value`);
  close(cell(refs.rows.moic), ground.returns.moic, 1e-4, `${label} MOIC`);
  close(cell(refs.rows.irr), ground.returns.irr, 1e-4, `${label} IRR`);

  // The bug this test guards against: Excel's IRR() must see Entry through
  // Year 5 (6 cells). If Year 5 were dropped, the exit proceeds (which only
  // land in the column matching the holding period) would be invisible to
  // IRR, which would then either throw #NUM! or silently misprice the deal.
  for (let yc = 0; yc <= 5; yc++) {
    const v = cell(refs.rows.cashFlow, yc + 1);
    assert.ok(Number.isFinite(v), `${label}: cash flow column for ${yc === 0 ? "Entry" : "Year " + yc} must be a real number, got ${v}`);
  }
  close(cell(refs.rows.cashFlow, assumptions.holdingPeriod + 1), ground.exit.sponsorProceeds, 1);
}

test("LBO workbook formulas match lib/lbo.js — base case (default assumptions)", () => {
  checkLBOWorkbook({}, "base case");
});

test("LBO workbook formulas match lib/lbo.js — higher interest rates", () => {
  // Confirms the formulas actually re-derive from the input cells rather
  // than baking in stale numbers: a materially different capital structure
  // must produce materially different (but still formula-correct) results.
  const base = validateLBOAssumptions({});
  const highRate = {
    ...base,
    seniorDebt: { ...base.seniorDebt, rate: 0.14 },
    subDebt: { ...base.subDebt, rate: 0.18 },
  };
  checkLBOWorkbook(highRate, "higher interest rates");

  const baseGround = computeLBO(base);
  const highGround = computeLBO(highRate);
  assert.notEqual(Math.round(baseGround.returns.irr * 1e6), Math.round(highGround.returns.irr * 1e6));
});

test("LBO sensitivity grid in the exported workbook is internally consistent", () => {
  const assumptions = validateLBOAssumptions({});
  const sensitivity = computeLBOSensitivity(assumptions);
  const { grid, entrySteps, exitSteps } = sensitivity;
  assert.equal(entrySteps.length, 5);
  assert.equal(exitSteps.length, 5);
  assert.equal(grid.length, 5);
  grid.forEach((row) => assert.equal(row.length, 5));
});

// ── DCF ──────────────────────────────────────────────────────────

function checkDCFWorkbook(rawAssumptions, label) {
  const assumptions = validateDCFAssumptions(rawAssumptions);
  const ground = computeDCF(assumptions);
  const sensitivity = computeDCFSensitivity(assumptions);
  const { wb, refs } = buildDCFWorkbook(assumptions, sensitivity);
  const cell = evaluate(wb, "DCF Model");

  assert.ok(ground.terminalValid, `${label}: fixture must have WACC > terminal growth`);

  // Every one of the 5 projection years, specifically free cash flow.
  for (let yc = 1; yc <= 5; yc++) {
    close(cell(refs.rows.ufcf, yc), ground.years[yc - 1].ufcf, 1, `${label} year ${yc} UFCF`);
    close(cell(refs.rows.revenue, yc), ground.years[yc - 1].revenue, 1);
    close(cell(refs.rows.pv, yc), ground.years[yc - 1].pv, 1);
  }

  // Terminal value must be derived from the final (year 5) UFCF.
  close(cell(refs.rows.terminalValue), ground.terminalValue, 1);
  close(cell(refs.rows.sumPV), ground.sumPV, 1);
  close(cell(refs.rows.enterpriseValue), ground.enterpriseValue, 1);
  close(cell(refs.rows.equityValue), ground.equityValue, 1);
  close(cell(refs.rows.impliedSharePrice), ground.impliedSharePrice, 0.01);
}

test("DCF workbook formulas match lib/dcf.js — base case (default assumptions)", () => {
  checkDCFWorkbook({}, "base case");
});

test("DCF workbook formulas match lib/dcf.js — higher WACC", () => {
  const base = validateDCFAssumptions({});
  const highWacc = { ...base, wacc: 0.22 };
  checkDCFWorkbook(highWacc, "higher WACC");

  const baseGround = computeDCF(base);
  const highGround = computeDCF(highWacc);
  // A higher discount rate must produce a lower enterprise value, and the
  // formulas (not a stale cached number) must be what produces that.
  assert.ok(highGround.enterpriseValue < baseGround.enterpriseValue);

  const { wb: wbBase, refs: refsBase } = buildDCFWorkbook(base, computeDCFSensitivity(base));
  const { wb: wbHigh, refs: refsHigh } = buildDCFWorkbook(highWacc, computeDCFSensitivity(highWacc));
  const evBase = evaluate(wbBase, "DCF Model")(refsBase.rows.enterpriseValue);
  const evHigh = evaluate(wbHigh, "DCF Model")(refsHigh.rows.enterpriseValue);
  assert.ok(evHigh < evBase);
});

test("DCF enterprise value equals the sum of all 5 PVs plus the PV of terminal value", () => {
  const assumptions = validateDCFAssumptions({});
  const sensitivity = computeDCFSensitivity(assumptions);
  const { wb, refs } = buildDCFWorkbook(assumptions, sensitivity);
  const cell = evaluate(wb, "DCF Model");

  let sumPV = 0;
  for (let yc = 1; yc <= 5; yc++) sumPV += cell(refs.rows.pv, yc);
  close(sumPV, cell(refs.rows.sumPV), 0.01);
  close(sumPV + cell(refs.rows.pvTerminalValue), cell(refs.rows.enterpriseValue), 0.01);
});
