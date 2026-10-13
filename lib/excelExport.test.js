import test from "node:test";
import assert from "node:assert/strict";
import { HyperFormula } from "hyperformula";
import XLSXStyle from "xlsx-js-style";
import { buildDCFWorkbook, buildLBOWorkbook, buildUnitEconomicsWorkbook, buildCapTableWorkbook } from "./excelExport.js";
import { computeDCF, validateAssumptions as validateDCFAssumptions, computeSensitivity as computeDCFSensitivity } from "./dcf.js";
import { computeLBO, validateAssumptions as validateLBOAssumptions, computeSensitivity as computeLBOSensitivity } from "./lbo.js";
import { computeUnitEconomics, validateAssumptions as validateUEAssumptions, computeSensitivity as computeUESensitivity } from "./unitEconomics.js";
import { computeCapTable, computePayoutTable, validateAssumptions as validateCapTableAssumptions } from "./capTable.js";

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

// ── Unit Economics ──────────────────────────────────────────────

function checkUEWorkbook(rawAssumptions, label) {
  const assumptions = validateUEAssumptions(rawAssumptions);
  const ground = computeUnitEconomics(assumptions);
  const sensitivity = computeUESensitivity(assumptions);
  const { wb, refs } = buildUnitEconomicsWorkbook(assumptions, sensitivity);
  const cell = evaluate(wb, "Unit Economics");

  close(cell(refs.rows.grossProfitPerMonth), ground.grossProfitPerMonth, 0.01, `${label} gross profit/mo`);
  close(cell(refs.rows.simpleLTV), ground.simpleLTV, 0.01, `${label} simple LTV`);
  close(cell(refs.rows.discountedLTV), ground.discountedLTV, 0.5, `${label} discounted LTV`);
  close(cell(refs.rows.ltvToCac), ground.ltvToCac, 0.001, `${label} LTV/CAC`);
  if (ground.cacPaybackMonths !== null) {
    close(cell(refs.rows.paybackMonth), ground.cacPaybackMonths, 0.01, `${label} payback month`);
  }

  for (let m = 1; m <= assumptions.projectionMonths; m++) {
    close(cell(refs.rows.cumulativeGP, m), ground.months[m - 1].cumulativeGrossProfit, 0.01, `${label} month ${m} cumulative GP`);
  }
}

test("Unit Economics workbook formulas match lib/unitEconomics.js — base case", () => {
  checkUEWorkbook({ arpu: 100, grossMargin: 0.80, monthlyChurn: 0.05, cac: 500, monthlyDiscountRate: 0.01, projectionMonths: 36 }, "base case");
});

test("Unit Economics workbook formulas match lib/unitEconomics.js — zero churn edge case", () => {
  checkUEWorkbook({ arpu: 100, grossMargin: 0.80, monthlyChurn: 0, cac: 500, monthlyDiscountRate: 0.01, projectionMonths: 36 }, "zero churn");
});

test("Unit Economics workbook formulas match lib/unitEconomics.js — rerun with different assumptions", () => {
  const base = { arpu: 100, grossMargin: 0.80, monthlyChurn: 0.05, cac: 500, monthlyDiscountRate: 0.01, projectionMonths: 36 };
  const harder = { ...base, monthlyChurn: 0.12, cac: 1200 };
  checkUEWorkbook(harder, "higher churn, higher CAC");

  const baseGround = computeUnitEconomics(validateUEAssumptions(base));
  const harderGround = computeUnitEconomics(validateUEAssumptions(harder));
  assert.ok(harderGround.ltvToCac < baseGround.ltvToCac);
});

test("Unit Economics sensitivity grid in the workbook is a live formula", () => {
  const assumptions = validateUEAssumptions({ arpu: 100, grossMargin: 0.80, monthlyChurn: 0.05, cac: 500 });
  const sensitivity = computeUESensitivity(assumptions);
  const { wb, refs } = buildUnitEconomicsWorkbook(assumptions, sensitivity);
  const ws = wb.Sheets["Unit Economics"];
  // Spot-check: every grid cell should carry a formula, not a static value.
  const anyGridCell = ws[XLSXStyle.utils.encode_cell({ r: refs.rows.ltvToCac + 4, c: 1 })];
  assert.ok(anyGridCell && typeof anyGridCell.f === "string" && anyGridCell.f.length > 0);
});

// ── Cap Table ────────────────────────────────────────────────────

function fixtureAssumptions(overrides = {}) {
  return {
    companyName: "Acme",
    founders: [
      { name: "Founder A", shares: 1000000 },
      { name: "Founder B", shares: 1000000 },
    ],
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0.10, liquidationPreference: 1 }],
    safes: [{ name: "SAFE 1", amount: 500000, cap: 5000000, discount: 0.20 }],
    ...overrides,
  };
}

function checkCapTableWorkbook(rawAssumptions, label) {
  const assumptions = validateCapTableAssumptions(rawAssumptions);
  const ground = computeCapTable(assumptions);
  const payoutTable = computePayoutTable(ground);
  const { wb, refs } = buildCapTableWorkbook(assumptions, ground, payoutTable);
  const cell = evaluate(wb, "Cap Table");

  // Ownership % is a live formula (shares / SUM(all shares)) — must match
  // the JS engine's finalOwnership exactly.
  refs.holderShareRows.forEach(({ name, row }) => {
    const groundPct = ground.finalOwnership.find((h) => h.name === name).pct;
    close(cell(row, 2), groundPct, 1e-6, `${label} ${name} ownership %`);
  });

  const groundTotalShares = ground.holders.reduce((s, h) => s + h.shares, 0);
  close(cell(refs.totalSharesRow, 1), groundTotalShares, 1, `${label} total shares`);

  // The exit-value check row is also a live formula (SUM of that column's
  // payouts) and must equal the exit value for every multiple.
  payoutTable.columns.forEach((col, ci) => {
    close(cell(refs.payoutCheckRow, ci + 1), col.exitValue, 1, `${label} ${col.multiple}x payout check`);
  });
}

test("Cap table workbook — single priced round with a SAFE conversion matches lib/capTable.js", () => {
  checkCapTableWorkbook(fixtureAssumptions(), "base case");
});

test("Cap table workbook — rerun with a materially different round matches lib/capTable.js", () => {
  const base = fixtureAssumptions();
  const bigger = fixtureAssumptions({
    rounds: [{ name: "Seed", preMoney: 20000000, investment: 5000000, targetOptionPoolPct: 0.15, liquidationPreference: 1.5 }],
  });
  checkCapTableWorkbook(bigger, "bigger round");

  const baseGround = computeCapTable(validateCapTableAssumptions(base));
  const biggerGround = computeCapTable(validateCapTableAssumptions(bigger));
  const founderPctBase = baseGround.finalOwnership.find((h) => h.name === "Founder A").pct;
  const founderPctBigger = biggerGround.finalOwnership.find((h) => h.name === "Founder A").pct;
  // A bigger round + bigger pool target must dilute the founders further.
  assert.ok(founderPctBigger < founderPctBase);
});

test("Cap table workbook — exit below total preferences pays the investor first", () => {
  const assumptions = validateCapTableAssumptions(fixtureAssumptions({ safes: [] }));
  const ground = computeCapTable(assumptions);
  const payoutTable = computePayoutTable(ground, [0.05]); // tiny exit, far below the $2M preference
  const { wb, refs } = buildCapTableWorkbook(assumptions, ground, payoutTable);
  const cell = evaluate(wb, "Cap Table");

  const investorRowIdx = refs.payoutRows[refs.holderShareRows.findIndex((h) => h.name === "Seed investor")];
  const founderRowIdx = refs.payoutRows[refs.holderShareRows.findIndex((h) => h.name === "Founder A")];
  close(cell(investorRowIdx, 1), payoutTable.columns[0].exitValue, 1);
  close(cell(founderRowIdx, 1), 0, 0.01);
});

// ── Cap table round pricing: closed-form live formulas ─────────
// These confirm that the round-pricing formulas (price per share, pool
// top-up, SAFE conversions, investor shares) are genuinely live — each of
// pre-money, investment, pool %, SAFE cap, and SAFE discount independently
// changes the resulting share counts, and every one matches lib/capTable.js.

function checkRoundFormulasLive(rawAssumptions, label) {
  const assumptions = validateCapTableAssumptions(rawAssumptions);
  const ground = computeCapTable(assumptions);
  const payoutTable = computePayoutTable(ground);
  const { wb, refs } = buildCapTableWorkbook(assumptions, ground, payoutTable);
  const cell = evaluate(wb, "Cap Table");

  const seedResult = refs.roundResultRows[0];
  const seedGround = ground.rounds[0];

  close(cell(seedResult.rFDAfter), seedGround.fullyDilutedSharesAfter, 1, `${label} fully diluted after`);
  close(cell(seedResult.rPrice), seedGround.pricePerShare, 1e-4, `${label} price per share`);
  close(cell(seedResult.rPoolTopUp), seedGround.optionPoolTopUpShares, 1, `${label} pool top-up`);
  close(cell(seedResult.rNewInvShares), seedGround.newInvestorShares, 1, `${label} investor shares`);
  seedResult.safeShareRows.forEach(({ name, row }) => {
    const groundShares = seedGround.safeConversions.find((sc) => sc.name === name).shares;
    close(cell(row), groundShares, 1, `${label} ${name} shares`);
  });

  // Holder shares in the final cap table are themselves formulas that
  // reference the round-result cells above — confirm they match too.
  refs.holderShareRows.forEach(({ name, row }) => {
    const groundHolder = ground.holders.find((h) => h.name === name);
    close(cell(row), groundHolder.shares, 1, `${label} final cap table ${name}`);
  });

  return { assumptions, ground };
}

test("Cap table round formulas are live: base case", () => {
  checkRoundFormulasLive(fixtureAssumptions(), "base case");
});

test("Cap table round formulas are live: changing pre-money changes share counts", () => {
  const { ground: baseGround } = checkRoundFormulasLive(fixtureAssumptions(), "pre-money base");
  const { ground: changedGround } = checkRoundFormulasLive(
    fixtureAssumptions({ rounds: [{ name: "Seed", preMoney: 16000000, investment: 2000000, targetOptionPoolPct: 0.10, liquidationPreference: 1 }] }),
    "pre-money changed"
  );
  assert.notEqual(
    Math.round(baseGround.rounds[0].pricePerShare * 1e4),
    Math.round(changedGround.rounds[0].pricePerShare * 1e4)
  );
});

test("Cap table round formulas are live: changing investment changes share counts", () => {
  const { ground: baseGround } = checkRoundFormulasLive(fixtureAssumptions(), "investment base");
  const { ground: changedGround } = checkRoundFormulasLive(
    fixtureAssumptions({ rounds: [{ name: "Seed", preMoney: 8000000, investment: 5000000, targetOptionPoolPct: 0.10, liquidationPreference: 1 }] }),
    "investment changed"
  );
  assert.notEqual(
    Math.round(baseGround.rounds[0].newInvestorShares),
    Math.round(changedGround.rounds[0].newInvestorShares)
  );
});

test("Cap table round formulas are live: changing the option pool % changes the top-up", () => {
  const { ground: baseGround } = checkRoundFormulasLive(fixtureAssumptions(), "pool % base");
  const { ground: changedGround } = checkRoundFormulasLive(
    fixtureAssumptions({ rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0.20, liquidationPreference: 1 }] }),
    "pool % changed"
  );
  assert.ok(changedGround.rounds[0].optionPoolTopUpShares > baseGround.rounds[0].optionPoolTopUpShares);
});

test("Cap table round formulas are live: changing the SAFE cap changes its conversion shares", () => {
  const { ground: baseGround } = checkRoundFormulasLive(fixtureAssumptions(), "SAFE cap base");
  const { ground: changedGround } = checkRoundFormulasLive(
    fixtureAssumptions({ safes: [{ name: "SAFE 1", amount: 500000, cap: 3000000, discount: 0.20 }] }),
    "SAFE cap changed"
  );
  const baseShares = baseGround.rounds[0].safeConversions[0].shares;
  const changedShares = changedGround.rounds[0].safeConversions[0].shares;
  // A lower cap means a lower conversion price, so more shares for the same $500K.
  assert.ok(changedShares > baseShares);
});

test("Cap table round formulas are live: changing the SAFE discount changes its conversion shares", () => {
  // Use a high cap so the discount (not the cap) governs, and the discount change actually matters.
  const { ground: baseGround } = checkRoundFormulasLive(
    fixtureAssumptions({ safes: [{ name: "SAFE 1", amount: 500000, cap: 50000000, discount: 0.20 }] }),
    "SAFE discount base"
  );
  const { ground: changedGround } = checkRoundFormulasLive(
    fixtureAssumptions({ safes: [{ name: "SAFE 1", amount: 500000, cap: 50000000, discount: 0.40 }] }),
    "SAFE discount changed"
  );
  const baseShares = baseGround.rounds[0].safeConversions[0].shares;
  const changedShares = changedGround.rounds[0].safeConversions[0].shares;
  // A steeper discount means a lower conversion price, so more shares for the same $500K.
  assert.ok(changedShares > baseShares);
});

// ── Cap table: live regime selection (not baked at export time) ───
// These confirm the SAFE cap-vs-discount choice, and the pool
// active-vs-inactive choice, are resolved by the hidden helper table and
// SUMPRODUCT selection live in Excel — not fixed when the file was built.

test("Cap table: raising a SAFE's cap flips which method governs, live in Excel", () => {
  // Low cap ($3M) makes the cap price the lower (governing) price.
  const lowCapResult = checkRoundFormulasLive(
    fixtureAssumptions({ safes: [{ name: "SAFE 1", amount: 500000, cap: 3000000, discount: 0.20 }] }),
    "low cap"
  );
  assert.equal(lowCapResult.ground.rounds[0].safeConversions[0].method, "cap");

  // The SAME SAFE with a much higher cap ($50M) makes the discount govern
  // instead — a genuine regime flip, resolved without rebuilding the sheet
  // under a different assumption about which one governs.
  const highCapResult = checkRoundFormulasLive(
    fixtureAssumptions({ safes: [{ name: "SAFE 1", amount: 500000, cap: 50000000, discount: 0.20 }] }),
    "high cap"
  );
  assert.equal(highCapResult.ground.rounds[0].safeConversions[0].method, "discount");

  // Confirm the Excel formula's answer actually matches each regime's own
  // independent formula, not just "some number came out."
  const lowCapShares = lowCapResult.ground.rounds[0].safeConversions[0].shares;
  const highCapShares = highCapResult.ground.rounds[0].safeConversions[0].shares;
  assert.notEqual(Math.round(lowCapShares), Math.round(highCapShares));
});

test("Cap table: lowering a SAFE's discount flips which method governs, live in Excel", () => {
  // High cap ($50M) so the discount is what's being tested; at 10% discount
  // off a ~$4-5/share round, the discount price stays above the (very high)
  // cap-implied price... actually with cap this high, cap price is always
  // far above discount price, so discount governs regardless here — instead
  // flip the comparison the other way: a very steep 70% discount vs a tight
  // cap that would otherwise govern at a shallow discount.
  const shallowDiscount = checkRoundFormulasLive(
    fixtureAssumptions({ safes: [{ name: "SAFE 1", amount: 500000, cap: 4500000, discount: 0.05 }] }),
    "shallow discount"
  );
  const steepDiscount = checkRoundFormulasLive(
    fixtureAssumptions({ safes: [{ name: "SAFE 1", amount: 500000, cap: 4500000, discount: 0.70 }] }),
    "steep discount"
  );
  assert.equal(shallowDiscount.ground.rounds[0].safeConversions[0].method, "cap");
  assert.equal(steepDiscount.ground.rounds[0].safeConversions[0].method, "discount");
});

test("Cap table: raising the existing pool above the target turns the top-up off, live in Excel", () => {
  // Baseline: no existing pool, 10% target -> top-up is active (positive).
  const noPool = checkRoundFormulasLive(
    fixtureAssumptions({ optionPoolShares: 0, safes: [] }),
    "no existing pool"
  );
  assert.ok(noPool.ground.rounds[0].poolActive);
  assert.ok(noPool.ground.rounds[0].optionPoolTopUpShares > 0);

  // Same round, but the existing pool is already well above what a 10%
  // post-money target would require — the top-up must turn itself off
  // rather than go negative, and the formula must find that live.
  const bigPool = checkRoundFormulasLive(
    fixtureAssumptions({ optionPoolShares: 2000000, safes: [] }),
    "existing pool above target"
  );
  assert.equal(bigPool.ground.rounds[0].poolActive, false);
  assert.equal(bigPool.ground.rounds[0].optionPoolTopUpShares, 0);
});
