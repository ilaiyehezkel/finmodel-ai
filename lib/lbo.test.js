import test from "node:test";
import assert from "node:assert/strict";
import { computeLBO, computeSensitivity, validateAssumptions, irr, DEFAULT_ASSUMPTIONS } from "./lbo.js";

const close = (actual, expected, tolerance = 0.01) => {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `expected ${actual} to be within ${tolerance} of ${expected}`
  );
};

// Fixture 1: flat revenue (0% growth) with D&A == CapEx (they cancel in FCF)
// and a 0% cash sweep, so the only moving part each year is mandatory senior
// amortization (10% of the original $900 principal = $90/yr) and the
// resulting interest roll-down — all hand-checkable without a calculator.
const fixture1 = {
  entryRevenue: 1000,
  revenueGrowth: [0, 0, 0, 0, 0],
  ebitdaMargin: 0.30,
  daPct: 0.05,
  capexPct: 0.05,
  nwcPct: 0,
  taxRate: 0.25,
  entryMultiple: 5,
  exitMultiple: 5,
  holdingPeriod: 5,
  feePct: 0,
  minCash: 0,
  seniorDebt: { multiple: 3, rate: 0.10, amortPct: 0.10 },
  subDebt: { multiple: 1, rate: 0.12 },
  sweepPct: 0,
};

test("fixture 1: sources & uses — sponsor equity is the plug", () => {
  const { sourcesUses } = computeLBO(fixture1);
  close(sourcesUses.entryEbitda, 300);
  close(sourcesUses.entryEV, 1500);
  close(sourcesUses.seniorDebt, 900);
  close(sourcesUses.subDebt, 300);
  close(sourcesUses.sponsorEquity, 300); // 1500 - 900 - 300
});

test("fixture 1: year 1 and year 5 projections match hand calculation", () => {
  const { years } = computeLBO(fixture1);
  close(years[0].interestExpense, 126.00); // 900*0.10 + 300*0.12
  close(years[0].netIncome, 93.00);
  close(years[0].fcfForDebtPaydown, 93.00); // D&A and CapEx cancel, NWC is 0
  close(years[0].mandatoryAmort, 90.00); // 10% of original $900
  close(years[0].seniorEndBalance, 810.00);
  close(years[0].cash, 3.00); // 93 - 90 mandatory, no sweep

  close(years[4].seniorEndBalance, 450.00); // 900 - 5 x 90
  close(years[4].cash, 82.50);
  close(years[4].netDebt, 667.50); // (450 + 300) - 82.50
});

test("fixture 1: exit, MOIC, and IRR", () => {
  const r = computeLBO(fixture1);
  assert.equal(r.valid, true);
  close(r.exit.exitEV, 1500);
  close(r.exit.exitNetDebt, 667.50);
  close(r.exit.exitEquityValue, 832.50);
  close(r.returns.moic, 2.775, 0.001);
  close(r.returns.irr * 100, 22.6458, 0.05); // 2.775^(1/5) - 1
});

// Fixture 2: a tiny, 50%-of-EBITDA senior tranche with a 50% mandatory
// amortization rate and a 100% cash sweep pays itself off entirely in year 1
// — this checks that mandatory amort + sweep floor at the beginning balance
// (never go negative) and that once debt hits zero, 100% of FCF flows to cash
// with zero interest expense in every following year.
const fixture2 = {
  entryRevenue: 1000,
  revenueGrowth: [0, 0, 0, 0, 0],
  ebitdaMargin: 0.40,
  daPct: 0.05,
  capexPct: 0.05,
  nwcPct: 0,
  taxRate: 0.25,
  entryMultiple: 4,
  exitMultiple: 4,
  holdingPeriod: 5,
  feePct: 0,
  minCash: 0,
  seniorDebt: { multiple: 0.5, rate: 0.08, amortPct: 0.50 },
  subDebt: { multiple: 0, rate: 0.11 },
  sweepPct: 1.0,
};

test("fixture 2: senior debt is fully retired in year 1", () => {
  const { years, sourcesUses } = computeLBO(fixture2);
  close(sourcesUses.seniorDebt, 200);
  close(years[0].mandatoryAmort, 100);
  close(years[0].cashSweep, 100);
  close(years[0].seniorEndBalance, 0);
  close(years[0].netDebt, -50.50); // net cash position already
});

test("fixture 2: zero debt and zero interest hold for every subsequent year", () => {
  const { years } = computeLBO(fixture2);
  for (let i = 1; i < years.length; i++) {
    assert.equal(years[i].seniorBeginBalance, 0);
    assert.equal(years[i].seniorEndBalance, 0);
    assert.equal(years[i].interestExpense, 0);
    assert.equal(years[i].mandatoryAmort, 0);
    assert.equal(years[i].cashSweep, 0);
  }
  // Cash keeps accumulating year over year once debt is gone.
  for (let i = 1; i < years.length; i++) {
    assert.ok(years[i].cash > years[i - 1].cash);
  }
  close(years[4].cash, 1100.50);
});

test("fixture 2: net cash position at exit and resulting MOIC/IRR", () => {
  const r = computeLBO(fixture2);
  close(r.sourcesUses.sponsorEquity, 1400);
  close(r.exit.exitNetDebt, -1100.50);
  close(r.exit.exitEquityValue, 2700.50);
  close(r.returns.moic, 1.928929, 0.001);
  close(r.returns.irr * 100, 14.0416, 0.05);
});

test("generic IRR solver matches the two-cash-flow closed-form result", () => {
  // For a simple -equity ... +proceeds series, IRR = (proceeds/equity)^(1/n) - 1.
  const cfs = [-300, 0, 0, 0, 0, 832.50];
  const expected = Math.pow(832.50 / 300, 1 / 5) - 1;
  close(irr(cfs), expected, 1e-6);
});

test("sponsor equity goes negative when debt exceeds the purchase price", () => {
  // Same deal, but entry multiple (3x) is below total leverage (3x + 1x = 4x).
  const r = computeLBO({ ...fixture1, entryMultiple: 3 });
  assert.equal(r.valid, false);
  close(r.sourcesUses.sponsorEquity, -300);
  assert.equal(r.returns.moic, null);
  assert.equal(r.returns.irr, null);
});

test("sensitivity grid flags sponsor-equity-negative cells as errors", () => {
  const { entrySteps, exitSteps, grid } = computeSensitivity(fixture1, {
    entrySteps: [3, 3.5, 4, 4.5, 5],
    exitSteps: [4, 4.5, 5, 5.5, 6],
  });
  assert.equal(entrySteps.length, 5);
  assert.equal(exitSteps.length, 5);

  // entryMultiple=3 -> sponsor equity negative at every exit multiple.
  grid[0].forEach((cell) => assert.equal(cell.error, true));

  // entrySteps[4]=5, exitSteps[2]=5 reproduces fixture 1's own base case exactly.
  const base = computeLBO(fixture1).returns.irr;
  close(grid[4][2].irr, base, 1e-6);
});

test("validateAssumptions fills in missing fields, including nested debt tranches", () => {
  const a = validateAssumptions({ companyName: "Acme", entryMultiple: 9, seniorDebt: { rate: 0.09 } });
  assert.equal(a.companyName, "Acme");
  assert.equal(a.entryMultiple, 9);
  assert.equal(a.seniorDebt.rate, 0.09);
  assert.equal(a.seniorDebt.multiple, DEFAULT_ASSUMPTIONS.seniorDebt.multiple);
  assert.deepEqual(a.subDebt, DEFAULT_ASSUMPTIONS.subDebt);
  assert.deepEqual(a.revenueGrowth, DEFAULT_ASSUMPTIONS.revenueGrowth);
});

test("validateAssumptions tolerates garbage input entirely and clamps holding period", () => {
  const a = validateAssumptions(null);
  assert.deepEqual(a, DEFAULT_ASSUMPTIONS);

  const b = validateAssumptions({ holdingPeriod: 11, seniorDebt: "nonsense" });
  assert.equal(b.holdingPeriod, 5); // clamped to the 1-5 range
  assert.deepEqual(b.seniorDebt, DEFAULT_ASSUMPTIONS.seniorDebt);
});
