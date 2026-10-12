import test from "node:test";
import assert from "node:assert/strict";
import { computeDCF, computeSensitivity, validateAssumptions, DEFAULT_ASSUMPTIONS } from "./dcf.js";

const close = (actual, expected, tolerance = 0.01) => {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `expected ${actual} to be within ${tolerance} of ${expected}`
  );
};

// Fixture 1: WACC (10%) deliberately equals revenue growth (10%), which makes
// every year's unlevered FCF grow at exactly the discount rate. That forces
// every year's present value to collapse to UFCF(year 1) / (1 + WACC), a
// constant — 101.75 / 1.1 = 92.5 — so the whole 5-year PV sum is hand-checkable
// as 5 x 92.50 = 462.50 without a calculator.
const fixture1 = {
  baseRevenue: 1000,
  revenueGrowth: [0.10, 0.10, 0.10, 0.10, 0.10],
  ebitdaMargin: 0.20,
  taxRate: 0.25,
  daPct: 0.05,
  capexPct: 0.05,
  nwcPct: 0.02,
  wacc: 0.10,
  terminalGrowth: 0.02,
  netDebt: 100,
  sharesOutstanding: 100,
};

test("fixture 1: year-by-year projections match hand calculation", () => {
  const { years } = computeDCF(fixture1);
  close(years[0].revenue, 1100.00);
  close(years[0].ebitda, 220.00);
  close(years[0].ebit, 165.00);
  close(years[0].nopat, 123.75);
  close(years[0].ufcf, 101.75);
  close(years[0].pv, 92.50);

  close(years[4].revenue, 1610.51);
  close(years[4].ufcf, 148.972175, 0.001);
  close(years[4].pv, 92.50);
});

test("fixture 1: sum of PV, terminal value, and implied share price", () => {
  const r = computeDCF(fixture1);
  close(r.sumPV, 462.50);
  close(r.terminalValue, 1899.395231, 0.01);
  close(r.pvTerminalValue, 1179.375, 0.01);
  close(r.enterpriseValue, 1641.875, 0.01);
  close(r.equityValue, 1541.875, 0.01); // enterprise value - $100 net debt
  close(r.impliedSharePrice, 15.41875, 0.001); // equity value / 100 shares
});

// Fixture 2: zero revenue growth keeps every year flat at 100, so each line
// item is identical year over year — only the discount factor changes.
const fixture2 = {
  baseRevenue: 100,
  revenueGrowth: [0, 0, 0, 0, 0],
  ebitdaMargin: 0.30,
  taxRate: 0.20,
  daPct: 0.04,
  capexPct: 0.04,
  nwcPct: 0,
  wacc: 0.08,
  terminalGrowth: 0.02,
  netDebt: 0,
  sharesOutstanding: 10,
};

test("fixture 2: flat revenue keeps every year's fundamentals identical", () => {
  const { years } = computeDCF(fixture2);
  for (const y of years) {
    close(y.revenue, 100.00);
    close(y.ebitda, 30.00);
    close(y.ebit, 26.00);
    close(y.nopat, 20.80);
    close(y.ufcf, 20.80);
  }
  // Discount factors must strictly decrease year over year.
  for (let i = 1; i < years.length; i++) {
    assert.ok(years[i].discountFactor < years[i - 1].discountFactor);
  }
});

test("fixture 2: enterprise value, equity value (no net debt), and price", () => {
  const r = computeDCF(fixture2);
  close(r.sumPV, 83.048369, 0.01);
  close(r.terminalValue, 353.60, 0.01);
  close(r.pvTerminalValue, 240.654218, 0.01);
  close(r.enterpriseValue, 323.702587, 0.01);
  close(r.equityValue, r.enterpriseValue); // net debt is zero
  close(r.impliedSharePrice, 32.370259, 0.001);
});

test("WACC <= terminal growth is flagged invalid, not Infinity/NaN", () => {
  const r = computeDCF({ ...fixture1, wacc: 0.05, terminalGrowth: 0.05 });
  assert.equal(r.terminalValid, false);
  assert.equal(r.terminalValue, null);
  assert.equal(r.enterpriseValue, null);
  assert.equal(r.equityValue, null);
  assert.equal(r.impliedSharePrice, null);

  const r2 = computeDCF({ ...fixture1, wacc: 0.04, terminalGrowth: 0.05 });
  assert.equal(r2.terminalValid, false);
});

test("sensitivity grid marks WACC <= terminal growth cells as errors", () => {
  const { waccSteps, growthSteps, grid } = computeSensitivity(fixture1, {
    waccSteps: [0.04, 0.06, 0.08, 0.10, 0.12],
    growthSteps: [0.02, 0.04, 0.06, 0.08, 0.10],
  });

  assert.equal(waccSteps.length, 5);
  assert.equal(growthSteps.length, 5);

  // wacc=0.04 vs growth=0.04 and above -> error (wacc <= growth)
  assert.equal(grid[0][1].error, true); // 0.04 wacc, 0.04 growth
  assert.equal(grid[0][0].error, false); // 0.04 wacc, 0.02 growth is valid

  // waccSteps[3] = 0.10 and growthSteps[0] = 0.02 exactly match fixture 1's own
  // wacc/terminalGrowth, so that cell must reproduce fixture 1's price exactly.
  const base = computeDCF(fixture1).impliedSharePrice;
  close(grid[3][0].price, base, 0.01);
});

test("validateAssumptions fills in missing fields with defaults", () => {
  const a = validateAssumptions({ companyName: "Acme", wacc: 0.15 });
  assert.equal(a.companyName, "Acme");
  assert.equal(a.wacc, 0.15);
  assert.equal(a.baseRevenue, DEFAULT_ASSUMPTIONS.baseRevenue);
  assert.deepEqual(a.revenueGrowth, DEFAULT_ASSUMPTIONS.revenueGrowth);
});

test("validateAssumptions tolerates garbage input entirely", () => {
  const a = validateAssumptions(null);
  assert.deepEqual(a, DEFAULT_ASSUMPTIONS);

  const b = validateAssumptions({ revenueGrowth: "not an array", wacc: "12%" });
  assert.deepEqual(b.revenueGrowth, DEFAULT_ASSUMPTIONS.revenueGrowth);
  assert.equal(b.wacc, DEFAULT_ASSUMPTIONS.wacc);
});
