import test from "node:test";
import assert from "node:assert/strict";
import { computeUnitEconomics, computeSensitivity, validateAssumptions, DEFAULT_ASSUMPTIONS } from "./unitEconomics.js";

const close = (actual, expected, tolerance = 0.01) => {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `expected ${actual} to be within ${tolerance} of ${expected}`
  );
};

// Fixture 1: $100 ARPU, 80% gross margin -> $80 gross profit/month, 5% monthly
// churn, $500 CAC. Simple LTV = 80/0.05 = $1,600 and LTV/CAC = 3.2 are both
// exact single-division hand calculations. The payback month is checkable by
// summing 0.95^(m-1)*80 by hand through month 7 ($482.66, not yet paid back)
// and month 8 ($538.53, paid back), then interpolating between them.
const fixture1 = {
  arpu: 100,
  grossMargin: 0.80,
  monthlyChurn: 0.05,
  cac: 500,
  monthlyDiscountRate: 0.01,
  projectionMonths: 36,
};

test("fixture 1: gross profit, average lifetime, and simple LTV/CAC", () => {
  const r = computeUnitEconomics(fixture1);
  close(r.grossProfitPerMonth, 80);
  close(r.avgLifetimeMonths, 20);
  close(r.simpleLTV, 1600);
  close(r.ltvToCac, 3.2);
});

test("fixture 1: cohort table survival and cumulative gross profit", () => {
  const { months } = computeUnitEconomics(fixture1);
  close(months[0].survival, 1);
  close(months[0].expectedGrossProfit, 80);
  close(months[0].cumulativeGrossProfit, 80);
  assert.equal(months[0].paybackReached, false);

  close(months[6].cumulativeGrossProfit, 482.6603, 0.01); // month 7, not yet paid back
  assert.equal(months[6].paybackReached, false);

  close(months[7].cumulativeGrossProfit, 538.5273, 0.01); // month 8, paid back
  assert.equal(months[7].paybackReached, true);
});

test("fixture 1: CAC payback month is interpolated between month 7 and 8", () => {
  const { cacPaybackMonths } = computeUnitEconomics(fixture1);
  close(cacPaybackMonths, 7.3104, 0.01);
});

test("fixture 1: discounted LTV is less than simple (undiscounted) LTV", () => {
  const r = computeUnitEconomics(fixture1);
  close(r.discountedLTV, 1186.2989, 0.1);
  assert.ok(r.discountedLTV < r.simpleLTV);
});

// Fixture 2: zero churn means the customer never leaves, so the textbook
// LTV formula (gross profit / churn) divides by zero. The engine must fall
// back to a finite, horizon-capped LTV instead of Infinity, and the cohort
// table's survival should stay at 1 (100%) for every month.
const fixture2 = {
  arpu: 100,
  grossMargin: 0.80,
  monthlyChurn: 0,
  cac: 500,
  monthlyDiscountRate: 0.01,
  projectionMonths: 36,
};

test("fixture 2 (zero churn): avgLifetimeMonths is null, not Infinity", () => {
  const r = computeUnitEconomics(fixture2);
  assert.equal(r.avgLifetimeMonths, null);
  assert.ok(Number.isFinite(r.simpleLTV));
  close(r.simpleLTV, 80 * 36); // horizon-capped: $80/mo for all 36 months
  close(r.ltvToCac, 5.76);
});

test("fixture 2 (zero churn): survival never decays and payback is a simple division", () => {
  const { months, cacPaybackMonths } = computeUnitEconomics(fixture2);
  months.forEach((m) => close(m.survival, 1));
  close(cacPaybackMonths, 500 / 80); // 6.25 months, no interpolation needed
});

test("sensitivity grid never divides by zero even near the edges", () => {
  const { churnSteps, cacSteps, grid } = computeSensitivity(fixture1);
  assert.equal(churnSteps.length, 5);
  assert.equal(cacSteps.length, 5);
  grid.forEach((row) =>
    row.forEach((cell) => {
      assert.ok(cell.monthlyChurn >= 0);
      assert.ok(cell.cac > 0);
      if (!cell.error) assert.ok(Number.isFinite(cell.ltvToCac));
    })
  );
});

test("sensitivity grid: higher churn or higher CAC both lower LTV/CAC", () => {
  const { grid } = computeSensitivity(fixture1, {
    churnSteps: [0.02, 0.05, 0.08],
    cacSteps: [300, 500, 700],
  });
  // Lower churn (row 0) beats higher churn (row 2) at the same CAC (col 1).
  assert.ok(grid[0][1].ltvToCac > grid[2][1].ltvToCac);
  // Lower CAC (col 0) beats higher CAC (col 2) at the same churn (row 1).
  assert.ok(grid[1][0].ltvToCac > grid[1][2].ltvToCac);
});

test("validateAssumptions fills in missing fields with defaults", () => {
  const a = validateAssumptions({ companyName: "Acme", cac: 999 });
  assert.equal(a.companyName, "Acme");
  assert.equal(a.cac, 999);
  assert.equal(a.arpu, DEFAULT_ASSUMPTIONS.arpu);
});

test("validateAssumptions tolerates garbage input and clamps negative churn to zero", () => {
  const a = validateAssumptions(null);
  assert.deepEqual(a, DEFAULT_ASSUMPTIONS);

  const b = validateAssumptions({ monthlyChurn: -0.5, projectionMonths: "bad" });
  assert.equal(b.monthlyChurn, 0);
  assert.equal(b.projectionMonths, DEFAULT_ASSUMPTIONS.projectionMonths);
});
