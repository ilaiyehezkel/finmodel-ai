import test from "node:test";
import assert from "node:assert/strict";
import { computeCapTable, computeExitWaterfall, computePayoutTable, validateAssumptions, solveRoundIterative, DEFAULT_ASSUMPTIONS } from "./capTable.js";

const close = (actual, expected, tolerance = 1) => {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `expected ${actual} to be within ${tolerance} of ${expected}`
  );
};

function founders() {
  return [
    { name: "Founder A", shares: 1000000 },
    { name: "Founder B", shares: 1000000 },
  ];
}

// Fixture 1: a single $8M pre-money / $2M round on a clean 2,000,000-share
// cap table, no pool and no SAFEs. Price per share = 8,000,000 / 2,000,000 =
// $4.00 exactly; the investor buys 2,000,000/4 = 500,000 shares, landing at
// an exact, hand-checkable 20% post-round (2,500,000 fully diluted total).
test("fixture 1: single priced round, no pool, no SAFEs", () => {
  const assumptions = {
    companyName: "Acme",
    founders: founders(),
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [],
  };
  const r = computeCapTable(assumptions);
  const seed = r.rounds[0];
  close(seed.pricePerShare, 4.0, 0.0001);
  close(seed.newInvestorShares, 500000);
  close(seed.fullyDilutedSharesAfter, 2500000);

  const after = seed.ownershipAfter;
  close(after.find((h) => h.name === "Founder A").pct, 0.40, 0.0001);
  close(after.find((h) => h.name === "Founder B").pct, 0.40, 0.0001);
  close(after.find((h) => h.name === "Seed investor").pct, 0.20, 0.0001);

  const before = seed.ownershipBefore;
  close(before.find((h) => h.name === "Founder A").pct, 0.50, 0.0001);
});

// Fixture 2: a SAFE with a $5M cap and 20% discount, in a round priced at
// $10M pre-money on the same 2,000,000-share base. The cap price ($5M/2M =
// $2.50) ends up lower than the discount price (~$3.64), so the cap governs
// and the SAFE converts at exactly $2.50/share -> 500,000/2.50 = 200,000
// shares, independent of the round's own price.
test("fixture 2: SAFE conversion where the cap governs", () => {
  const assumptions = {
    companyName: "Acme",
    founders: founders(),
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 10000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [{ name: "SAFE 1", amount: 500000, cap: 5000000, discount: 0.20 }],
  };
  const r = computeCapTable(assumptions);
  const seed = r.rounds[0];
  close(seed.safeConversions[0].shares, 200000, 1);
  close(seed.pricePerShare, 10000000 / 2200000, 0.001);
  close(seed.fullyDilutedSharesAfter, 2640000, 1);
});

// Fixture 3: same setup but the cap is set very high ($50M), so the 20%
// discount off the round price governs instead. Solving the price
// algebraically: p = (preMoney) / (prior + 500000/(0.8p)) gives p = $4.6875
// exactly, and the SAFE converts at 0.8*4.6875 = $3.75/share -> 133,333.33
// shares. The three post-round share counts still sum to a clean 2,560,000.
test("fixture 3: SAFE conversion where the discount governs", () => {
  const assumptions = {
    companyName: "Acme",
    founders: founders(),
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 10000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [{ name: "SAFE 1", amount: 500000, cap: 50000000, discount: 0.20 }],
  };
  const r = computeCapTable(assumptions);
  const seed = r.rounds[0];
  close(seed.pricePerShare, 4.6875, 0.0001);
  close(seed.safeConversions[0].shares, 133333.33, 1);
  close(seed.fullyDilutedSharesAfter, 2560000, 1);
});

// Fixture 4: a 10% post-money option pool target with no existing pool and
// no SAFEs. Solving the pool-shuffle system algebraically gives an exact
// pool top-up of 2,000,000/7 shares and a fully diluted total of
// 20,000,000/7 — and 2,000,000/7 divided by 20,000,000/7 is exactly 10%.
test("fixture 4: option pool shuffle hits the exact target percentage", () => {
  const assumptions = {
    companyName: "Acme",
    founders: founders(),
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0.10, liquidationPreference: 1 }],
    safes: [],
  };
  const r = computeCapTable(assumptions);
  const seed = r.rounds[0];
  close(seed.optionPoolTopUpShares, 2000000 / 7, 1);
  close(seed.fullyDilutedSharesAfter, 20000000 / 7, 1);

  const poolPct = seed.ownershipAfter.find((h) => h.name === "Option Pool").pct;
  close(poolPct, 0.10, 0.0001);
});

// Fixture 5: a $2M investment at 1x non-participating preference, exit at
// only $1M (below the $2M preference claim). The investor must be paid
// first — getting the entire $1M — and the founders, who only have common
// stock, get $0.
test("fixture 5: exit below total preferences pays the investor first", () => {
  const assumptions = {
    companyName: "Acme",
    founders: founders(),
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [],
  };
  const r = computeCapTable(assumptions);
  const payouts = computeExitWaterfall(r, 1000000);

  const investor = payouts.find((p) => p.name === "Seed investor");
  close(investor.payout, 1000000);
  assert.equal(investor.converted, false);

  payouts.filter((p) => p.type === "founder").forEach((p) => close(p.payout, 0));
});

test("exit well above preferences: investor converts and shares pro rata", () => {
  const assumptions = {
    companyName: "Acme",
    founders: founders(),
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [],
  };
  const r = computeCapTable(assumptions);
  const payouts = computeExitWaterfall(r, 50000000);
  const investor = payouts.find((p) => p.name === "Seed investor");
  assert.equal(investor.converted, true);
  close(investor.payout, 50000000 * 0.20, 1); // 20% as-converted ownership

  const total = payouts.reduce((s, p) => s + p.payout, 0);
  close(total, 50000000, 1);
});

test("computePayoutTable sweeps several multiples of the last round's post-money", () => {
  const assumptions = {
    companyName: "Acme",
    founders: founders(),
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [],
  };
  const r = computeCapTable(assumptions);
  const table = computePayoutTable(r, [0.5, 1, 2, 5, 10]);
  close(table.basis, 10000000); // last round's post-money: 8M pre + 2M invested
  assert.equal(table.columns.length, 5);
  close(table.columns[1].exitValue, 10000000); // 1x
  close(table.columns[4].exitValue, 100000000); // 10x

  // Every column's payouts must sum back to the exit value.
  table.columns.forEach((col) => {
    const total = col.payouts.reduce((s, p) => s + p.payout, 0);
    close(total, col.exitValue, 1);
  });
});

test("validateAssumptions fills in missing fields with defaults", () => {
  const a = validateAssumptions({ companyName: "Acme" });
  assert.equal(a.companyName, "Acme");
  assert.deepEqual(a.founders, DEFAULT_ASSUMPTIONS.founders);
  assert.equal(a.rounds.length, 1);
  assert.deepEqual(a.safes, []);
});

test("validateAssumptions tolerates garbage input entirely", () => {
  const a = validateAssumptions(null);
  assert.deepEqual(a, DEFAULT_ASSUMPTIONS);

  const b = validateAssumptions({ founders: "nonsense", rounds: [{ preMoney: "bad" }] });
  assert.deepEqual(b.founders, DEFAULT_ASSUMPTIONS.founders);
  assert.equal(b.rounds[0].preMoney, DEFAULT_ASSUMPTIONS.rounds[0].preMoney);
});

// ── Closed form vs. iterative cross-check ───────────────────────
// computeCapTable uses the closed-form solver (solveRoundClosedForm,
// internal to lib/capTable.js). These tests feed the exact same inputs to
// the kept-for-testing iterative solver and confirm the two always agree —
// across a plain round, both SAFE-governance cases, a pool-only shuffle,
// and a round where the pool shuffle and a SAFE both apply at once.

function closeTo(actual, expected, tolerance, label) {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `${label}: expected ${actual} to be within ${tolerance} of ${expected}`
  );
}

function crossCheckFirstRound(assumptions, label) {
  const result = computeCapTable(assumptions);
  const seed = result.rounds[0];
  const priorFDShares = assumptions.founders.reduce((s, f) => s + f.shares, 0) + assumptions.optionPoolShares;
  const iterative = solveRoundIterative(priorFDShares, assumptions.optionPoolShares, assumptions.rounds[0], assumptions.safes);

  closeTo(seed.pricePerShare, iterative.price, 1e-4, `${label} price`);
  closeTo(seed.optionPoolTopUpShares, iterative.poolTopUp, 1, `${label} pool top-up`);
  closeTo(seed.newInvestorShares, iterative.newInvestorShares, 1, `${label} investor shares`);
  seed.safeConversions.forEach((sc) => {
    closeTo(sc.shares, iterative.safeSharesByName[sc.name], 1, `${label} ${sc.name} shares`);
  });
}

test("closed form matches iterative: plain round, no pool, no SAFEs", () => {
  crossCheckFirstRound(validateAssumptions({
    founders: [{ name: "A", shares: 1000000 }, { name: "B", shares: 1000000 }],
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [],
  }), "plain round");
});

test("closed form matches iterative: SAFE where the cap governs", () => {
  crossCheckFirstRound(validateAssumptions({
    founders: [{ name: "A", shares: 1000000 }, { name: "B", shares: 1000000 }],
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 10000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [{ name: "SAFE 1", amount: 500000, cap: 5000000, discount: 0.20 }],
  }), "cap governs");
});

test("closed form matches iterative: SAFE where the discount governs", () => {
  crossCheckFirstRound(validateAssumptions({
    founders: [{ name: "A", shares: 1000000 }, { name: "B", shares: 1000000 }],
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 10000000, investment: 2000000, targetOptionPoolPct: 0, liquidationPreference: 1 }],
    safes: [{ name: "SAFE 1", amount: 500000, cap: 50000000, discount: 0.20 }],
  }), "discount governs");
});

test("closed form matches iterative: option pool shuffle only", () => {
  crossCheckFirstRound(validateAssumptions({
    founders: [{ name: "A", shares: 1000000 }, { name: "B", shares: 1000000 }],
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0.10, liquidationPreference: 1 }],
    safes: [],
  }), "pool shuffle only");
});

test("closed form matches iterative: pool shuffle and a SAFE together", () => {
  crossCheckFirstRound(validateAssumptions({
    founders: [{ name: "A", shares: 1000000 }, { name: "B", shares: 1000000 }],
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 10000000, investment: 2000000, targetOptionPoolPct: 0.10, liquidationPreference: 1 }],
    safes: [{ name: "SAFE 1", amount: 500000, cap: 5000000, discount: 0.20 }],
  }), "pool + SAFE");
});

test("closed form matches iterative: existing option pool already above target", () => {
  // existingPoolShares is large enough that the target is already met without
  // a top-up — exercises the poolActive=false branch with a nonzero pool.
  crossCheckFirstRound(validateAssumptions({
    founders: [{ name: "A", shares: 1000000 }, { name: "B", shares: 1000000 }],
    optionPoolShares: 2000000,
    rounds: [{ name: "Seed", preMoney: 8000000, investment: 2000000, targetOptionPoolPct: 0.10, liquidationPreference: 1 }],
    safes: [],
  }), "pool already above target");
});

test("closed form matches iterative: multiple SAFEs with mixed governance", () => {
  crossCheckFirstRound(validateAssumptions({
    founders: [{ name: "A", shares: 1000000 }, { name: "B", shares: 1000000 }],
    optionPoolShares: 0,
    rounds: [{ name: "Seed", preMoney: 12000000, investment: 3000000, targetOptionPoolPct: 0.08, liquidationPreference: 1 }],
    safes: [
      { name: "SAFE 1", amount: 500000, cap: 5000000, discount: 0.20 }, // cap governs
      { name: "SAFE 2", amount: 300000, cap: 50000000, discount: 0.20 }, // discount governs
    ],
  }), "mixed governance");
});
