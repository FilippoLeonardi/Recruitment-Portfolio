// calc.js — live in-browser LBO calculation engine.
// A faithful port of the model math. No circular references (interest uses
// beginning-of-period balances), so it's a clean sequential computation.
// Works as a CommonJS module (Node tests) and as a plain <script> (sets window.LBOCalc).
(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  if (typeof window !== "undefined") window.LBOCalc = api;
})(this, function () {
  "use strict";

  const DEFAULTS = {
    entryMultiple: 10.0, transFees: 0.020, finFees: 0.015, mgmtRollover: 0.10,
    revolverPct: 0.10, revolverRate: 0.045, revolverAmrt: 0.00,
    tlbPct: 0.35, tlbRate: 0.055, tlbAmrt: 0.01,
    ssnPct: 0.20, ssnRate: 0.065, ssnAmrt: 0.00,
    mezPct: 0.10, mezRate: 0.090, mezAmrt: 0.00, mezPik: 0.02,
    revenueLTM: 400.0, ebitdaLTM: 100.0, daLTM: 20.0, capexLTM: 16.0, nwcLTM: 8.0,
    taxRate: 0.25, growth: [0.05, 0.05, 0.05, 0.05, 0.05], marginExp: 0.005,
    exitYear: 5, exitMultiple: 11.0, exitFee: 0.015,
    openingPPE: 80.0, openingNWC: 60.0,
  };

  // Excel-style IRR (Newton–Raphson, bisection fallback), seeded at 0.20.
  function irr(cashflows, guess) {
    const npv = (r) => cashflows.reduce((s, cf, i) => s + cf / Math.pow(1 + r, i), 0);
    const dnpv = (r) => cashflows.reduce((s, cf, i) => s - (i * cf) / Math.pow(1 + r, i + 1), 0);
    let r = guess == null ? 0.2 : guess;
    for (let k = 0; k < 100; k++) {
      const f = npv(r), d = dnpv(r);
      if (Math.abs(f) < 1e-10) return r;
      if (d === 0) break;
      const rn = r - f / d;
      if (!isFinite(rn)) break;
      if (Math.abs(rn - r) < 1e-12) return rn;
      r = rn;
    }
    let lo = -0.9999, hi = 10, flo = npv(lo), fhi = npv(hi);
    if (flo * fhi > 0) return r;
    for (let k = 0; k < 200; k++) {
      const mid = (lo + hi) / 2, fm = npv(mid);
      if (Math.abs(fm) < 1e-11) return mid;
      if (flo * fm < 0) { hi = mid; } else { lo = mid; flo = fm; }
    }
    return (lo + hi) / 2;
  }

  function computeLBO(userInputs) {
    const V = Object.assign({}, DEFAULTS, userInputs || {});
    const N = 5;
    const tax = V.taxRate;
    const g = Array.isArray(V.growth)
      ? [0, 1, 2, 3, 4].map((i) => (V.growth[i] == null ? 0 : V.growth[i]))
      : [V.growth, V.growth, V.growth, V.growth, V.growth];

    // ── Assumptions ──────────────────────────────────────────────────────────
    const ev = V.ebitdaLTM * V.entryMultiple;
    const tlb0 = V.tlbPct * ev, ssn0 = V.ssnPct * ev, mez0 = V.mezPct * ev;
    const totalDebt = tlb0 + ssn0 + mez0;

    // Operating build (index 0 = LTM, 1..5 = years)
    const rev = [V.revenueLTM];
    for (let t = 1; t <= N; t++) rev[t] = rev[t - 1] * (1 + g[t - 1]);
    const marginLTM = V.revenueLTM ? V.ebitdaLTM / V.revenueLTM : 0;
    const margin = [marginLTM];
    for (let t = 1; t <= N; t++) margin[t] = margin[t - 1] + V.marginExp;
    const ebitda = [V.ebitdaLTM];
    for (let t = 1; t <= N; t++) ebitda[t] = rev[t] * margin[t];
    const daRatio = V.revenueLTM ? V.daLTM / V.revenueLTM : 0;
    const capexRatio = V.revenueLTM ? V.capexLTM / V.revenueLTM : 0;
    const nwcRatio = V.revenueLTM ? V.nwcLTM / V.revenueLTM : 0;
    const da = [V.daLTM], capex = [V.capexLTM], nwc = [V.nwcLTM];
    for (let t = 1; t <= N; t++) { da[t] = rev[t] * daRatio; capex[t] = rev[t] * capexRatio; nwc[t] = rev[t] * nwcRatio; }
    const ebit = rev.map((_, t) => ebitda[t] - da[t]);

    // ── Sources & Uses ───────────────────────────────────────────────────────
    const eqPool = ev * (1 + V.transFees) + totalDebt * V.finFees - totalDebt;
    const totalEquity = eqPool;
    const sponsorEquity = eqPool * (1 - V.mgmtRollover);
    const mgmtEquity = eqPool * V.mgmtRollover;
    const totalSources = totalDebt + totalEquity;
    const transFee$ = ev * V.transFees;
    const finFee$ = totalDebt * V.finFees;
    const totalUses = ev + transFee$ + finFee$;
    const suCheck = totalSources - totalUses;
    const dfc = finFee$;

    // ── Debt Schedule + Cash Flow (year by year) ─────────────────────────────
    const tlbBeg = [], tlbAmrt = [], tlbOpt = [], tlbEnd = [], tlbInt = [];
    const ssnBeg = [], ssnOpt = [], ssnEnd = [], ssnInt = [];
    const mezBeg = [], mezPik = [], mezOpt = [], mezEnd = [], mezCashInt = [], mezPikInt = [];
    const totCashInt = [], totPik = [], totInt = [], totMand = [], totDebtEnd = [];
    const cfInt = [], cfTax = [], ofcf = [], avail = [], sweep = [], ncf = [], begCash = [], endCash = [];

    tlbBeg[0] = tlb0; tlbAmrt[0] = 0; tlbOpt[0] = 0; tlbEnd[0] = tlb0; tlbInt[0] = 0;
    ssnBeg[0] = ssn0; ssnOpt[0] = 0; ssnEnd[0] = ssn0; ssnInt[0] = 0;
    mezBeg[0] = mez0; mezPik[0] = 0; mezOpt[0] = 0; mezEnd[0] = mez0; mezCashInt[0] = 0; mezPikInt[0] = 0;
    totCashInt[0] = 0; totPik[0] = 0; totInt[0] = 0; totMand[0] = 0;
    totDebtEnd[0] = tlbEnd[0] + ssnEnd[0] + mezEnd[0];
    begCash[0] = 0; endCash[0] = 0;

    for (let t = 1; t <= N; t++) {
      tlbBeg[t] = tlbEnd[t - 1];
      ssnBeg[t] = ssnEnd[t - 1];
      mezBeg[t] = mezEnd[t - 1];

      tlbInt[t] = tlbBeg[t] * V.tlbRate;
      ssnInt[t] = ssnBeg[t] * V.ssnRate;
      mezPik[t] = mezBeg[t] * V.mezPik;
      mezCashInt[t] = mezBeg[t] * (V.mezRate - V.mezPik);
      mezPikInt[t] = mezPik[t];
      tlbAmrt[t] = Math.min(tlb0 * V.tlbAmrt, tlbBeg[t]); // capped at outstanding balance

      totCashInt[t] = tlbInt[t] + ssnInt[t] + mezCashInt[t];
      totPik[t] = mezPikInt[t];
      totInt[t] = totCashInt[t] + totPik[t];
      totMand[t] = tlbAmrt[t];                    // + SSN (0)

      cfInt[t] = totCashInt[t];
      cfTax[t] = Math.max((ebitda[t] - da[t] - totInt[t]) * tax, 0);
      ofcf[t] = ebitda[t] - cfInt[t] - cfTax[t] - capex[t] - nwc[t];
      avail[t] = Math.max(ofcf[t] - totMand[t], 0);

      tlbOpt[t] = Math.min(Math.max(avail[t], 0), Math.max(tlbBeg[t] - tlbAmrt[t], 0));
      ssnOpt[t] = Math.min(Math.max(avail[t] - tlbOpt[t], 0), Math.max(ssnBeg[t], 0));
      mezOpt[t] = Math.min(Math.max(avail[t] - tlbOpt[t] - ssnOpt[t], 0), Math.max(mezBeg[t] + mezPik[t], 0));

      tlbEnd[t] = Math.max(tlbBeg[t] - tlbAmrt[t] - tlbOpt[t], 0);
      ssnEnd[t] = Math.max(ssnBeg[t] - ssnOpt[t], 0);
      mezEnd[t] = Math.max(mezBeg[t] + mezPik[t] - mezOpt[t], 0);
      totDebtEnd[t] = tlbEnd[t] + ssnEnd[t] + mezEnd[t];

      sweep[t] = tlbOpt[t] + ssnOpt[t] + mezOpt[t];
      ncf[t] = avail[t] - sweep[t];
      begCash[t] = endCash[t - 1];
      endCash[t] = begCash[t] + ncf[t];
    }

    // ── Income Statement ─────────────────────────────────────────────────────
    const isEbt = [null], isTax = [null], isNi = [null];
    for (let t = 1; t <= N; t++) {
      const ebt = ebit[t] - totInt[t];
      isEbt[t] = ebt;
      isTax[t] = Math.max(ebt * tax, 0);
      isNi[t] = ebt - isTax[t];
    }

    // ── Balance Sheet (self-check must be ~0 every year) ─────────────────────
    const gw = ev * (1 + V.transFees) - V.openingNWC - V.openingPPE;
    const bsCash = [], bsNwc = [], bsPpe = [], bsTast = [], bsTdbt = [], bsRe = [], bsTeq = [], bsChk = [];
    bsCash[0] = 0; bsNwc[0] = V.openingNWC; bsPpe[0] = V.openingPPE;
    bsTast[0] = bsCash[0] + bsNwc[0] + bsPpe[0] + gw + dfc;
    bsTdbt[0] = totDebtEnd[0]; bsRe[0] = 0; bsTeq[0] = totalEquity + bsRe[0];
    bsChk[0] = bsTast[0] - (bsTdbt[0] + bsTeq[0]);
    for (let t = 1; t <= N; t++) {
      bsCash[t] = endCash[t];
      bsNwc[t] = bsNwc[t - 1] + nwc[t];
      bsPpe[t] = bsPpe[t - 1] + capex[t] - da[t];
      bsTast[t] = bsCash[t] + bsNwc[t] + bsPpe[t] + gw + dfc;
      bsTdbt[t] = totDebtEnd[t];
      bsRe[t] = bsRe[t - 1] + isNi[t];
      bsTeq[t] = totalEquity + bsRe[t];
      bsChk[t] = bsTast[t] - (bsTdbt[t] + bsTeq[t]);
    }

    // ── Returns ──────────────────────────────────────────────────────────────
    const exitYr = Math.min(5, Math.max(1, Math.round(V.exitYear)));
    const entryEquity = totalEquity;
    const exitEbitda = ebitda[exitYr];
    const exitEV = exitEbitda * V.exitMultiple;
    const exitDebt = totDebtEnd[exitYr];
    const exitCash = endCash[exitYr];
    const exitFee$ = exitEV * V.exitFee;
    const exitEquity = Math.max(exitEV - exitFee$ - exitDebt + exitCash, 0);
    const moic = entryEquity !== 0 ? exitEquity / entryEquity : 0;

    // IRR from the cash-flow series (−entry at t0, +exit at exit year)
    const cf = [-entryEquity];
    for (let t = 1; t <= exitYr; t++) cf[t] = t === exitYr ? exitEquity : 0;
    const irrVal = entryEquity > 0 && exitEquity > 0 ? irr(cf, 0.2) : 0;

    // Leverage / equity metrics
    const leverage = V.ebitdaLTM ? totalDebt / V.ebitdaLTM : 0;
    const equityPct = ev ? totalEquity / ev : 0;


    // ── Sensitivity tables ─────────────────────────────────────────────────────
  const entryMults = [8, 9, 10, 11, 12];
  const exitMults = [8, 9, 10, 11, 12, 13];
  const irrTable = [], moicTable = [];

  for (const em of entryMults) {
    const irrRow = [], moicRow = [];

  // Recalculate EV at each entry multiple
  const evSen = V.ebitdaLTM * em;

  // Keep debt structure consistent with main model
  const tlbSen = V.tlbPct * evSen;
  const ssnSen = V.ssnPct * evSen;
  const mezSen = V.mezPct * evSen;
  const totalDebtSen = tlbSen + ssnSen + mezSen;

  // IMPORTANT: include financing fees, just like main model
  const entryEqSen =
    evSen * (1 + V.transFees) +
    totalDebtSen * V.finFees -
    totalDebtSen;

  for (const xm of exitMults) {
    const exitEVSen = exitEbitda * xm;
    const exitFeeSen = exitEVSen * V.exitFee;

    // Use actual model debt/cash at exit year
    const exitEqSen = Math.max(
      exitEVSen - exitFeeSen - exitDebt + exitCash,
      0
    );

    const moicSen = entryEqSen !== 0 ? exitEqSen / entryEqSen : 0;
    const irrSen =
      entryEqSen > 0 && exitEqSen > 0
        ? Math.pow(moicSen, 1 / Math.max(exitYr, 1)) - 1
        : 0;

    irrRow.push(irrSen);
    moicRow.push(moicSen);
  }

  irrTable.push(irrRow);
  moicTable.push(moicRow);
}

    const bsBalanced = bsChk.every((c) => Math.abs(c) < 1e-6);

    return {
      inputs: V,
      transaction: { ev, totalDebt, tlb0, ssn0, mez0, leverage, equityPct },
      sources: { totalDebt, totalEquity, sponsorEquity, mgmtEquity, totalSources, totalUses, suCheck, transFee$, finFee$, dfc },
      years: {
        rev, ebitda, da, ebit, capex, nwc, margin,
        tlbEnd, ssnEnd, mezEnd, totDebtEnd, totCashInt, totPik, totInt, totMand,
        ofcf, avail, sweep, endCash, isNi, isEbt, bsChk, bsTast, bsTeq, bsTdbt,
      },
      returns: { entryEquity, exitYr, exitEbitda, exitEV, exitDebt, exitCash, exitFee$, exitEquity, moic, irr: irrVal },
      sensitivity: { entryMults, exitMults, irrTable, moicTable, baseEntry: 10, baseExit: 11 },
      bsBalanced, bsCheck: bsChk,
    };
  }

  return { computeLBO, irr, DEFAULTS };
});
