# LBO Model + Live Calculator (JavaScript)

A JavaScript port of the Python/openpyxl LBO model, in three usable pieces:

1. **Live web calculator** (`index.html`) — IRR / MOIC / sensitivity update
   instantly in the browser as you change assumptions.
2. **Excel export** — generate the full 7-sheet workbook (`.xlsx`) client-side,
   opens in Excel for Mac, Windows, or Microsoft 365 online.
3. **Node CLI** (`node main.js`) — write the workbook from the command line.

Model sheets: `Assumptions → Sources_Uses → Income_Stmt → Debt_Schedule → Cash_Flow → Balance_Sheet → Returns`.

---

## Quick start

Open `index.html` in a browser — that's the whole calculator, no build step.
It loads three scripts:

```html
<script src="calc.js"></script>            <!-- live math → window.LBOCalc -->
<script src="dist/exceljs.min.js"></script> <!-- ExcelJS → window.ExcelJS -->
<script src="dist/lbo-model.js"></script>   <!-- workbook builder → window.LBOModel -->
```

Host it on any static host (GitHub Pages, Netlify, Vercel, your existing site).
No backend, no server, no cold starts.

---

## Two engines, one set of inputs

The page has **two independent engines** that take the same `inputs` object:

| | `window.LBOCalc.computeLBO(inputs)` | `window.LBOModel.downloadModel(inputs)` |
|---|---|---|
| Runs | in the browser, instantly | in the browser, on click |
| Output | a results object (IRR, MOIC, per-year arrays, sensitivity grid) | a downloaded `.xlsx` |
| Computes | the math **in JavaScript** | writes **Excel formulas**; Excel computes on open |

`computeLBO` is what makes the numbers live on the page; `downloadModel` is the
Excel export. They're kept consistent — same defaults, same formulas.

### Minimal integration

```html
<script src="calc.js"></script>
<script src="dist/exceljs.min.js"></script>
<script src="dist/lbo-model.js"></script>

<script>
  const inputs = { entryMultiple:10, exitMultiple:11, exitYear:5,
                   revenueLTM:400, ebitdaLTM:100, growth:0.05, taxRate:0.25 };

  // live numbers on the page:
  const r = LBOCalc.computeLBO(inputs);
  console.log(r.returns.irr, r.returns.moic);   // 0.233…, 2.85…

  // export the workbook:
  LBOModel.downloadModel(inputs, "lbo_model.xlsx");
</script>
```

### What `computeLBO` returns

`returns` (entryEquity, exitEquity, exitEV, moic, irr, exitYr) ·
`transaction` (ev, totalDebt, leverage, equityPct) ·
`sources` (all Sources & Uses lines + suCheck) ·
`years` (per-year arrays: rev, ebitda, da, ebit, capex, nwc, totDebtEnd,
totCashInt, ofcf, endCash, isNi, bsChk, …) ·
`sensitivity` (entryMults, exitMults, irrTable, moicTable) ·
`bsBalanced` + `bsCheck` (the balance-sheet reconciliation per year).

---

## Node CLI

```bash
npm install
node main.js        # writes lbo_model.xlsx
```

Customise by passing inputs in `main.js`: `buildWorkbook(ExcelJS, { entryMultiple: 9, … })`.
Rebuild the browser bundle after editing any builder: `npm run build:browser`.

---

## Why the Excel file opens correctly everywhere

`.xlsx` is universal. Like openpyxl, ExcelJS writes formula *text* but no cached
results, so the workbook sets **`fullCalcOnLoad`** — Excel recomputes on open, on
Mac, Windows, or 365 web. (The Windows-only `os.startfile()` from the Python
version is gone.) To open online: upload to OneDrive → click → Excel for the web.

---

## One thing worth knowing: a balance-check edge case

The live calculator shows an honest **balance-sheet reconciliation** badge. In
most scenarios it reads "reconciles" (assets = liabilities + equity to the
penny). In scenarios where a debt tranche **fully repays before the exit year**,
you may see a small residual (e.g. −2.0).

Cause (present in the original Python model too): TLB mandatory amortisation is
`initial balance × 1%` **every year**, even after the balance hits zero. The cash
flow keeps subtracting it, but the balance can't go below zero, so assets fall by
that amount more than liabilities.

One-line fix, if you want it applied (cap amortisation at the outstanding
balance) — in both engines:

- `debtSchedule.js` (Excel): `MIN(ASM_TLB_AMT*ASM_TLB_CAMRT, TLB_beg)` for the
  TLB mandatory-amort formula.
- `calc.js` (live): `tlbAmrt[t] = Math.min(tlb0 * V.tlbAmrt, tlbBeg[t])`.

I left it faithful to your original model by default rather than changing its
behaviour silently. Say the word and I'll apply the fix to both.

---

## Files

| File | Purpose |
|------|---------|
| `index.html` | The live calculator page |
| `calc.js` | Live calculation engine (`window.LBOCalc`) |
| `dist/lbo-model.js` | Workbook builder, bundled (`window.LBOModel`) |
| `dist/exceljs.min.js` | ExcelJS browser build |
| `styles.js`, `helpers.js`, `constants.js` | Formatting, cell writers, row/col registry |
| `assumptions.js` … `returns.js` | One builder per sheet |
| `buildModel.js` | Assembles the workbook (Node + browser) |
| `browser-entry.js` | Browser API, bundled into `dist/lbo-model.js` |
| `main.js` | Node CLI entry |

Library: **ExcelJS 4.4.0** (MIT) for the workbook; the live engine is dependency-free.
