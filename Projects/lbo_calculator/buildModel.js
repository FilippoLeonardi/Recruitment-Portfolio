// buildModel.js — shared workbook factory (works in Node and the browser).
// Pass in the ExcelJS constructor so the same code runs in both environments:
//   Node:    buildWorkbook(require("exceljs"), inputs)
//   Browser: buildWorkbook(window.ExcelJS, inputs)

const { buildAssumptions } = require("./assumptions");
const { buildSourcesUses } = require("./sourcesUses");
const { buildIncomeStatement } = require("./incomeStmt");
const { buildDebtSchedule } = require("./debtSchedule");
const { buildCashFlow } = require("./cashFlow");
const { buildBalanceSheet } = require("./balanceSheet");
const { buildReturns } = require("./returns");
const { buildAllNamedRanges } = require("./constants");
const { registerNamedRanges } = require("./helpers");

function buildWorkbook(ExcelJS, inputs = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "LBO Model";
  wb.created = new Date();

  const wsAsm = wb.addWorksheet("Assumptions");
  const wsSu  = wb.addWorksheet("Sources_Uses");
  const wsIs  = wb.addWorksheet("Income_Stmt");
  const wsDs  = wb.addWorksheet("Debt_Schedule");
  const wsCf  = wb.addWorksheet("Cash_Flow");
  const wsBs  = wb.addWorksheet("Balance_Sheet");
  const wsRet = wb.addWorksheet("Returns");

  buildAssumptions(wsAsm, inputs);   // ← only sheet with inputs
  buildSourcesUses(wsSu);
  buildIncomeStatement(wsIs);
  buildDebtSchedule(wsDs);
  buildCashFlow(wsCf);
  buildBalanceSheet(wsBs);
  buildReturns(wsRet);

  registerNamedRanges(wb, buildAllNamedRanges());
  wb.calcProperties.fullCalcOnLoad = true;  // recalc on open, every platform
  return wb;
}

module.exports = { buildWorkbook };
