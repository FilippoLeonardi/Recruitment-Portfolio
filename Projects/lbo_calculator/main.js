// main.js — Node entry point. Builds lbo_model.xlsx from the command line.
// Cross-platform: output opens in Excel for Mac, Windows, or Microsoft 365 web.

const path = require("path");
const ExcelJS = require("exceljs");
const { buildWorkbook } = require("./buildModel");

async function main() {
  const filename = path.join(__dirname, "lbo_model.xlsx");
  console.log("Building all 7 sheets...");
  const wb = buildWorkbook(ExcelJS);          // pass {...inputs} to customise
  await wb.xlsx.writeFile(filename);
  console.log(`\n✅  Model complete — ${filename}`);
  console.log("   Open in Excel (Mac/Windows), or upload to OneDrive and open in Excel for the web.");
}

main().catch((err) => { console.error("\nERROR building model:", err); process.exit(1); });
