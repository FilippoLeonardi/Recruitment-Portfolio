var LBOModel = (() => {
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __commonJS = (cb, mod) => function __require() {
    try {
      return mod || (0, cb[__getOwnPropNames(cb)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
    } catch (e) {
      throw mod = 0, e;
    }
  };

  // styles.js
  var require_styles = __commonJS({
    "styles.js"(exports, module) {
      var DARK_BLUE = "1F4E79";
      var MID_BLUE = "2E75B6";
      var LIGHT_BLUE = "D6E4F0";
      var INPUT_YLW = "FFF2CC";
      var CALC_GREEN = "E2EFDA";
      var TOTAL_GREY = "D9D9D9";
      var WHITE = "FFFFFF";
      var BLACK = "000000";
      var ORANGE = "FCE4D6";
      var INPUT_BLUE = "0070C0";
      var FORMULA_BLACK = "000000";
      var LINK_GREEN = "00B050";
      var EXTERNAL_PURPLE = "7030A0";
      var INPUT_YELLOW = "FFF2CC";
      var FMT_DOLLAR = "#,##0.0";
      var FMT_PERCENT = "0.0%";
      var FMT_MULTIPLE = '0.0"x"';
      var FMT_INTEGER = "0";
      module.exports = {
        DARK_BLUE,
        MID_BLUE,
        LIGHT_BLUE,
        INPUT_YLW,
        CALC_GREEN,
        TOTAL_GREY,
        WHITE,
        BLACK,
        ORANGE,
        INPUT_BLUE,
        FORMULA_BLACK,
        LINK_GREEN,
        EXTERNAL_PURPLE,
        INPUT_YELLOW,
        FMT_DOLLAR,
        FMT_PERCENT,
        FMT_MULTIPLE,
        FMT_INTEGER,
      };
    }
  });

  // helpers.js
  var require_helpers = __commonJS({
    "helpers.js"(exports, module) {
      var S = require_styles();
      function colLetter(n) {
        let s = "";
        while (n > 0) {
          const m = (n - 1) % 26;
          s = String.fromCharCode(65 + m) + s;
          n = Math.floor((n - 1) / 26);
        }
        return s;
      }
      function crossRef(sheetName, row, col) {
        return `${sheetName}!$${colLetter(col)}$${row}`;
      }
      function sameRef(row, col) {
        return `$${colLetter(col)}$${row}`;
      }
      function crossRefRange(sheetName, row, colStart, colEnd) {
        return `'${sheetName}'!$${colLetter(colStart)}$${row}:$${colLetter(colEnd)}$${row}`;
      }
      function argb(hex) {
        if (!hex) return void 0;
        return hex.length === 6 ? "FF" + hex : hex;
      }
      function cellRC(ws, row, col) {
        return ws.getRow(row).getCell(col);
      }
      function applyStyle(cell, {
        bold = false,
        size = 10,
        colour = S.BLACK,
        bg = null,
        fmt = null,
        align = "left",
        italic = false
      } = {}) {
        cell.font = { name: "Calibri", bold, size, italic, color: { argb: argb(colour) } };
        cell.alignment = { horizontal: align, vertical: "middle" };
        if (bg) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(bg) } };
        if (fmt) cell.numFmt = fmt;
      }
      function writeLabel(ws, row, col, text, opts = {}) {
        const cell = cellRC(ws, row, col);
        cell.value = text;
        applyStyle(cell, opts);
        return cell;
      }
      function writeInput(ws, row, col, value, fmt = S.FMT_DOLLAR, opts = {}) {
        const { bold = false, bg = S.INPUT_YELLOW, colour = S.BLACK } = opts;
        const cell = cellRC(ws, row, col);
        cell.value = value;
        applyStyle(cell, { bold, colour, bg, fmt, align: "right" });
        return cell;
      }
      function writeFormula(ws, row, col, formula, fmt = S.FMT_DOLLAR, opts = {}) {
        const { bold = false, bg = null, colour = S.BLACK } = opts;
        const cell = cellRC(ws, row, col);
        cell.value = { formula: formula.replace(/^=/, "") };
        applyStyle(cell, { bold, colour, bg, fmt, align: "right" });
        return cell;
      }
      function writeHeader(ws, row, col, text, span = 7) {
        const cell = cellRC(ws, row, col);
        cell.value = text;
        applyStyle(cell, { bold: true, size: 11, colour: S.WHITE, bg: S.DARK_BLUE });
        if (span > 1) ws.mergeCells(row, col, row, col + span - 1);
      }
      function writeTotal(ws, row, col, formula, fmt = S.FMT_DOLLAR, opts = {}) {
        const { label = null, colour = S.BLACK } = opts;
        const border = {
          top: { style: "medium", color: { argb: "FF404040" } },
          bottom: { style: "medium", color: { argb: "FF404040" } }
        };
        const cell = cellRC(ws, row, col);
        cell.value = { formula: formula.replace(/^=/, "") };
        applyStyle(cell, { bold: true, colour, fmt, bg: S.TOTAL_GREY, align: "right" });
        cell.border = border;
        if (label) {
          const lbl = cellRC(ws, row, 2);
          lbl.value = label;
          applyStyle(lbl, { bold: true, colour, bg: S.TOTAL_GREY });
          lbl.border = border;
        }
      }
      function writeInfo(ws, row, col, formula, fmt = S.FMT_PERCENT) {
        const cell = cellRC(ws, row, col);
        cell.value = { formula: formula.replace(/^=/, "") };
        applyStyle(cell, { fmt, align: "right" });
        return cell;
      }
      function registerNamedRanges(wb, namedMap) {
        let n = 0;
        for (const [name, [sheet, row, col]] of Object.entries(namedMap)) {
          const addr = `'${sheet}'!$${colLetter(col)}$${row}`;
          wb.definedNames.add(addr, name);
          n++;
        }
        console.log(`  \u2192 ${n} named ranges registered`);
      }
      module.exports = {
        colLetter,
        crossRef,
        sameRef,
        crossRefRange,
        argb,
        cellRC,
        applyStyle,
        writeLabel,
        writeInput,
        writeFormula,
        writeHeader,
        writeTotal,
        writeInfo,
        registerNamedRanges
      };
    }
  });

  // constants.js
  var require_constants = __commonJS({
    "constants.js"(exports, module) {
      var SHEET_ASM = "Assumptions";
      var SHEET_SU = "Sources_Uses";
      var SHEET_IS = "Income_Stmt";
      var SHEET_DS = "Debt_Schedule";
      var SHEET_CF = "Cash_Flow";
      var SHEET_BS = "Balance_Sheet";
      var SHEET_RET = "Returns";
      var LTM_COL = 3;
      var YR1_COL = 4;
      var N_YEARS = 5;
      var ASM_INPUT_COL = 3;
      var ASM_DEBT_PCT_C = 3;
      var ASM_DEBT_AMT_C = 4;
      var ASM_RATE_C = 5;
      var ASM_CASH_AMRT_C = 6;
      var ASM_PIK_RATE_C = 7;
      var ASM_LTM_EBITDA_R = 5;
      var ASM_ENTRY_MULT_R = 6;
      var ASM_ENTRY_EV_R = 7;
      var ASM_TRANS_FEES_R = 8;
      var ASM_FIN_FEES_R = 9;
      var ASM_MGMT_ROLLOVER_R = 10;
      var ASM_REVOLVER_R = 14;
      var ASM_TLB_R = 15;
      var ASM_SSN_R = 16;
      var ASM_MEZ_R = 17;
      var ASM_TOT_DEBT_R = 19;
      var ASM_EQUITY_R = 20;
      var ASM_REVENUE_R = 27;
      var ASM_EBITDA_R = 28;
      var ASM_DA_R = 30;
      var ASM_EBIT_R = 31;
      var ASM_CAPEX_R = 33;
      var ASM_NWC_R = 34;
      var ASM_TAX_R = 35;
      var ASM_EBITDA_MRG_R = 38;
      var ASM_DA_PCT_R = 39;
      var ASM_EBIT_MRG_R = 40;
      var ASM_CAPEX_PCT_R = 41;
      var ASM_NWC_PCT_R = 42;
      var ASM_GROWTH_R = 45;
      var ASM_MARGIN_EXP_R = 46;
      var ASM_EXIT_YEAR_R = 50;
      var ASM_EXIT_MULT_R = 51;
      var ASM_EXIT_FEE_R = 52;
      var ASM_BS_HDR_R = 56;
      var ASM_INIT_PPE_R = 58;
      var ASM_INIT_NWC_BAL_R = 59;
      var SU_VAL_C = 3;
      var SU_PCT_C = 4;
      var SU_EV_PCT_C = 5;
      var SU_TLB_R = 6;
      var SU_SSN_R = 7;
      var SU_MEZ_R = 8;
      var SU_TOT_DEBT_R = 9;
      var SU_SPONS_EQ_R = 11;
      var SU_MGMT_EQ_R = 12;
      var SU_TOT_EQ_R = 13;
      var SU_TOT_SRC_R = 14;
      var SU_PRICE_R = 18;
      var SU_TRANS_FEE_R = 19;
      var SU_FIN_FEE_R = 20;
      var SU_TOT_USE_R = 21;
      var SU_CHECK_R = 23;
      var IS_REV_R = 6;
      var IS_GRW_R = 7;
      var IS_EBD_R = 9;
      var IS_MRG_R = 10;
      var IS_DA_R = 11;
      var IS_EBI_R = 12;
      var IS_EBI_MRG_R = 13;
      var IS_INT_TLB_R = 15;
      var IS_INT_SSN_R = 16;
      var IS_INT_MEZ_R = 17;
      var IS_INT_TOT_R = 18;
      var IS_EBT_R = 20;
      var IS_TAX_R = 21;
      var IS_NI_R = 22;
      var IS_NI_MRG_R = 23;
      var DS_YR0_C = 3;
      var DS_TLB_HDR_R = 6;
      var DS_TLB_BEG_R = 7;
      var DS_TLB_AMRT_R = 8;
      var DS_TLB_OPT_R = 9;
      var DS_TLB_END_R = 10;
      var DS_TLB_INT_R = 12;
      var DS_SSN_HDR_R = 14;
      var DS_SSN_BEG_R = 15;
      var DS_SSN_AMRT_R = 16;
      var DS_SSN_OPT_R = 17;
      var DS_SSN_END_R = 18;
      var DS_SSN_INT_R = 20;
      var DS_MEZ_HDR_R = 22;
      var DS_MEZ_BEG_R = 23;
      var DS_MEZ_PIK_R = 24;
      var DS_MEZ_OPT_R = 25;
      var DS_MEZ_END_R = 26;
      var DS_MEZ_CASH_INT_R = 28;
      var DS_MEZ_PIK_INT_R = 29;
      var DS_SUM_HDR_R = 31;
      var DS_TDEBT_R = 33;
      var DS_TCASH_INT_R = 34;
      var DS_TPIK_R = 35;
      var DS_TINT_R = 36;
      var DS_TMAND_R = 37;
      var CF_EBITDA_R = 6;
      var CF_INT_R = 7;
      var CF_TAX_R = 8;
      var CF_CAPEX_R = 9;
      var CF_NWC_R = 10;
      var CF_OFCF_R = 11;
      var CF_MAND_R = 12;
      var CF_AVAIL_R = 13;
      var CF_SWEEP_R = 14;
      var CF_NCF_R = 15;
      var CF_BEG_CASH_R = 16;
      var CF_END_CASH_R = 17;
      var BS_YR0_C = 3;
      var BS_AST_HDR_R = 5;
      var BS_CASH_R = 6;
      var BS_NWC_R = 7;
      var BS_PPE_R = 8;
      var BS_GW_R = 9;
      var BS_DFC_R = 10;
      var BS_TAST_R = 11;
      var BS_LBL_HDR_R = 13;
      var BS_TLB_R = 14;
      var BS_SSN_R = 15;
      var BS_MEZ_R = 16;
      var BS_TDBT_R = 17;
      var BS_EQ_HDR_R = 19;
      var BS_INIT_EQ_R = 20;
      var BS_RE_R = 21;
      var BS_TEQ_R = 22;
      var BS_TLE_R = 24;
      var BS_CHK_R = 25;
      var RET_VAL_C = 3;
      var RET_HDR_R = 4;
      var RET_ENTRY_EQ_R = 5;
      var RET_EXIT_YR_R = 6;
      var RET_EXIT_EBD_R = 7;
      var RET_EXIT_MULT_R = 8;
      var RET_EXIT_EV_R = 9;
      var RET_EXIT_DBT_R = 10;
      var RET_EXIT_CSH_R = 11;
      var RET_EXIT_FEE_R = 12;
      var RET_EXIT_EQ_R = 13;
      var RET_MOIC_R = 14;
      var RET_IRR_R = 15;
      var RET_CF_HDR_R = 17;
      var RET_CF0_R = 18;
      var RET_CF1_R = 19;
      var RET_CF2_R = 20;
      var RET_CF3_R = 21;
      var RET_CF4_R = 22;
      var RET_CF5_R = 23;
      var SEN_IRR_HDR_R = 27;
      var SEN_IRR_R1 = 30;
      var SEN_IRR_C1 = 4;
      var SEN_MOIC_HDR_R = 40;
      var SEN_MOIC_R1 = 43;
      var SEN_MOIC_C1 = 4;
      var SEN_ENTRY_MULTS = [8, 9, 10, 11, 12];
      var SEN_EXIT_MULTS = [8, 9, 10, 11, 12, 13];
      var NAMED_RANGES = {
        ASM_LTM_EBITDA: [SHEET_ASM, ASM_LTM_EBITDA_R, ASM_INPUT_COL],
        ASM_ENTRY_MULT: [SHEET_ASM, ASM_ENTRY_MULT_R, ASM_INPUT_COL],
        ASM_ENTRY_EV: [SHEET_ASM, ASM_ENTRY_EV_R, ASM_INPUT_COL],
        ASM_TRANS_FEES: [SHEET_ASM, ASM_TRANS_FEES_R, ASM_INPUT_COL],
        ASM_FIN_FEES: [SHEET_ASM, ASM_FIN_FEES_R, ASM_INPUT_COL],
        ASM_MGMT_ROLLOVER: [SHEET_ASM, ASM_MGMT_ROLLOVER_R, ASM_INPUT_COL],
        ASM_TLB_AMT: [SHEET_ASM, ASM_TLB_R, ASM_DEBT_AMT_C],
        ASM_TLB_RATE: [SHEET_ASM, ASM_TLB_R, ASM_RATE_C],
        ASM_TLB_CAMRT: [SHEET_ASM, ASM_TLB_R, ASM_CASH_AMRT_C],
        ASM_SSN_AMT: [SHEET_ASM, ASM_SSN_R, ASM_DEBT_AMT_C],
        ASM_SSN_RATE: [SHEET_ASM, ASM_SSN_R, ASM_RATE_C],
        ASM_MEZ_AMT: [SHEET_ASM, ASM_MEZ_R, ASM_DEBT_AMT_C],
        ASM_MEZ_RATE: [SHEET_ASM, ASM_MEZ_R, ASM_RATE_C],
        ASM_MEZ_PIK: [SHEET_ASM, ASM_MEZ_R, ASM_PIK_RATE_C],
        ASM_TOT_DEBT: [SHEET_ASM, ASM_TOT_DEBT_R, ASM_DEBT_AMT_C],
        ASM_EQUITY: [SHEET_ASM, ASM_EQUITY_R, ASM_DEBT_AMT_C],
        ASM_REV_LTM: [SHEET_ASM, ASM_REVENUE_R, LTM_COL],
        ASM_EBITDA_LTM: [SHEET_ASM, ASM_EBITDA_R, LTM_COL],
        ASM_DA_LTM: [SHEET_ASM, ASM_DA_R, LTM_COL],
        ASM_EBIT_LTM: [SHEET_ASM, ASM_EBIT_R, LTM_COL],
        ASM_CAPEX_LTM: [SHEET_ASM, ASM_CAPEX_R, LTM_COL],
        ASM_NWC_LTM: [SHEET_ASM, ASM_NWC_R, LTM_COL],
        ASM_EXIT_YEAR: [SHEET_ASM, ASM_EXIT_YEAR_R, ASM_INPUT_COL],
        ASM_EXIT_MULT: [SHEET_ASM, ASM_EXIT_MULT_R, ASM_INPUT_COL],
        ASM_EXIT_FEE: [SHEET_ASM, ASM_EXIT_FEE_R, ASM_INPUT_COL]
      };
      var YEAR_RANGE_ROWS = [
        ["ASM_REV", SHEET_ASM, ASM_REVENUE_R],
        ["ASM_EBITDA", SHEET_ASM, ASM_EBITDA_R],
        ["ASM_DA", SHEET_ASM, ASM_DA_R],
        ["ASM_EBIT", SHEET_ASM, ASM_EBIT_R],
        ["ASM_CAPEX", SHEET_ASM, ASM_CAPEX_R],
        ["ASM_NWC", SHEET_ASM, ASM_NWC_R],
        ["ASM_TAX", SHEET_ASM, ASM_TAX_R],
        ["ASM_MARGIN", SHEET_ASM, ASM_EBITDA_MRG_R]
      ];
      function buildAllNamedRanges() {
        const all = { ...NAMED_RANGES };
        for (const [base, sheet, row] of YEAR_RANGE_ROWS) {
          for (let yr = 1; yr <= N_YEARS; yr++) {
            all[`${base}_Y${yr}`] = [sheet, row, YR1_COL + yr - 1];
          }
        }
        return all;
      }
      module.exports = {
        SHEET_ASM,
        SHEET_SU,
        SHEET_IS,
        SHEET_DS,
        SHEET_CF,
        SHEET_BS,
        SHEET_RET,
        LTM_COL,
        YR1_COL,
        N_YEARS,
        ASM_INPUT_COL,
        ASM_DEBT_PCT_C,
        ASM_DEBT_AMT_C,
        ASM_RATE_C,
        ASM_CASH_AMRT_C,
        ASM_PIK_RATE_C,
        ASM_LTM_EBITDA_R,
        ASM_ENTRY_MULT_R,
        ASM_ENTRY_EV_R,
        ASM_TRANS_FEES_R,
        ASM_FIN_FEES_R,
        ASM_MGMT_ROLLOVER_R,
        ASM_REVOLVER_R,
        ASM_TLB_R,
        ASM_SSN_R,
        ASM_MEZ_R,
        ASM_TOT_DEBT_R,
        ASM_EQUITY_R,
        ASM_REVENUE_R,
        ASM_EBITDA_R,
        ASM_DA_R,
        ASM_EBIT_R,
        ASM_CAPEX_R,
        ASM_NWC_R,
        ASM_TAX_R,
        ASM_EBITDA_MRG_R,
        ASM_DA_PCT_R,
        ASM_EBIT_MRG_R,
        ASM_CAPEX_PCT_R,
        ASM_NWC_PCT_R,
        ASM_GROWTH_R,
        ASM_MARGIN_EXP_R,
        ASM_EXIT_YEAR_R,
        ASM_EXIT_MULT_R,
        ASM_EXIT_FEE_R,
        ASM_BS_HDR_R,
        ASM_INIT_PPE_R,
        ASM_INIT_NWC_BAL_R,
        SU_VAL_C,
        SU_PCT_C,
        SU_EV_PCT_C,
        SU_TLB_R,
        SU_SSN_R,
        SU_MEZ_R,
        SU_TOT_DEBT_R,
        SU_SPONS_EQ_R,
        SU_MGMT_EQ_R,
        SU_TOT_EQ_R,
        SU_TOT_SRC_R,
        SU_PRICE_R,
        SU_TRANS_FEE_R,
        SU_FIN_FEE_R,
        SU_TOT_USE_R,
        SU_CHECK_R,
        IS_REV_R,
        IS_GRW_R,
        IS_EBD_R,
        IS_MRG_R,
        IS_DA_R,
        IS_EBI_R,
        IS_EBI_MRG_R,
        IS_INT_TLB_R,
        IS_INT_SSN_R,
        IS_INT_MEZ_R,
        IS_INT_TOT_R,
        IS_EBT_R,
        IS_TAX_R,
        IS_NI_R,
        IS_NI_MRG_R,
        DS_YR0_C,
        DS_TLB_HDR_R,
        DS_TLB_BEG_R,
        DS_TLB_AMRT_R,
        DS_TLB_OPT_R,
        DS_TLB_END_R,
        DS_TLB_INT_R,
        DS_SSN_HDR_R,
        DS_SSN_BEG_R,
        DS_SSN_AMRT_R,
        DS_SSN_OPT_R,
        DS_SSN_END_R,
        DS_SSN_INT_R,
        DS_MEZ_HDR_R,
        DS_MEZ_BEG_R,
        DS_MEZ_PIK_R,
        DS_MEZ_OPT_R,
        DS_MEZ_END_R,
        DS_MEZ_CASH_INT_R,
        DS_MEZ_PIK_INT_R,
        DS_SUM_HDR_R,
        DS_TDEBT_R,
        DS_TCASH_INT_R,
        DS_TPIK_R,
        DS_TINT_R,
        DS_TMAND_R,
        CF_EBITDA_R,
        CF_INT_R,
        CF_TAX_R,
        CF_CAPEX_R,
        CF_NWC_R,
        CF_OFCF_R,
        CF_MAND_R,
        CF_AVAIL_R,
        CF_SWEEP_R,
        CF_NCF_R,
        CF_BEG_CASH_R,
        CF_END_CASH_R,
        BS_YR0_C,
        BS_AST_HDR_R,
        BS_CASH_R,
        BS_NWC_R,
        BS_PPE_R,
        BS_GW_R,
        BS_DFC_R,
        BS_TAST_R,
        BS_LBL_HDR_R,
        BS_TLB_R,
        BS_SSN_R,
        BS_MEZ_R,
        BS_TDBT_R,
        BS_EQ_HDR_R,
        BS_INIT_EQ_R,
        BS_RE_R,
        BS_TEQ_R,
        BS_TLE_R,
        BS_CHK_R,
        RET_VAL_C,
        RET_HDR_R,
        RET_ENTRY_EQ_R,
        RET_EXIT_YR_R,
        RET_EXIT_EBD_R,
        RET_EXIT_MULT_R,
        RET_EXIT_EV_R,
        RET_EXIT_DBT_R,
        RET_EXIT_CSH_R,
        RET_EXIT_FEE_R,
        RET_EXIT_EQ_R,
        RET_MOIC_R,
        RET_IRR_R,
        RET_CF_HDR_R,
        RET_CF0_R,
        RET_CF1_R,
        RET_CF2_R,
        RET_CF3_R,
        RET_CF4_R,
        RET_CF5_R,
        SEN_IRR_HDR_R,
        SEN_IRR_R1,
        SEN_IRR_C1,
        SEN_MOIC_HDR_R,
        SEN_MOIC_R1,
        SEN_MOIC_C1,
        SEN_ENTRY_MULTS,
        SEN_EXIT_MULTS,
        NAMED_RANGES,
        YEAR_RANGE_ROWS,
        buildAllNamedRanges
      };
    }
  });

  // assumptions.js
  var require_assumptions = __commonJS({
    "assumptions.js"(exports, module) {
      var H = require_helpers();
      var S = require_styles();
      var C = require_constants();
      var { sameRef, writeLabel, writeInput, writeFormula, writeHeader, writeTotal } = H;
      var {
        INPUT_BLUE,
        FORMULA_BLACK,
        INPUT_YELLOW,
        DARK_BLUE,
        LIGHT_BLUE,
        FMT_DOLLAR,
        FMT_PERCENT,
        FMT_MULTIPLE,
        FMT_INTEGER
      } = S;
      var INPUT_COL = 3;
      var ROW_LTM_EBITDA = 5;
      var ROW_ENTRY_MULT = 6;
      var ROW_ENTRY_EV = 7;
      var ROW_TRANS_FEES = 8;
      var ROW_FIN_FEES = 9;
      var ROW_MGMT_ROLLOVER = 10;
      var ROW_REVOLVER = 14;
      var ROW_TLB = 15;
      var ROW_SSN = 16;
      var ROW_MEZ = 17;
      var ROW_TOT_DEBT = 19;
      var ROW_EQUITY = 20;
      var COL_DEBT_PCT = 3;
      var COL_DEBT_AMT = 4;
      var COL_RATE = 5;
      var COL_CASH_AMRT = 6;
      var COL_PIK_RATE = 7;
      var ROW_REVENUE = 27;
      var ROW_EBITDA = 28;
      var ROW_DA = 30;
      var ROW_EBIT = 31;
      var ROW_CAPEX = 33;
      var ROW_NWC = 34;
      var ROW_TAX_RATE = 35;
      var ROW_EBITDA_MRG = 38;
      var ROW_DA_PCT = 39;
      var ROW_EBIT_MRG = 40;
      var ROW_CAPEX_PCT = 41;
      var ROW_NWC_PCT = 42;
      var ROW_GROWTH = 45;
      var ROW_MARGIN_EXP = 46;
      var ROW_EXIT_YEAR = 50;
      var ROW_EXIT_MULT = 51;
      var ROW_EXIT_FEE = 52;
      var DEFAULTS = {
        entryMultiple: 10,
        transFees: 0.02,
        finFees: 0.015,
        mgmtRollover: 0.1,
        revolverPct: 0.1,
        revolverRate: 0.045,
        revolverAmrt: 0,
        tlbPct: 0.35,
        tlbRate: 0.055,
        tlbAmrt: 0.01,
        ssnPct: 0.2,
        ssnRate: 0.065,
        ssnAmrt: 0,
        mezPct: 0.1,
        mezRate: 0.09,
        mezAmrt: 0,
        mezPik: 0.02,
        revenueLTM: 400,
        ebitdaLTM: 100,
        daLTM: 20,
        capexLTM: 16,
        nwcLTM: 8,
        taxRate: 0.25,
        growth: [0.05, 0.05, 0.05, 0.05, 0.05],
        marginExp: 5e-3,
        exitYear: 5,
        exitMultiple: 11,
        exitFee: 0.015,
        openingPPE: 80,
        openingNWC: 60
      };
      function buildAssumptions(ws, inputs = {}) {
        const V = { ...DEFAULTS, ...inputs };
        const g = Array.isArray(V.growth) ? V.growth : [V.growth, V.growth, V.growth, V.growth, V.growth];
        ws.getColumn(1).width = 2;
        ws.getColumn(2).width = 45;
        ws.getColumn(3).width = 16;
        [4, 5, 6, 7, 8].forEach((n) => ws.getColumn(n).width = 14);
        writeLabel(ws, 1, 2, "LBO MODEL \u2014 ASSUMPTIONS", { bold: true, size: 14, colour: DARK_BLUE });
        writeLabel(
          ws,
          2,
          2,
          "Yellow + Blue font = hard-coded input (you must fill these)   Black = same-sheet formula   Green = linked from another sheet   Purple = external file link",
          { italic: true, size: 9, colour: "808080" }
        );
        writeLabel(ws, ROW_LTM_EBITDA, 2, "LTM EBITDA ($M)  [links to Block 3 \u2014 enter value there]", { bold: true });
        writeFormula(
          ws,
          ROW_LTM_EBITDA,
          INPUT_COL,
          `=${sameRef(ROW_EBITDA, INPUT_COL)}`,
          FMT_DOLLAR,
          { bold: true, colour: FORMULA_BLACK }
        );
        writeLabel(ws, ROW_ENTRY_MULT, 2, "Entry EV / EBITDA Multiple (x)");
        writeInput(ws, ROW_ENTRY_MULT, INPUT_COL, V.entryMultiple, FMT_MULTIPLE, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(ws, ROW_ENTRY_EV, 2, "Entry Enterprise Value ($M)  [= LTM EBITDA \xD7 Multiple]", { bold: true });
        writeFormula(
          ws,
          ROW_ENTRY_EV,
          INPUT_COL,
          `=${sameRef(ROW_LTM_EBITDA, INPUT_COL)}*${sameRef(ROW_ENTRY_MULT, INPUT_COL)}`,
          FMT_DOLLAR,
          { bold: true, colour: FORMULA_BLACK }
        );
        writeLabel(ws, ROW_TRANS_FEES, 2, "Transaction Fees (% of EV)");
        writeInput(ws, ROW_TRANS_FEES, INPUT_COL, V.transFees, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(ws, ROW_FIN_FEES, 2, "Financing Fees (% of Total Debt)");
        writeInput(ws, ROW_FIN_FEES, INPUT_COL, V.finFees, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(ws, ROW_MGMT_ROLLOVER, 2, "Management Rollover (% of Equity)");
        writeInput(ws, ROW_MGMT_ROLLOVER, INPUT_COL, V.mgmtRollover, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeHeader(ws, 12, 2, "2.  DEBT STRUCTURE", 6);
        for (const [colN, label] of [
          [2, "Tranche"],
          [COL_DEBT_PCT, "% of EV"],
          [COL_DEBT_AMT, "$ Amount ($M)"],
          [COL_RATE, "Interest Rate"],
          [COL_CASH_AMRT, "Cash Amort %/yr"],
          [COL_PIK_RATE, "PIK Rate"]
        ]) {
          writeLabel(ws, 13, colN, label, { bold: true, bg: LIGHT_BLUE, align: "center" });
        }
        const debtRow = (row, label, pct, rate, cashAmrt, pik = 0) => {
          writeLabel(ws, row, 2, label);
          writeInput(ws, row, COL_DEBT_PCT, pct, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
          writeFormula(
            ws,
            row,
            COL_DEBT_AMT,
            `=${sameRef(row, COL_DEBT_PCT)}*${sameRef(ROW_ENTRY_EV, INPUT_COL)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
          writeInput(ws, row, COL_RATE, rate, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
          writeInput(ws, row, COL_CASH_AMRT, cashAmrt, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
          writeInput(ws, row, COL_PIK_RATE, pik, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        };
        debtRow(ROW_REVOLVER, "Revolving Credit (capacity, undrawn)", V.revolverPct, V.revolverRate, V.revolverAmrt);
        debtRow(ROW_TLB, "Term Loan B", V.tlbPct, V.tlbRate, V.tlbAmrt);
        debtRow(ROW_SSN, "Senior Secured Notes", V.ssnPct, V.ssnRate, V.ssnAmrt);
        debtRow(ROW_MEZ, "Mezzanine", V.mezPct, V.mezRate, V.mezAmrt, V.mezPik);
        writeTotal(
          ws,
          ROW_TOT_DEBT,
          COL_DEBT_AMT,
          `=SUM(${sameRef(ROW_TLB, COL_DEBT_AMT)}:${sameRef(ROW_MEZ, COL_DEBT_AMT)})`,
          FMT_DOLLAR,
          { label: "Total Drawn Debt", colour: FORMULA_BLACK }
        );
        writeLabel(ws, ROW_EQUITY, 2, "Sponsor Equity (plug)");
        writeFormula(
          ws,
          ROW_EQUITY,
          COL_DEBT_AMT,
          `=${sameRef(ROW_ENTRY_EV, INPUT_COL)}*(1+${sameRef(ROW_TRANS_FEES, INPUT_COL)})-${sameRef(ROW_TOT_DEBT, COL_DEBT_AMT)}`,
          FMT_DOLLAR,
          { bold: true, colour: FORMULA_BLACK }
        );
        writeHeader(ws, 23, 2, "3.  OPERATING ASSUMPTIONS", 7);
        writeLabel(
          ws,
          24,
          2,
          "SECTION A  \u2014  LTM actuals in $M (yellow = enter here)  +  Year 1-5 auto-projected",
          { italic: true, size: 9, colour: "1F4E79" }
        );
        writeLabel(ws, 26, INPUT_COL, "LTM", { bold: true, bg: LIGHT_BLUE, align: "center" });
        for (let i = 1; i <= 5; i++)
          writeLabel(ws, 26, INPUT_COL + i, `Year ${i}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeLabel(ws, ROW_REVENUE, 2, "Revenue  ($M)");
        writeInput(ws, ROW_REVENUE, INPUT_COL, V.revenueLTM, FMT_DOLLAR, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        for (let col = 4; col <= 8; col++)
          writeFormula(
            ws,
            ROW_REVENUE,
            col,
            `=${sameRef(ROW_REVENUE, col - 1)}*(1+${sameRef(ROW_GROWTH, col)})`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_EBITDA, 2, "EBITDA  ($M)");
        writeInput(ws, ROW_EBITDA, INPUT_COL, V.ebitdaLTM, FMT_DOLLAR, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        for (let col = 4; col <= 8; col++)
          writeFormula(
            ws,
            ROW_EBITDA,
            col,
            `=${sameRef(ROW_REVENUE, col)}*${sameRef(ROW_EBITDA_MRG, col)}`,
            FMT_DOLLAR,
            { bold: true, colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_DA, 2, "D&A  ($M)");
        writeInput(ws, ROW_DA, INPUT_COL, V.daLTM, FMT_DOLLAR, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        for (let col = 4; col <= 8; col++)
          writeFormula(
            ws,
            ROW_DA,
            col,
            `=${sameRef(ROW_REVENUE, col)}*${sameRef(ROW_DA_PCT, INPUT_COL)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_EBIT, 2, "EBIT  ($M)  [= EBITDA \u2212 D&A]", { bold: true });
        for (let col = 3; col <= 8; col++)
          writeFormula(
            ws,
            ROW_EBIT,
            col,
            `=${sameRef(ROW_EBITDA, col)}-${sameRef(ROW_DA, col)}`,
            FMT_DOLLAR,
            { bold: true, colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_CAPEX, 2, "CapEx  ($M)");
        writeInput(ws, ROW_CAPEX, INPUT_COL, V.capexLTM, FMT_DOLLAR, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        for (let col = 4; col <= 8; col++)
          writeFormula(
            ws,
            ROW_CAPEX,
            col,
            `=${sameRef(ROW_REVENUE, col)}*${sameRef(ROW_CAPEX_PCT, INPUT_COL)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_NWC, 2, "Change in NWC  ($M)");
        writeInput(ws, ROW_NWC, INPUT_COL, V.nwcLTM, FMT_DOLLAR, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        for (let col = 4; col <= 8; col++)
          writeFormula(
            ws,
            ROW_NWC,
            col,
            `=${sameRef(ROW_REVENUE, col)}*${sameRef(ROW_NWC_PCT, INPUT_COL)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_TAX_RATE, 2, "Tax Rate  (%)  [external: set by legislation \u2014 must be input]");
        for (let col = 3; col <= 8; col++)
          writeInput(ws, ROW_TAX_RATE, col, V.taxRate, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(
          ws,
          37,
          2,
          "SECTION B  \u2014  Derived % ratios (black = same-sheet formula)",
          { italic: true, size: 9, colour: "CC0000" }
        );
        writeLabel(ws, 37, INPUT_COL, "LTM", { bold: true, bg: LIGHT_BLUE, align: "center" });
        for (let i = 1; i <= 5; i++)
          writeLabel(ws, 37, INPUT_COL + i, `Year ${i}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeLabel(ws, ROW_EBITDA_MRG, 2, "EBITDA Margin  (%)  [LTM = actual; Y1-5 = LTM + cumulative expansion]");
        writeFormula(
          ws,
          ROW_EBITDA_MRG,
          INPUT_COL,
          `=IFERROR(${sameRef(ROW_EBITDA, INPUT_COL)}/${sameRef(ROW_REVENUE, INPUT_COL)},0)`,
          FMT_PERCENT,
          { colour: FORMULA_BLACK }
        );
        writeFormula(
          ws,
          ROW_EBITDA_MRG,
          4,
          `=${sameRef(ROW_EBITDA_MRG, INPUT_COL)}+${sameRef(ROW_MARGIN_EXP, INPUT_COL)}`,
          FMT_PERCENT,
          { colour: FORMULA_BLACK }
        );
        for (let col = 5; col <= 8; col++)
          writeFormula(
            ws,
            ROW_EBITDA_MRG,
            col,
            `=${sameRef(ROW_EBITDA_MRG, col - 1)}+${sameRef(ROW_MARGIN_EXP, INPUT_COL)}`,
            FMT_PERCENT,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_DA_PCT, 2, "D&A  (% of Revenue)  [constant = LTM ratio applied to all years]");
        for (let col = 3; col <= 8; col++)
          writeFormula(
            ws,
            ROW_DA_PCT,
            col,
            `=IFERROR(${sameRef(ROW_DA, col)}/${sameRef(ROW_REVENUE, col)},0)`,
            FMT_PERCENT,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_EBIT_MRG, 2, "EBIT Margin  (%)");
        for (let col = 3; col <= 8; col++)
          writeFormula(
            ws,
            ROW_EBIT_MRG,
            col,
            `=IFERROR(${sameRef(ROW_EBIT, col)}/${sameRef(ROW_REVENUE, col)},0)`,
            FMT_PERCENT,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_CAPEX_PCT, 2, "CapEx  (% of Revenue)  [constant = LTM ratio]");
        for (let col = 3; col <= 8; col++)
          writeFormula(
            ws,
            ROW_CAPEX_PCT,
            col,
            `=IFERROR(${sameRef(ROW_CAPEX, col)}/${sameRef(ROW_REVENUE, col)},0)`,
            FMT_PERCENT,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, ROW_NWC_PCT, 2, "NWC Change  (% of Revenue)  [constant = LTM ratio]");
        for (let col = 3; col <= 8; col++)
          writeFormula(
            ws,
            ROW_NWC_PCT,
            col,
            `=IFERROR(${sameRef(ROW_NWC, col)}/${sameRef(ROW_REVENUE, col)},0)`,
            FMT_PERCENT,
            { colour: FORMULA_BLACK }
          );
        writeLabel(
          ws,
          44,
          2,
          "SECTION C  \u2014  Forward growth drivers (yellow cells drive all of Section A Years 1-5)",
          { italic: true, size: 9, colour: "1F4E79" }
        );
        writeLabel(ws, 44, INPUT_COL, "LTM", { bold: true, bg: LIGHT_BLUE, align: "center" });
        for (let i = 1; i <= 5; i++)
          writeLabel(ws, 44, INPUT_COL + i, `Year ${i}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeLabel(ws, ROW_GROWTH, 2, "Revenue Growth Rate  (%/year)");
        writeLabel(ws, ROW_GROWTH, INPUT_COL, "n/a", { italic: true, size: 9, colour: "909090", align: "center" });
        for (let col = 4; col <= 8; col++)
          writeInput(ws, ROW_GROWTH, col, g[col - 4], FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(
          ws,
          ROW_MARGIN_EXP,
          2,
          "Annual EBITDA Margin Improvement  (%/yr)  [0 = flat | 0.005 = +50bps/yr | drives row 38]"
        );
        writeInput(ws, ROW_MARGIN_EXP, INPUT_COL, V.marginExp, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeHeader(ws, 48, 2, "4.  EXIT ASSUMPTIONS", 4);
        writeLabel(ws, ROW_EXIT_YEAR, 2, "Exit Year (1-5)");
        writeInput(ws, ROW_EXIT_YEAR, INPUT_COL, V.exitYear, FMT_INTEGER, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(ws, ROW_EXIT_MULT, 2, "Exit EV / EBITDA Multiple (x)");
        writeInput(ws, ROW_EXIT_MULT, INPUT_COL, V.exitMultiple, FMT_MULTIPLE, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(ws, ROW_EXIT_FEE, 2, "Exit Fee (% of Exit EV)");
        writeInput(ws, ROW_EXIT_FEE, INPUT_COL, V.exitFee, FMT_PERCENT, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeHeader(ws, C.ASM_BS_HDR_R, 2, "5.  BALANCE SHEET OPENING BALANCES AT CLOSE", 5);
        writeLabel(
          ws,
          57,
          2,
          "These are the target company's asset balances on Day 1 of ownership",
          { italic: true, size: 9, colour: "606060" }
        );
        writeLabel(ws, C.ASM_INIT_PPE_R, 2, "Opening PP&E Balance ($M)  [default = D&A \xD7 4]");
        writeInput(ws, C.ASM_INIT_PPE_R, INPUT_COL, V.openingPPE, FMT_DOLLAR, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
        writeLabel(ws, C.ASM_INIT_NWC_BAL_R, 2, "Opening NWC Balance ($M)  [default = Revenue \xD7 15%]");
        writeInput(ws, C.ASM_INIT_NWC_BAL_R, INPUT_COL, V.openingNWC, FMT_DOLLAR, { colour: INPUT_BLUE, bg: INPUT_YELLOW });
      }
      module.exports = { buildAssumptions, DEFAULTS };
    }
  });

  // sourcesUses.js
  var require_sourcesUses = __commonJS({
    "sourcesUses.js"(exports, module) {
      var H = require_helpers();
      var S = require_styles();
      var C = require_constants();
      var { crossRef, sameRef, writeLabel, writeFormula, writeHeader, writeTotal, applyStyle, cellRC } = H;
      var {
        FORMULA_BLACK,
        LINK_GREEN,
        DARK_BLUE,
        LIGHT_BLUE,
        TOTAL_GREY,
        FMT_DOLLAR,
        FMT_PERCENT,
        FMT_MULTIPLE
      } = S;
      var {
        SHEET_ASM,
        ASM_TLB_R,
        ASM_SSN_R,
        ASM_MEZ_R,
        ASM_DEBT_AMT_C,
        ASM_RATE_C,
        ASM_PIK_RATE_C,
        ASM_ENTRY_EV_R,
        ASM_TRANS_FEES_R,
        ASM_FIN_FEES_R,
        ASM_MGMT_ROLLOVER_R,
        ASM_INPUT_COL,
        ASM_LTM_EBITDA_R,
        ASM_EBITDA_R,
        SU_VAL_C,
        SU_PCT_C,
        SU_EV_PCT_C,
        SU_TLB_R,
        SU_SSN_R,
        SU_MEZ_R,
        SU_TOT_DEBT_R,
        SU_SPONS_EQ_R,
        SU_MGMT_EQ_R,
        SU_TOT_EQ_R,
        SU_TOT_SRC_R,
        SU_PRICE_R,
        SU_TRANS_FEE_R,
        SU_FIN_FEE_R,
        SU_TOT_USE_R,
        SU_CHECK_R
      } = C;
      var TOTAL_BORDER = {
        top: { style: "medium", color: { argb: "FF404040" } },
        bottom: { style: "medium", color: { argb: "FF404040" } }
      };
      function buildSourcesUses(ws) {
        ws.getColumn(1).width = 2;
        ws.getColumn(2).width = 45;
        ws.getColumn(3).width = 16;
        ws.getColumn(4).width = 14;
        ws.getColumn(5).width = 14;
        writeLabel(ws, 1, 2, "SOURCES & USES", { bold: true, size: 14, colour: DARK_BLUE });
        writeLabel(
          ws,
          2,
          2,
          "Transaction financing at close  |  Sources must equal Uses  |  ",
          { italic: true, size: 9, colour: "808080" }
        );
        for (const [col, label] of [[SU_VAL_C, "Amount  ($M)"], [SU_PCT_C, "% of Total"], [SU_EV_PCT_C, "% of EV"]])
          writeLabel(ws, 3, col, label, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeHeader(ws, 4, 2, "SOURCES  \u2014  Where the money comes from", 4);
        writeLabel(ws, 5, 2, "Debt", { bold: true, italic: true, colour: "404040" });
        writeLabel(ws, SU_TLB_R, 2, "Term Loan B  (senior, floating rate, ~1% annual amortisation)");
        writeFormula(ws, SU_TLB_R, SU_VAL_C, `=${crossRef(SHEET_ASM, ASM_TLB_R, ASM_DEBT_AMT_C)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, SU_SSN_R, 2, "Senior Secured Notes  (fixed rate, bullet maturity)");
        writeFormula(ws, SU_SSN_R, SU_VAL_C, `=${crossRef(SHEET_ASM, ASM_SSN_R, ASM_DEBT_AMT_C)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, SU_MEZ_R, 2, "Mezzanine  (junior debt, highest rate, partial PIK)");
        writeFormula(ws, SU_MEZ_R, SU_VAL_C, `=${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_DEBT_AMT_C)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeTotal(
          ws,
          SU_TOT_DEBT_R,
          SU_VAL_C,
          `=SUM(${sameRef(SU_TLB_R, SU_VAL_C)}:${sameRef(SU_MEZ_R, SU_VAL_C)})`,
          FMT_DOLLAR,
          { label: "Total Debt", colour: FORMULA_BLACK }
        );
        writeLabel(ws, 10, 2, "Equity", { bold: true, italic: true, colour: "404040" });
        const _ev = crossRef(SHEET_ASM, ASM_ENTRY_EV_R, ASM_INPUT_COL);
        const _tf = crossRef(SHEET_ASM, ASM_TRANS_FEES_R, ASM_INPUT_COL);
        const _ff = crossRef(SHEET_ASM, ASM_FIN_FEES_R, ASM_INPUT_COL);
        const _ro = crossRef(SHEET_ASM, ASM_MGMT_ROLLOVER_R, ASM_INPUT_COL);
        const _td = sameRef(SU_TOT_DEBT_R, SU_VAL_C);
        const _eqPool = `(${_ev}*(1+${_tf})+${_td}*${_ff}-${_td})`;
        writeLabel(ws, SU_SPONS_EQ_R, 2, "Sponsor Equity  (PE fund cash contribution)");
        writeFormula(ws, SU_SPONS_EQ_R, SU_VAL_C, `=${_eqPool}*(1-${_ro})`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, SU_MGMT_EQ_R, 2, "Management Rollover  (non-cash \u2014 existing equity retained by mgmt)");
        writeFormula(ws, SU_MGMT_EQ_R, SU_VAL_C, `=${_eqPool}*${_ro}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeTotal(
          ws,
          SU_TOT_EQ_R,
          SU_VAL_C,
          `=${sameRef(SU_SPONS_EQ_R, SU_VAL_C)}+${sameRef(SU_MGMT_EQ_R, SU_VAL_C)}`,
          FMT_DOLLAR,
          { label: "Total Equity", colour: FORMULA_BLACK }
        );
        writeTotal(
          ws,
          SU_TOT_SRC_R,
          SU_VAL_C,
          `=${sameRef(SU_TOT_DEBT_R, SU_VAL_C)}+${sameRef(SU_TOT_EQ_R, SU_VAL_C)}`,
          FMT_DOLLAR,
          { label: "TOTAL SOURCES", colour: FORMULA_BLACK }
        );
        const _totSrc = sameRef(SU_TOT_SRC_R, SU_VAL_C);
        const _evRef = crossRef(SHEET_ASM, ASM_ENTRY_EV_R, ASM_INPUT_COL);
        for (const row of [SU_TLB_R, SU_SSN_R, SU_MEZ_R, SU_TOT_DEBT_R, SU_SPONS_EQ_R, SU_MGMT_EQ_R, SU_TOT_EQ_R, SU_TOT_SRC_R]) {
          writeFormula(ws, row, SU_PCT_C, `=${sameRef(row, SU_VAL_C)}/${_totSrc}`, FMT_PERCENT, { colour: FORMULA_BLACK });
          writeFormula(ws, row, SU_EV_PCT_C, `=${sameRef(row, SU_VAL_C)}/${_evRef}`, FMT_PERCENT, { colour: LINK_GREEN });
        }
        for (const row of [SU_TOT_DEBT_R, SU_TOT_EQ_R, SU_TOT_SRC_R]) {
          const cPct2 = cellRC(ws, row, SU_PCT_C);
          applyStyle(cPct2, { bold: true, fmt: FMT_PERCENT, bg: TOTAL_GREY, align: "right", colour: FORMULA_BLACK });
          cPct2.border = TOTAL_BORDER;
          const cEv2 = cellRC(ws, row, SU_EV_PCT_C);
          applyStyle(cEv2, { bold: true, fmt: FMT_PERCENT, bg: TOTAL_GREY, align: "right", colour: LINK_GREEN });
          cEv2.border = TOTAL_BORDER;
        }
        writeHeader(ws, 16, 2, "USES  \u2014  Where the money goes", 4);
        writeLabel(ws, 17, 2, "Transaction Costs", { bold: true, italic: true, colour: "404040" });
        writeLabel(ws, SU_PRICE_R, 2, "Purchase Price  (= Entry Enterprise Value)");
        writeFormula(ws, SU_PRICE_R, SU_VAL_C, `=${_ev}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, SU_TRANS_FEE_R, 2, "Transaction Fees  (advisory, legal, accounting)");
        writeFormula(ws, SU_TRANS_FEE_R, SU_VAL_C, `=${_ev}*${_tf}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, SU_FIN_FEE_R, 2, "Financing Fees  (debt arrangement, OID, syndication)");
        writeFormula(ws, SU_FIN_FEE_R, SU_VAL_C, `=${sameRef(SU_TOT_DEBT_R, SU_VAL_C)}*${_ff}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeTotal(
          ws,
          SU_TOT_USE_R,
          SU_VAL_C,
          `=SUM(${sameRef(SU_PRICE_R, SU_VAL_C)}:${sameRef(SU_FIN_FEE_R, SU_VAL_C)})`,
          FMT_DOLLAR,
          { label: "TOTAL USES", colour: FORMULA_BLACK }
        );
        const _totUse = sameRef(SU_TOT_USE_R, SU_VAL_C);
        for (const row of [SU_PRICE_R, SU_TRANS_FEE_R, SU_FIN_FEE_R, SU_TOT_USE_R]) {
          writeFormula(ws, row, SU_PCT_C, `=${sameRef(row, SU_VAL_C)}/${_totUse}`, FMT_PERCENT, { colour: FORMULA_BLACK });
          writeFormula(ws, row, SU_EV_PCT_C, `=${sameRef(row, SU_VAL_C)}/${_evRef}`, FMT_PERCENT, { colour: LINK_GREEN });
        }
        let cPct = cellRC(ws, SU_TOT_USE_R, SU_PCT_C);
        applyStyle(cPct, { bold: true, fmt: FMT_PERCENT, bg: TOTAL_GREY, align: "right", colour: FORMULA_BLACK });
        cPct.border = TOTAL_BORDER;
        let cEv = cellRC(ws, SU_TOT_USE_R, SU_EV_PCT_C);
        applyStyle(cEv, { bold: true, fmt: FMT_PERCENT, bg: TOTAL_GREY, align: "right", colour: LINK_GREEN });
        cEv.border = TOTAL_BORDER;
        writeLabel(ws, SU_CHECK_R, 2, "CHECK: Sources \u2212 Uses  (must = $0.0)", { bold: true, colour: "CC0000" });
        writeFormula(
          ws,
          SU_CHECK_R,
          SU_VAL_C,
          `=${sameRef(SU_TOT_SRC_R, SU_VAL_C)}-${sameRef(SU_TOT_USE_R, SU_VAL_C)}`,
          FMT_DOLLAR,
          { bold: true, colour: FORMULA_BLACK }
        );
        writeHeader(ws, 25, 2, "KEY LEVERAGE METRICS AT CLOSE", 4);
        writeLabel(ws, 26, 2, "Total Debt / LTM EBITDA  (Leverage Ratio)");
        writeFormula(
          ws,
          26,
          SU_VAL_C,
          `=${sameRef(SU_TOT_DEBT_R, SU_VAL_C)}/${crossRef(SHEET_ASM, ASM_LTM_EBITDA_R, ASM_INPUT_COL)}`,
          FMT_MULTIPLE,
          { bold: true, colour: LINK_GREEN }
        );
        writeLabel(ws, 26, SU_PCT_C, "Typical range: 4.5x \u2013 6.5x", { italic: true, size: 9, colour: "707070" });
        writeLabel(ws, 27, 2, "Total Equity / Entry EV  (Equity Contribution)");
        writeFormula(ws, 27, SU_VAL_C, `=${sameRef(SU_TOT_EQ_R, SU_VAL_C)}/${_evRef}`, FMT_PERCENT, { bold: true, colour: LINK_GREEN });
        writeLabel(ws, 27, SU_PCT_C, "Typical range: 30% \u2013 45%", { italic: true, size: 9, colour: "707070" });
        const _y1Ebitda = crossRef(SHEET_ASM, ASM_EBITDA_R, 3);
        const _tlbInt = `${crossRef(SHEET_ASM, ASM_TLB_R, ASM_DEBT_AMT_C)}*${crossRef(SHEET_ASM, ASM_TLB_R, ASM_RATE_C)}`;
        const _ssnInt = `${crossRef(SHEET_ASM, ASM_SSN_R, ASM_DEBT_AMT_C)}*${crossRef(SHEET_ASM, ASM_SSN_R, ASM_RATE_C)}`;
        const _mezInt = `${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_DEBT_AMT_C)}*(${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_RATE_C)}-${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_PIK_RATE_C)})`;
        writeLabel(ws, 28, 2, "EBITDA / Interest  (Coverage Ratio, Year 1 approx.)");
        writeFormula(
          ws,
          28,
          SU_VAL_C,
          `=(${_y1Ebitda})/(${_tlbInt}+${_ssnInt}+${_mezInt})`,
          FMT_MULTIPLE,
          { bold: true, colour: LINK_GREEN }
        );
        writeLabel(ws, 28, SU_PCT_C, "Must be > 2.0x", { italic: true, size: 9, colour: "707070" });
      }
      module.exports = { buildSourcesUses };
    }
  });

  // incomeStmt.js
  var require_incomeStmt = __commonJS({
    "incomeStmt.js"(exports, module) {
      var H = require_helpers();
      var S = require_styles();
      var C = require_constants();
      var { crossRef, sameRef, writeLabel, writeFormula, writeHeader, writeTotal, writeInfo } = H;
      var { FORMULA_BLACK, LINK_GREEN, DARK_BLUE, LIGHT_BLUE, FMT_DOLLAR, FMT_PERCENT } = S;
      var {
        SHEET_ASM,
        SHEET_DS,
        LTM_COL,
        YR1_COL,
        N_YEARS,
        ASM_REVENUE_R,
        ASM_EBITDA_R,
        ASM_DA_R,
        ASM_TAX_R,
        ASM_TLB_R,
        ASM_SSN_R,
        ASM_MEZ_R,
        ASM_DEBT_AMT_C,
        ASM_RATE_C,
        DS_TLB_INT_R,
        DS_SSN_INT_R,
        DS_MEZ_CASH_INT_R,
        DS_MEZ_PIK_INT_R,
        IS_REV_R,
        IS_GRW_R,
        IS_EBD_R,
        IS_MRG_R,
        IS_DA_R,
        IS_EBI_R,
        IS_EBI_MRG_R,
        IS_INT_TLB_R,
        IS_INT_SSN_R,
        IS_INT_MEZ_R,
        IS_INT_TOT_R,
        IS_EBT_R,
        IS_TAX_R,
        IS_NI_R,
        IS_NI_MRG_R
      } = C;
      function buildIncomeStatement(ws) {
        ws.getColumn(1).width = 2;
        ws.getColumn(2).width = 45;
        ws.getColumn(3).width = 15;
        [4, 5, 6, 7, 8].forEach((n) => ws.getColumn(n).width = 13);
        writeLabel(ws, 1, 2, "INCOME STATEMENT", { bold: true, size: 14, colour: DARK_BLUE });
        writeLabel(
          ws,
          2,
          2,
          "All figures $M  |  Interest from Debt Schedule (Year 1-5)  |  ",
          { italic: true, size: 9, colour: "808080" }
        );
        writeHeader(ws, 3, 2, "PROJECTED INCOME STATEMENT  ($M)", 7);
        writeLabel(ws, 4, 2, "");
        writeLabel(ws, 4, LTM_COL, "LTM", { bold: true, bg: LIGHT_BLUE, align: "center" });
        for (let i = 0; i < N_YEARS; i++)
          writeLabel(ws, 4, YR1_COL + i, `Year ${i + 1}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeLabel(ws, IS_REV_R, 2, "Revenue", { bold: true });
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeFormula(ws, IS_REV_R, col, `=${crossRef(SHEET_ASM, ASM_REVENUE_R, col)}`, FMT_DOLLAR, { bold: true, colour: LINK_GREEN });
        writeLabel(ws, IS_GRW_R, 2, "   YoY Growth");
        writeLabel(ws, IS_GRW_R, LTM_COL, "\u2014", { align: "center" });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeInfo(ws, IS_GRW_R, col, `=${sameRef(IS_REV_R, col)}/${sameRef(IS_REV_R, col - 1)}-1`, FMT_PERCENT);
        writeLabel(ws, IS_EBD_R, 2, "EBITDA", { bold: true });
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeFormula(ws, IS_EBD_R, col, `=${crossRef(SHEET_ASM, ASM_EBITDA_R, col)}`, FMT_DOLLAR, { bold: true, colour: LINK_GREEN });
        writeLabel(ws, IS_MRG_R, 2, "   EBITDA Margin");
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeInfo(ws, IS_MRG_R, col, `=${sameRef(IS_EBD_R, col)}/${sameRef(IS_REV_R, col)}`);
        writeLabel(ws, IS_DA_R, 2, "Depreciation & Amortisation  (D&A)");
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeFormula(ws, IS_DA_R, col, `=${crossRef(SHEET_ASM, ASM_DA_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, IS_EBI_R, 2, "EBIT  (Operating Profit)", { bold: true });
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeFormula(ws, IS_EBI_R, col, `=${sameRef(IS_EBD_R, col)}-${sameRef(IS_DA_R, col)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        writeLabel(ws, IS_EBI_MRG_R, 2, "   EBIT Margin");
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeInfo(ws, IS_EBI_MRG_R, col, `=${sameRef(IS_EBI_R, col)}/${sameRef(IS_REV_R, col)}`);
        writeLabel(
          ws,
          14,
          2,
          "Interest Expense  (LTM = simplified  |  Year 1-5 = from Debt Schedule)",
          { italic: true, colour: "606060" }
        );
        writeLabel(ws, IS_INT_TLB_R, 2, "   Interest \u2014 Term Loan B");
        writeFormula(
          ws,
          IS_INT_TLB_R,
          LTM_COL,
          `=${crossRef(SHEET_ASM, ASM_TLB_R, ASM_DEBT_AMT_C)}*${crossRef(SHEET_ASM, ASM_TLB_R, ASM_RATE_C)}`,
          FMT_DOLLAR,
          { colour: LINK_GREEN }
        );
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, IS_INT_TLB_R, col, `=${crossRef(SHEET_DS, DS_TLB_INT_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, IS_INT_SSN_R, 2, "   Interest \u2014 Senior Secured Notes");
        writeFormula(
          ws,
          IS_INT_SSN_R,
          LTM_COL,
          `=${crossRef(SHEET_ASM, ASM_SSN_R, ASM_DEBT_AMT_C)}*${crossRef(SHEET_ASM, ASM_SSN_R, ASM_RATE_C)}`,
          FMT_DOLLAR,
          { colour: LINK_GREEN }
        );
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, IS_INT_SSN_R, col, `=${crossRef(SHEET_DS, DS_SSN_INT_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, IS_INT_MEZ_R, 2, "   Interest \u2014 Mezzanine  (cash + PIK = full P&L charge)");
        writeFormula(
          ws,
          IS_INT_MEZ_R,
          LTM_COL,
          `=${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_DEBT_AMT_C)}*${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_RATE_C)}`,
          FMT_DOLLAR,
          { colour: LINK_GREEN }
        );
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            IS_INT_MEZ_R,
            col,
            `=${crossRef(SHEET_DS, DS_MEZ_CASH_INT_R, col)}+${crossRef(SHEET_DS, DS_MEZ_PIK_INT_R, col)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        for (let i = 0; i <= N_YEARS; i++) {
          const col = LTM_COL + i;
          writeTotal(
            ws,
            IS_INT_TOT_R,
            col,
            `=SUM(${sameRef(IS_INT_TLB_R, col)}:${sameRef(IS_INT_MEZ_R, col)})`,
            FMT_DOLLAR,
            { label: i === 0 ? "Total Interest Expense" : null, colour: FORMULA_BLACK }
          );
        }
        writeLabel(ws, IS_EBT_R, 2, "EBT  (Earnings Before Tax)", { bold: true });
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeFormula(ws, IS_EBT_R, col, `=${sameRef(IS_EBI_R, col)}-${sameRef(IS_INT_TOT_R, col)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        writeLabel(ws, IS_TAX_R, 2, "Tax Expense");
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeFormula(
            ws,
            IS_TAX_R,
            col,
            `=MAX(${sameRef(IS_EBT_R, col)}*${crossRef(SHEET_ASM, ASM_TAX_R, col)},0)`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        for (let i = 0; i <= N_YEARS; i++) {
          const col = LTM_COL + i;
          writeTotal(
            ws,
            IS_NI_R,
            col,
            `=${sameRef(IS_EBT_R, col)}-${sameRef(IS_TAX_R, col)}`,
            FMT_DOLLAR,
            { label: i === 0 ? "NET INCOME" : null, colour: FORMULA_BLACK }
          );
        }
        writeLabel(ws, IS_NI_MRG_R, 2, "   Net Income Margin");
        for (let col = LTM_COL; col <= LTM_COL + N_YEARS; col++)
          writeInfo(ws, IS_NI_MRG_R, col, `=IFERROR(${sameRef(IS_NI_R, col)}/${sameRef(IS_REV_R, col)},0)`);
      }
      module.exports = { buildIncomeStatement };
    }
  });

  // debtSchedule.js
  var require_debtSchedule = __commonJS({
    "debtSchedule.js"(exports, module) {
      var H = require_helpers();
      var S = require_styles();
      var C = require_constants();
      var { crossRef, sameRef, writeLabel, writeFormula, writeHeader, writeTotal } = H;
      var { FORMULA_BLACK, LINK_GREEN, DARK_BLUE, LIGHT_BLUE, TOTAL_GREY, FMT_DOLLAR } = S;
      var {
        SHEET_ASM,
        SHEET_CF,
        YR1_COL,
        N_YEARS,
        ASM_TLB_R,
        ASM_SSN_R,
        ASM_MEZ_R,
        ASM_DEBT_AMT_C,
        ASM_RATE_C,
        ASM_CASH_AMRT_C,
        ASM_PIK_RATE_C,
        CF_AVAIL_R,
        DS_YR0_C,
        DS_TLB_HDR_R,
        DS_TLB_BEG_R,
        DS_TLB_AMRT_R,
        DS_TLB_OPT_R,
        DS_TLB_END_R,
        DS_TLB_INT_R,
        DS_SSN_HDR_R,
        DS_SSN_BEG_R,
        DS_SSN_AMRT_R,
        DS_SSN_OPT_R,
        DS_SSN_END_R,
        DS_SSN_INT_R,
        DS_MEZ_HDR_R,
        DS_MEZ_BEG_R,
        DS_MEZ_PIK_R,
        DS_MEZ_OPT_R,
        DS_MEZ_END_R,
        DS_MEZ_CASH_INT_R,
        DS_MEZ_PIK_INT_R,
        DS_SUM_HDR_R,
        DS_TDEBT_R,
        DS_TCASH_INT_R,
        DS_TPIK_R,
        DS_TINT_R,
        DS_TMAND_R
      } = C;
      function buildDebtSchedule(ws) {
        ws.getColumn(1).width = 2;
        ws.getColumn(2).width = 52;
        ws.getColumn(3).width = 15;
        [4, 5, 6, 7, 8].forEach((n) => ws.getColumn(n).width = 13);
        writeLabel(ws, 1, 2, "DEBT SCHEDULE", { bold: true, size: 14, colour: DARK_BLUE });
        writeLabel(
          ws,
          2,
          2,
          "Tracks balance, amortisation, cash sweep and interest for all tranches  |  ",
          { italic: true, size: 9, colour: "808080" }
        );
        writeHeader(ws, 3, 2, "DEBT SCHEDULE  ($M)", 7);
        writeLabel(ws, 4, 3, "Year 0  (close)", { bold: true, bg: TOTAL_GREY, align: "center" });
        for (let i = 1; i <= N_YEARS; i++)
          writeLabel(ws, 4, 3 + i, `Year ${i}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeHeader(ws, DS_TLB_HDR_R, 2, "TERM LOAN B", 7);
        writeLabel(ws, DS_TLB_BEG_R, 2, "Beginning Balance");
        writeFormula(ws, DS_TLB_BEG_R, DS_YR0_C, `=${crossRef(SHEET_ASM, ASM_TLB_R, ASM_DEBT_AMT_C)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, DS_TLB_BEG_R, col, `=${sameRef(DS_TLB_END_R, col - 1)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        writeLabel(ws, DS_TLB_AMRT_R, 2, "(-) Mandatory Amortisation");
        writeFormula(ws, DS_TLB_AMRT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_TLB_AMRT_R,
            col,
            `=MIN(${crossRef(SHEET_ASM, ASM_TLB_R, ASM_DEBT_AMT_C)}*${crossRef(SHEET_ASM, ASM_TLB_R, ASM_CASH_AMRT_C)},${sameRef(DS_TLB_BEG_R, col)})`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeLabel(ws, DS_TLB_OPT_R, 2, "(-) Optional Paydown  (cash sweep \u2014 FIRST priority in waterfall)");
        writeFormula(ws, DS_TLB_OPT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++) {
          const _avail = crossRef(SHEET_CF, CF_AVAIL_R, col);
          const _outstanding = `MAX(${sameRef(DS_TLB_BEG_R, col)}-${sameRef(DS_TLB_AMRT_R, col)},0)`;
          writeFormula(ws, DS_TLB_OPT_R, col, `=MIN(MAX(${_avail},0),${_outstanding})`, FMT_DOLLAR, { colour: LINK_GREEN });
        }
        writeLabel(ws, DS_TLB_END_R, 2, "Ending Balance", { bold: true });
        writeFormula(ws, DS_TLB_END_R, DS_YR0_C, `=${sameRef(DS_TLB_BEG_R, DS_YR0_C)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_TLB_END_R,
            col,
            `=MAX(${sameRef(DS_TLB_BEG_R, col)}-${sameRef(DS_TLB_AMRT_R, col)}-${sameRef(DS_TLB_OPT_R, col)},0)`,
            FMT_DOLLAR,
            { bold: true, colour: FORMULA_BLACK }
          );
        writeLabel(ws, DS_TLB_INT_R, 2, "Cash Interest  (= Beginning Balance \xD7 Rate)");
        writeFormula(ws, DS_TLB_INT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_TLB_INT_R,
            col,
            `=${sameRef(DS_TLB_BEG_R, col)}*${crossRef(SHEET_ASM, ASM_TLB_R, ASM_RATE_C)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeHeader(ws, DS_SSN_HDR_R, 2, "SENIOR SECURED NOTES", 7);
        writeLabel(ws, DS_SSN_BEG_R, 2, "Beginning Balance");
        writeFormula(ws, DS_SSN_BEG_R, DS_YR0_C, `=${crossRef(SHEET_ASM, ASM_SSN_R, ASM_DEBT_AMT_C)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, DS_SSN_BEG_R, col, `=${sameRef(DS_SSN_END_R, col - 1)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        writeLabel(ws, DS_SSN_AMRT_R, 2, "(-) Mandatory Amortisation  (none \u2014 bullet bond, repaid at maturity)");
        for (let col = DS_YR0_C; col <= DS_YR0_C + N_YEARS; col++)
          writeFormula(ws, DS_SSN_AMRT_R, col, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        writeLabel(ws, DS_SSN_OPT_R, 2, "(-) Optional Paydown  (SECOND priority \u2014 remainder after TLB)");
        writeFormula(ws, DS_SSN_OPT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++) {
          const _avail = crossRef(SHEET_CF, CF_AVAIL_R, col);
          const _usedTlb = sameRef(DS_TLB_OPT_R, col);
          const _remaining = `MAX(${_avail}-${_usedTlb},0)`;
          const _outstanding = `MAX(${sameRef(DS_SSN_BEG_R, col)},0)`;
          writeFormula(ws, DS_SSN_OPT_R, col, `=MIN(${_remaining},${_outstanding})`, FMT_DOLLAR, { colour: LINK_GREEN });
        }
        writeLabel(ws, DS_SSN_END_R, 2, "Ending Balance", { bold: true });
        writeFormula(ws, DS_SSN_END_R, DS_YR0_C, `=${sameRef(DS_SSN_BEG_R, DS_YR0_C)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_SSN_END_R,
            col,
            `=MAX(${sameRef(DS_SSN_BEG_R, col)}-${sameRef(DS_SSN_OPT_R, col)},0)`,
            FMT_DOLLAR,
            { bold: true, colour: FORMULA_BLACK }
          );
        writeLabel(ws, DS_SSN_INT_R, 2, "Cash Interest  (= Beginning Balance \xD7 Rate)");
        writeFormula(ws, DS_SSN_INT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_SSN_INT_R,
            col,
            `=${sameRef(DS_SSN_BEG_R, col)}*${crossRef(SHEET_ASM, ASM_SSN_R, ASM_RATE_C)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeHeader(ws, DS_MEZ_HDR_R, 2, "MEZZANINE", 7);
        writeLabel(ws, DS_MEZ_BEG_R, 2, "Beginning Balance");
        writeFormula(ws, DS_MEZ_BEG_R, DS_YR0_C, `=${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_DEBT_AMT_C)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, DS_MEZ_BEG_R, col, `=${sameRef(DS_MEZ_END_R, col - 1)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        writeLabel(ws, DS_MEZ_PIK_R, 2, "(+) PIK Accrual  (non-cash \u2014 accretes to principal balance)");
        writeFormula(ws, DS_MEZ_PIK_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_MEZ_PIK_R,
            col,
            `=${sameRef(DS_MEZ_BEG_R, col)}*${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_PIK_RATE_C)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeLabel(ws, DS_MEZ_OPT_R, 2, "(-) Optional Paydown  (THIRD priority \u2014 remainder after TLB and SSN)");
        writeFormula(ws, DS_MEZ_OPT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++) {
          const _avail = crossRef(SHEET_CF, CF_AVAIL_R, col);
          const _usedTlb = sameRef(DS_TLB_OPT_R, col);
          const _usedSsn = sameRef(DS_SSN_OPT_R, col);
          const _remaining = `MAX(${_avail}-${_usedTlb}-${_usedSsn},0)`;
          const _outstanding = `MAX(${sameRef(DS_MEZ_BEG_R, col)}+${sameRef(DS_MEZ_PIK_R, col)},0)`;
          writeFormula(ws, DS_MEZ_OPT_R, col, `=MIN(${_remaining},${_outstanding})`, FMT_DOLLAR, { colour: LINK_GREEN });
        }
        writeLabel(ws, DS_MEZ_END_R, 2, "Ending Balance", { bold: true });
        writeFormula(ws, DS_MEZ_END_R, DS_YR0_C, `=${sameRef(DS_MEZ_BEG_R, DS_YR0_C)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_MEZ_END_R,
            col,
            `=MAX(${sameRef(DS_MEZ_BEG_R, col)}+${sameRef(DS_MEZ_PIK_R, col)}-${sameRef(DS_MEZ_OPT_R, col)},0)`,
            FMT_DOLLAR,
            { bold: true, colour: FORMULA_BLACK }
          );
        writeLabel(ws, DS_MEZ_CASH_INT_R, 2, "Cash Interest  (= Beginning \xD7 (Total Rate \u2212 PIK Rate))");
        writeFormula(ws, DS_MEZ_CASH_INT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            DS_MEZ_CASH_INT_R,
            col,
            `=${sameRef(DS_MEZ_BEG_R, col)}*(${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_RATE_C)}-${crossRef(SHEET_ASM, ASM_MEZ_R, ASM_PIK_RATE_C)})`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeLabel(ws, DS_MEZ_PIK_INT_R, 2, "PIK Interest  (non-cash \u2014 same as PIK accrual above)");
        writeFormula(ws, DS_MEZ_PIK_INT_R, DS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, DS_MEZ_PIK_INT_R, col, `=${sameRef(DS_MEZ_PIK_R, col)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        writeHeader(ws, DS_SUM_HDR_R, 2, "SUMMARY", 7);
        writeLabel(ws, DS_SUM_HDR_R + 1, DS_YR0_C, "Year 0  (close)", { bold: true, bg: TOTAL_GREY, align: "center" });
        for (let i = 1; i <= N_YEARS; i++)
          writeLabel(ws, DS_SUM_HDR_R + 1, DS_YR0_C + i, `Year ${i}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeTotal(
          ws,
          DS_TDEBT_R,
          DS_YR0_C,
          `=${sameRef(DS_TLB_END_R, DS_YR0_C)}+${sameRef(DS_SSN_END_R, DS_YR0_C)}+${sameRef(DS_MEZ_END_R, DS_YR0_C)}`,
          FMT_DOLLAR,
          { label: "Total Debt  (ending balance)", colour: FORMULA_BLACK }
        );
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeTotal(
            ws,
            DS_TDEBT_R,
            col,
            `=${sameRef(DS_TLB_END_R, col)}+${sameRef(DS_SSN_END_R, col)}+${sameRef(DS_MEZ_END_R, col)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
        writeTotal(
          ws,
          DS_TCASH_INT_R,
          DS_YR0_C,
          `=${sameRef(DS_TLB_INT_R, DS_YR0_C)}+${sameRef(DS_SSN_INT_R, DS_YR0_C)}+${sameRef(DS_MEZ_CASH_INT_R, DS_YR0_C)}`,
          FMT_DOLLAR,
          { label: "Total Cash Interest  [\u2192 used by Cash Flow and IS]", colour: FORMULA_BLACK }
        );
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeTotal(
            ws,
            DS_TCASH_INT_R,
            col,
            `=${sameRef(DS_TLB_INT_R, col)}+${sameRef(DS_SSN_INT_R, col)}+${sameRef(DS_MEZ_CASH_INT_R, col)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, DS_TPIK_R, 2, "Total PIK Interest  (non-cash)");
        for (let col = DS_YR0_C; col <= DS_YR0_C + N_YEARS; col++)
          writeFormula(ws, DS_TPIK_R, col, `=${sameRef(DS_MEZ_PIK_INT_R, col)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        writeTotal(
          ws,
          DS_TINT_R,
          DS_YR0_C,
          `=${sameRef(DS_TCASH_INT_R, DS_YR0_C)}+${sameRef(DS_TPIK_R, DS_YR0_C)}`,
          FMT_DOLLAR,
          { label: "Total Interest  (cash + PIK)", colour: FORMULA_BLACK }
        );
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeTotal(
            ws,
            DS_TINT_R,
            col,
            `=${sameRef(DS_TCASH_INT_R, col)}+${sameRef(DS_TPIK_R, col)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
        writeLabel(ws, DS_TMAND_R, 2, "Total Mandatory Amortisation  [\u2192 used by Cash Flow]");
        for (let col = DS_YR0_C; col <= DS_YR0_C + N_YEARS; col++)
          writeFormula(
            ws,
            DS_TMAND_R,
            col,
            `=${sameRef(DS_TLB_AMRT_R, col)}+${sameRef(DS_SSN_AMRT_R, col)}`,
            FMT_DOLLAR,
            { colour: FORMULA_BLACK }
          );
      }
      module.exports = { buildDebtSchedule };
    }
  });

  // cashFlow.js
  var require_cashFlow = __commonJS({
    "cashFlow.js"(exports, module) {
      var H = require_helpers();
      var S = require_styles();
      var C = require_constants();
      var { crossRef, sameRef, writeLabel, writeFormula, writeHeader, writeTotal } = H;
      var { FORMULA_BLACK, LINK_GREEN, DARK_BLUE, LIGHT_BLUE, TOTAL_GREY, FMT_DOLLAR } = S;
      var {
        SHEET_ASM,
        SHEET_DS,
        DS_YR0_C,
        YR1_COL,
        N_YEARS,
        ASM_EBITDA_R,
        ASM_DA_R,
        ASM_CAPEX_R,
        ASM_NWC_R,
        ASM_TAX_R,
        DS_TCASH_INT_R,
        DS_TINT_R,
        DS_TMAND_R,
        DS_TLB_OPT_R,
        DS_SSN_OPT_R,
        DS_MEZ_OPT_R,
        CF_EBITDA_R,
        CF_INT_R,
        CF_TAX_R,
        CF_CAPEX_R,
        CF_NWC_R,
        CF_OFCF_R,
        CF_MAND_R,
        CF_AVAIL_R,
        CF_SWEEP_R,
        CF_NCF_R,
        CF_BEG_CASH_R,
        CF_END_CASH_R
      } = C;
      var GREY9 = { italic: true, size: 9, colour: "909090", align: "center" };
      function buildCashFlow(ws) {
        ws.getColumn(1).width = 2;
        ws.getColumn(2).width = 55;
        ws.getColumn(3).width = 15;
        [4, 5, 6, 7, 8].forEach((n) => ws.getColumn(n).width = 13);
        writeLabel(ws, 1, 2, "CASH FLOW STATEMENT", { bold: true, size: 14, colour: DARK_BLUE });
        writeLabel(
          ws,
          2,
          2,
          "EBITDA \u2192 Operating FCF \u2192 debt service \u2192 ending cash balance  |  ",
          { italic: true, size: 9, colour: "808080" }
        );
        writeHeader(ws, 3, 2, "FREE CASH FLOW  ($M)", 7);
        writeLabel(ws, 4, DS_YR0_C, "Year 0", { bold: true, bg: TOTAL_GREY, align: "center" });
        for (let i = 1; i <= N_YEARS; i++)
          writeLabel(ws, 4, DS_YR0_C + i, `Year ${i}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeLabel(ws, 5, DS_YR0_C, "\u2190 deal close", GREY9);
        writeLabel(ws, CF_EBITDA_R, 2, "EBITDA", { bold: true });
        writeLabel(ws, CF_EBITDA_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, CF_EBITDA_R, col, `=${crossRef(SHEET_ASM, ASM_EBITDA_R, col)}`, FMT_DOLLAR, { bold: true, colour: LINK_GREEN });
        writeLabel(ws, CF_INT_R, 2, "(-) Cash Interest Expense  [from Debt Schedule \u2014 actual declining balances]");
        writeLabel(ws, CF_INT_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, CF_INT_R, col, `=${crossRef(SHEET_DS, DS_TCASH_INT_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, CF_TAX_R, 2, "(-) Cash Taxes  [= MAX((EBITDA \u2212 D&A \u2212 Total Interest incl. PIK) \xD7 Rate, 0)]");
        writeLabel(ws, CF_TAX_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++) {
          const _ebitda = crossRef(SHEET_ASM, ASM_EBITDA_R, col);
          const _da = crossRef(SHEET_ASM, ASM_DA_R, col);
          const _int = crossRef(SHEET_DS, DS_TINT_R, col);
          const _rate = crossRef(SHEET_ASM, ASM_TAX_R, col);
          writeFormula(ws, CF_TAX_R, col, `=MAX((${_ebitda}-${_da}-${_int})*${_rate},0)`, FMT_DOLLAR, { colour: LINK_GREEN });
        }
        writeLabel(ws, CF_CAPEX_R, 2, "(-) Capital Expenditure  (CapEx)");
        writeLabel(ws, CF_CAPEX_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, CF_CAPEX_R, col, `=${crossRef(SHEET_ASM, ASM_CAPEX_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, CF_NWC_R, 2, "(-) Increase in Net Working Capital  (NWC)");
        writeLabel(ws, CF_NWC_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, CF_NWC_R, col, `=${crossRef(SHEET_ASM, ASM_NWC_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, CF_OFCF_R, DS_YR0_C, "\u2014", GREY9);
        for (let i = 0; i < N_YEARS; i++) {
          const col = YR1_COL + i;
          writeTotal(
            ws,
            CF_OFCF_R,
            col,
            `=${sameRef(CF_EBITDA_R, col)}-${sameRef(CF_INT_R, col)}-${sameRef(CF_TAX_R, col)}-${sameRef(CF_CAPEX_R, col)}-${sameRef(CF_NWC_R, col)}`,
            FMT_DOLLAR,
            { label: i === 0 ? "Operating Free Cash Flow" : null, colour: FORMULA_BLACK }
          );
        }
        writeLabel(ws, CF_MAND_R, 2, "(-) Mandatory Debt Amortisation  [from Debt Schedule]");
        writeLabel(ws, CF_MAND_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, CF_MAND_R, col, `=${crossRef(SHEET_DS, DS_TMAND_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, CF_AVAIL_R, 2, "Cash Available for Optional Debt Sweep  [\u2B05 Debt Schedule reads this row]", { bold: true });
        writeLabel(ws, CF_AVAIL_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            CF_AVAIL_R,
            col,
            `=MAX(${sameRef(CF_OFCF_R, col)}-${sameRef(CF_MAND_R, col)},0)`,
            FMT_DOLLAR,
            { bold: true, colour: FORMULA_BLACK }
          );
        writeLabel(ws, CF_SWEEP_R, 2, "(-) Optional Cash Sweep  [sum from Debt Schedule: TLB + SSN + MEZ]");
        writeLabel(ws, CF_SWEEP_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            CF_SWEEP_R,
            col,
            `= ${crossRef(SHEET_DS, DS_TLB_OPT_R, col)}+${crossRef(SHEET_DS, DS_SSN_OPT_R, col)}+${crossRef(SHEET_DS, DS_MEZ_OPT_R, col)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeLabel(ws, CF_NCF_R, 2, "Net Cash Retained on Balance Sheet");
        writeLabel(ws, CF_NCF_R, DS_YR0_C, "\u2014", GREY9);
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, CF_NCF_R, col, `=${sameRef(CF_AVAIL_R, col)}-${sameRef(CF_SWEEP_R, col)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        writeLabel(ws, CF_BEG_CASH_R, 2, "Beginning Cash Balance");
        for (let col = DS_YR0_C; col <= DS_YR0_C + N_YEARS; col++) {
          if (col === DS_YR0_C)
            writeFormula(ws, CF_BEG_CASH_R, col, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
          else
            writeFormula(ws, CF_BEG_CASH_R, col, `=${sameRef(CF_END_CASH_R, col - 1)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        }
        writeLabel(ws, CF_END_CASH_R, 2, "Ending Cash Balance", { bold: true });
        writeFormula(ws, CF_END_CASH_R, DS_YR0_C, "=0", FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, CF_END_CASH_R, col, `=${sameRef(CF_BEG_CASH_R, col)}+${sameRef(CF_NCF_R, col)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
      }
      module.exports = { buildCashFlow };
    }
  });

  // balanceSheet.js
  var require_balanceSheet = __commonJS({
    "balanceSheet.js"(exports, module) {
      var H = require_helpers();
      var S = require_styles();
      var C = require_constants();
      var { crossRef, sameRef, writeLabel, writeFormula, writeHeader } = H;
      var { FORMULA_BLACK, LINK_GREEN, DARK_BLUE, LIGHT_BLUE, TOTAL_GREY, FMT_DOLLAR } = S;
      var {
        SHEET_ASM,
        SHEET_CF,
        SHEET_DS,
        SHEET_IS,
        SHEET_SU,
        YR1_COL,
        N_YEARS,
        ASM_ENTRY_EV_R,
        ASM_TRANS_FEES_R,
        ASM_INIT_NWC_BAL_R,
        ASM_INIT_PPE_R,
        ASM_INPUT_COL,
        ASM_DA_R,
        CF_END_CASH_R,
        CF_NWC_R,
        CF_CAPEX_R,
        DS_TLB_END_R,
        DS_SSN_END_R,
        DS_MEZ_END_R,
        DS_TDEBT_R,
        IS_NI_R,
        SU_FIN_FEE_R,
        SU_VAL_C,
        SU_TOT_EQ_R,
        BS_YR0_C,
        BS_AST_HDR_R,
        BS_CASH_R,
        BS_NWC_R,
        BS_PPE_R,
        BS_GW_R,
        BS_DFC_R,
        BS_TAST_R,
        BS_LBL_HDR_R,
        BS_TLB_R,
        BS_SSN_R,
        BS_MEZ_R,
        BS_TDBT_R,
        BS_EQ_HDR_R,
        BS_INIT_EQ_R,
        BS_RE_R,
        BS_TEQ_R,
        BS_TLE_R,
        BS_CHK_R
      } = C;
      function buildBalanceSheet(ws) {
        ws.getColumn(1).width = 2;
        ws.getColumn(2).width = 55;
        ws.getColumn(3).width = 15;
        [4, 5, 6, 7, 8].forEach((n) => ws.getColumn(n).width = 13);
        writeLabel(ws, 1, 2, "BALANCE SHEET", { bold: true, size: 14, colour: DARK_BLUE });
        writeLabel(
          ws,
          2,
          2,
          "Year 0 = deal close  |  CHECK row must show exactly 0.0 every year  |  ",
          { italic: true, size: 9, colour: "808080" }
        );
        writeHeader(ws, 3, 2, "BALANCE SHEET  ($M)", 7);
        writeLabel(ws, 4, BS_YR0_C, "Year 0  (close)", { bold: true, bg: TOTAL_GREY, align: "center" });
        for (let i = 1; i <= N_YEARS; i++)
          writeLabel(ws, 4, BS_YR0_C + i, `Year ${i}`, { bold: true, bg: LIGHT_BLUE, align: "center" });
        writeHeader(ws, BS_AST_HDR_R, 2, "ASSETS", 7);
        writeLabel(ws, BS_CASH_R, 2, "Cash  ($M)");
        writeFormula(ws, BS_CASH_R, BS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(ws, BS_CASH_R, col, `=${crossRef(SHEET_CF, CF_END_CASH_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, BS_NWC_R, 2, "Net Working Capital  ($M)");
        writeFormula(ws, BS_NWC_R, BS_YR0_C, `=${crossRef(SHEET_ASM, ASM_INIT_NWC_BAL_R, ASM_INPUT_COL)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            BS_NWC_R,
            col,
            `=${sameRef(BS_NWC_R, col - 1)}+${crossRef(SHEET_CF, CF_NWC_R, col)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeLabel(ws, BS_PPE_R, 2, "PP&E  (net book value, $M)");
        writeFormula(ws, BS_PPE_R, BS_YR0_C, `=${crossRef(SHEET_ASM, ASM_INIT_PPE_R, ASM_INPUT_COL)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            BS_PPE_R,
            col,
            `=${sameRef(BS_PPE_R, col - 1)}+${crossRef(SHEET_CF, CF_CAPEX_R, col)}-${crossRef(SHEET_ASM, ASM_DA_R, col)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        const _ev = crossRef(SHEET_ASM, ASM_ENTRY_EV_R, ASM_INPUT_COL);
        const _tf = crossRef(SHEET_ASM, ASM_TRANS_FEES_R, ASM_INPUT_COL);
        const _nwc0 = crossRef(SHEET_ASM, ASM_INIT_NWC_BAL_R, ASM_INPUT_COL);
        const _ppe0 = crossRef(SHEET_ASM, ASM_INIT_PPE_R, ASM_INPUT_COL);
        const _gw = `=${_ev}*(1+${_tf})-${_nwc0}-${_ppe0}`;
        writeLabel(ws, BS_GW_R, 2, "Goodwill  ($M)  [= EV\xD7(1+fee%) \u2212 NWC \u2212 PP&E, constant]");
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_GW_R, col, _gw, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, BS_DFC_R, 2, "Deferred Financing Costs  ($M)  [simplified: held constant]");
        const _dfc0 = crossRef(SHEET_SU, SU_FIN_FEE_R, SU_VAL_C);
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_DFC_R, col, `=${_dfc0}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, BS_TAST_R, 2, "TOTAL ASSETS", { bold: true });
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_TAST_R, col, `=SUM(${sameRef(BS_CASH_R, col)}:${sameRef(BS_DFC_R, col)})`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        writeHeader(ws, BS_LBL_HDR_R, 2, "LIABILITIES", 7);
        writeLabel(ws, BS_TLB_R, 2, "Term Loan B");
        writeLabel(ws, BS_SSN_R, 2, "Senior Secured Notes");
        writeLabel(ws, BS_MEZ_R, 2, "Mezzanine  (includes PIK accrual each year)");
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++) {
          writeFormula(ws, BS_TLB_R, col, `=${crossRef(SHEET_DS, DS_TLB_END_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
          writeFormula(ws, BS_SSN_R, col, `=${crossRef(SHEET_DS, DS_SSN_END_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
          writeFormula(ws, BS_MEZ_R, col, `=${crossRef(SHEET_DS, DS_MEZ_END_R, col)}`, FMT_DOLLAR, { colour: LINK_GREEN });
        }
        writeLabel(ws, BS_TDBT_R, 2, "Total Debt", { bold: true });
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_TDBT_R, col, `=${crossRef(SHEET_DS, DS_TDEBT_R, col)}`, FMT_DOLLAR, { bold: true, colour: LINK_GREEN });
        writeHeader(ws, BS_EQ_HDR_R, 2, "EQUITY", 7);
        const _initEq = crossRef(SHEET_SU, SU_TOT_EQ_R, SU_VAL_C);
        writeLabel(ws, BS_INIT_EQ_R, 2, "Initial Equity Invested  (constant)");
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_INIT_EQ_R, col, `=${_initEq}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, BS_RE_R, 2, "Retained Earnings  (cumulative Net Income from IS)");
        writeFormula(ws, BS_RE_R, BS_YR0_C, "=0", FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let col = YR1_COL; col < YR1_COL + N_YEARS; col++)
          writeFormula(
            ws,
            BS_RE_R,
            col,
            `=${sameRef(BS_RE_R, col - 1)}+${crossRef(SHEET_IS, IS_NI_R, col)}`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        writeLabel(ws, BS_TEQ_R, 2, "TOTAL EQUITY", { bold: true });
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_TEQ_R, col, `=${sameRef(BS_INIT_EQ_R, col)}+${sameRef(BS_RE_R, col)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        writeLabel(ws, BS_TLE_R, 2, "Total Liabilities + Equity", { bold: true });
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_TLE_R, col, `=${sameRef(BS_TDBT_R, col)}+${sameRef(BS_TEQ_R, col)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
        writeLabel(ws, BS_CHK_R, 2, "CHECK  Assets \u2212 (L + E)  \u2192  must = 0.0 every year", { bold: true, colour: "CC0000" });
        for (let col = BS_YR0_C; col <= BS_YR0_C + N_YEARS; col++)
          writeFormula(ws, BS_CHK_R, col, `=${sameRef(BS_TAST_R, col)}-${sameRef(BS_TLE_R, col)}`, FMT_DOLLAR, { bold: true, colour: FORMULA_BLACK });
      }
      module.exports = { buildBalanceSheet };
    }
  });

  // returns.js
  var require_returns = __commonJS({
    "returns.js"(exports, module) {
      var H = require_helpers();
      var S = require_styles();
      var C = require_constants();
      var {
        crossRef,
        sameRef,
        crossRefRange,
        colLetter,
        argb,
        cellRC,
        writeLabel,
        writeFormula,
        writeHeader,
        applyStyle
      } = H;
      var {
        FORMULA_BLACK,
        LINK_GREEN,
        DARK_BLUE,
        LIGHT_BLUE,
        TOTAL_GREY,
        ORANGE,
        FMT_DOLLAR,
        FMT_PERCENT,
        FMT_MULTIPLE,
        FMT_INTEGER
      } = S;
      var {
        SHEET_ASM,
        SHEET_SU,
        SHEET_DS,
        SHEET_CF,
        LTM_COL,
        YR1_COL,
        N_YEARS,
        ASM_EBITDA_R,
        ASM_TRANS_FEES_R,
        ASM_EXIT_YEAR_R,
        ASM_EXIT_MULT_R,
        ASM_EXIT_FEE_R,
        ASM_ENTRY_MULT_R,
        ASM_INPUT_COL,
        SU_TOT_EQ_R,
        SU_TOT_DEBT_R,
        SU_VAL_C,
        DS_TDEBT_R,
        CF_END_CASH_R,
        RET_VAL_C,
        RET_HDR_R,
        RET_ENTRY_EQ_R,
        RET_EXIT_YR_R,
        RET_EXIT_EBD_R,
        RET_EXIT_MULT_R,
        RET_EXIT_EV_R,
        RET_EXIT_DBT_R,
        RET_EXIT_CSH_R,
        RET_EXIT_FEE_R,
        RET_EXIT_EQ_R,
        RET_MOIC_R,
        RET_IRR_R,
        RET_CF_HDR_R,
        RET_CF0_R,
        RET_CF5_R,
        SEN_IRR_HDR_R,
        SEN_IRR_R1,
        SEN_IRR_C1,
        SEN_MOIC_HDR_R,
        SEN_MOIC_R1,
        SEN_MOIC_C1,
        SEN_ENTRY_MULTS,
        SEN_EXIT_MULTS,
        ASM_FIN_FEES_R
      } = C;
      function buildReturns(ws) {
        ws.getColumn(1).width = 2;
        ws.getColumn(2).width = 52;
        ws.getColumn(3).width = 16;
        [4, 5, 6, 7, 8, 9, 10].forEach((n) => ws.getColumn(n).width = 13);
        writeLabel(ws, 1, 2, "RETURNS ANALYSIS", { bold: true, size: 14, colour: DARK_BLUE });
        writeLabel(
          ws,
          2,
          2,
          "Exit equity, MOIC, IRR  |  Sensitivity: entry multiple vs exit multiple  |  ",
          { italic: true, size: 9, colour: "808080" }
        );
        writeHeader(ws, RET_HDR_R, 2, "EXIT RETURNS SUMMARY", 3);
        const _entryEq = crossRef(SHEET_SU, SU_TOT_EQ_R, SU_VAL_C);
        const _exitYr = crossRef(SHEET_ASM, ASM_EXIT_YEAR_R, ASM_INPUT_COL);
        const _exitXm = crossRef(SHEET_ASM, ASM_EXIT_MULT_R, ASM_INPUT_COL);
        const _exitXf = crossRef(SHEET_ASM, ASM_EXIT_FEE_R, ASM_INPUT_COL);
        const _ebitdaRng = crossRefRange(SHEET_ASM, ASM_EBITDA_R, YR1_COL, YR1_COL + N_YEARS - 1);
        const _debtRng = crossRefRange(SHEET_DS, DS_TDEBT_R, YR1_COL, YR1_COL + N_YEARS - 1);
        const _cashRng = crossRefRange(SHEET_CF, CF_END_CASH_R, YR1_COL, YR1_COL + N_YEARS - 1);
        writeLabel(ws, RET_ENTRY_EQ_R, 2, "Entry Equity Invested  ($M)  [from Sources & Uses]", { bold: true });
        writeFormula(ws, RET_ENTRY_EQ_R, RET_VAL_C, `=${_entryEq}`, FMT_DOLLAR, { bold: true, colour: LINK_GREEN });
        writeLabel(ws, RET_EXIT_YR_R, 2, "Exit Year  [from Assumptions]");
        writeFormula(ws, RET_EXIT_YR_R, RET_VAL_C, `=${_exitYr}`, FMT_INTEGER, { colour: LINK_GREEN });
        writeLabel(ws, RET_EXIT_EBD_R, 2, "Exit EBITDA  ($M)  [INDEX from Assumptions Y1-5]");
        writeFormula(ws, RET_EXIT_EBD_R, RET_VAL_C, `=INDEX(${_ebitdaRng},${_exitYr})`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, RET_EXIT_MULT_R, 2, "Exit EV / EBITDA Multiple  (\xD7)  [from Assumptions]");
        writeFormula(ws, RET_EXIT_MULT_R, RET_VAL_C, `=${_exitXm}`, FMT_MULTIPLE, { colour: LINK_GREEN });
        writeLabel(ws, RET_EXIT_EV_R, 2, "Exit Enterprise Value  ($M)  [= Exit EBITDA \xD7 Exit Multiple]", { bold: true });
        writeFormula(
          ws,
          RET_EXIT_EV_R,
          RET_VAL_C,
          `=${sameRef(RET_EXIT_EBD_R, RET_VAL_C)}*${sameRef(RET_EXIT_MULT_R, RET_VAL_C)}`,
          FMT_DOLLAR,
          { bold: true, colour: FORMULA_BLACK }
        );
        writeLabel(ws, RET_EXIT_DBT_R, 2, "(-) Exit Total Debt  ($M)  [INDEX from Debt Schedule]");
        writeFormula(ws, RET_EXIT_DBT_R, RET_VAL_C, `=INDEX(${_debtRng},${_exitYr})`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, RET_EXIT_CSH_R, 2, "(+) Exit Cash Balance  ($M)  [INDEX from Cash Flow]");
        writeFormula(ws, RET_EXIT_CSH_R, RET_VAL_C, `=INDEX(${_cashRng},${_exitYr})`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, RET_EXIT_FEE_R, 2, "(-) Exit Fee  ($M)  [= Exit EV \xD7 Exit Fee %]");
        writeFormula(ws, RET_EXIT_FEE_R, RET_VAL_C, `=${sameRef(RET_EXIT_EV_R, RET_VAL_C)}*${_exitXf}`, FMT_DOLLAR, { colour: LINK_GREEN });
        writeLabel(ws, RET_EXIT_EQ_R, 2, "Exit Equity  ($M)  [= EV \u2212 Fee \u2212 Debt + Cash]", { bold: true });
        writeFormula(
          ws,
          RET_EXIT_EQ_R,
          RET_VAL_C,
          `=MAX(${sameRef(RET_EXIT_EV_R, RET_VAL_C)}-${sameRef(RET_EXIT_FEE_R, RET_VAL_C)}-${sameRef(RET_EXIT_DBT_R, RET_VAL_C)}+${sameRef(RET_EXIT_CSH_R, RET_VAL_C)},0)`,
          FMT_DOLLAR,
          { bold: true, colour: FORMULA_BLACK }
        );
        writeLabel(ws, RET_MOIC_R, 2, "MOIC  (Multiple on Invested Capital)", { bold: true });
        writeFormula(
          ws,
          RET_MOIC_R,
          RET_VAL_C,
          `=IFERROR(${sameRef(RET_EXIT_EQ_R, RET_VAL_C)}/${sameRef(RET_ENTRY_EQ_R, RET_VAL_C)},0)`,
          FMT_MULTIPLE,
          { bold: true, bg: ORANGE, colour: FORMULA_BLACK }
        );
        writeLabel(ws, RET_IRR_R, 2, "IRR  (Internal Rate of Return)", { bold: true });
        writeFormula(
          ws,
          RET_IRR_R,
          RET_VAL_C,
          `=IFERROR(IRR(${sameRef(RET_CF0_R, RET_VAL_C)}:${sameRef(RET_CF5_R, RET_VAL_C)},0.20),0)`,
          FMT_PERCENT,
          { bold: true, bg: ORANGE, colour: FORMULA_BLACK }
        );
        writeHeader(
          ws,
          RET_CF_HDR_R,
          2,
          "IRR HELPER CASH FLOWS  (Year 0 = equity in; exit year = equity out)",
          3
        );
        writeLabel(ws, RET_CF0_R, 2, "Year 0:  \u2212Entry Equity  (fund invests)");
        writeFormula(ws, RET_CF0_R, RET_VAL_C, `=-${sameRef(RET_ENTRY_EQ_R, RET_VAL_C)}`, FMT_DOLLAR, { colour: FORMULA_BLACK });
        for (let yr = 1; yr <= N_YEARS; yr++) {
          const row = RET_CF0_R + yr;
          writeLabel(ws, row, 2, `Year ${yr}:  + Exit Equity  (only if exit year = ${yr})`);
          writeFormula(
            ws,
            row,
            RET_VAL_C,
            `=IF(${_exitYr}=${yr},${sameRef(RET_EXIT_EQ_R, RET_VAL_C)},0)`,
            FMT_DOLLAR,
            { colour: LINK_GREEN }
          );
        }
        buildSensitivityTables(ws, _ebitdaRng, _debtRng, _cashRng, _exitYr, _exitXf);
      }
      function buildSensitivityTables(ws, ebitdaRng, debtRng, cashRng, exitYr, exitXf) {
        const _ltmEbitda = crossRef(SHEET_ASM, ASM_EBITDA_R, LTM_COL);
        const _tf = crossRef(SHEET_ASM, ASM_TRANS_FEES_R, ASM_INPUT_COL);
        const _ff = crossRef(SHEET_ASM, ASM_FIN_FEES_R, ASM_INPUT_COL);
        const _debtClose = crossRef(SHEET_SU, SU_TOT_DEBT_R, SU_VAL_C);
        const _asmEntryRef = `'${SHEET_ASM}'!$${colLetter(ASM_INPUT_COL)}$${ASM_ENTRY_MULT_R}`;
        const _asmExitRef = `'${SHEET_ASM}'!$${colLetter(ASM_INPUT_COL)}$${ASM_EXIT_MULT_R}`;
        const tables = [
          [SEN_IRR_HDR_R, SEN_IRR_R1, SEN_IRR_C1, "IRR SENSITIVITY  (%)", FMT_PERCENT, true],
          [SEN_MOIC_HDR_R, SEN_MOIC_R1, SEN_MOIC_C1, "MOIC SENSITIVITY  (\xD7)", FMT_MULTIPLE, false]
        ];
        for (const [hdrRow, r1Row, c1Col, title, fmt, isIrr] of tables) {
          writeHeader(ws, hdrRow, 2, `${title}  |  Entry Multiple \u2193  vs  Exit Multiple \u2192`, 9);
          writeLabel(
            ws,
            hdrRow + 1,
            2,
            "Debt held constant  |  Orange = base case  |  All else equal \u2014 only multiples vary",
            { italic: true, size: 9, colour: "404040" }
          );
          writeLabel(ws, hdrRow + 2, 2, "Entry \u2193 / Exit \u2192", { bold: true, bg: TOTAL_GREY, align: "center" });
          for (let j = 0; j < SEN_EXIT_MULTS.length; j++) {
            const cell = cellRC(ws, hdrRow + 2, c1Col + j);
            cell.value = SEN_EXIT_MULTS[j];
            applyStyle(cell, { bold: true, bg: LIGHT_BLUE, align: "center", fmt: '0"\xD7"' });
          }
          for (let i = 0; i < SEN_ENTRY_MULTS.length; i++) {
            const row = r1Row + i;
            const em = SEN_ENTRY_MULTS[i];
            const hCell = cellRC(ws, row, 3);
            hCell.value = em;
            applyStyle(hCell, { bold: true, bg: LIGHT_BLUE, align: "center", fmt: '0"\xD7"' });
            for (let j = 0; j < SEN_EXIT_MULTS.length; j++) {
              const col = c1Col + j;
              const xm = SEN_EXIT_MULTS[j];
              const _entryEqSen = `MAX(${_ltmEbitda}*${em}*(1+${_tf})+${_debtClose}*${_ff}-${_debtClose},1)`;
              const _exitEbd = `INDEX(${ebitdaRng},${exitYr})`;
              const _exitDbt = `INDEX(${debtRng},${exitYr})`;
              const _exitCsh = `INDEX(${cashRng},${exitYr})`;
              const _exitEqSen = `MAX(${_exitEbd}*${xm}*(1-${exitXf})-${_exitDbt}+${_exitCsh},0)`;
              const formula = isIrr ? `=IFERROR((${_exitEqSen}/${_entryEqSen})^(1/MAX(${exitYr},1))-1,"N/A")` : `=IFERROR(${_exitEqSen}/${_entryEqSen},0)`;
              writeFormula(ws, row, col, formula, fmt, { bold: false, bg: null, colour: LINK_GREEN });
            }
          }
          const nEntry = SEN_ENTRY_MULTS.length;
          const nExit = SEN_EXIT_MULTS.length;
          const dataRange = `${colLetter(c1Col)}${r1Row}:${colLetter(c1Col + nExit - 1)}${r1Row + nEntry - 1}`;
          const cfFormula = `AND($C${r1Row}=${_asmEntryRef},${colLetter(c1Col)}$${hdrRow + 2}=${_asmExitRef})`;
          ws.addConditionalFormatting({
            ref: dataRange,
            rules: [{
              type: "expression",
              formulae: [cfFormula],
              style: {
                fill: { type: "pattern", pattern: "solid", bgColor: { argb: argb(ORANGE) } },
                font: { bold: true }
              }
            }]
          });
        }
      }
      module.exports = { buildReturns };
    }
  });

  // buildModel.js
  var require_buildModel = __commonJS({
    "buildModel.js"(exports, module) {
      var { buildAssumptions } = require_assumptions();
      var { buildSourcesUses } = require_sourcesUses();
      var { buildIncomeStatement } = require_incomeStmt();
      var { buildDebtSchedule } = require_debtSchedule();
      var { buildCashFlow } = require_cashFlow();
      var { buildBalanceSheet } = require_balanceSheet();
      var { buildReturns } = require_returns();
      var { buildAllNamedRanges } = require_constants();
      var { registerNamedRanges } = require_helpers();
      function buildWorkbook(ExcelJS, inputs = {}) {
        const wb = new ExcelJS.Workbook();
        wb.creator = "LBO Model";
        wb.created = /* @__PURE__ */ new Date();
        const wsAsm = wb.addWorksheet("Assumptions");
        const wsSu = wb.addWorksheet("Sources_Uses");
        const wsIs = wb.addWorksheet("Income_Stmt");
        const wsDs = wb.addWorksheet("Debt_Schedule");
        const wsCf = wb.addWorksheet("Cash_Flow");
        const wsBs = wb.addWorksheet("Balance_Sheet");
        const wsRet = wb.addWorksheet("Returns");
        buildAssumptions(wsAsm, inputs);
        buildSourcesUses(wsSu);
        buildIncomeStatement(wsIs);
        buildDebtSchedule(wsDs);
        buildCashFlow(wsCf);
        buildBalanceSheet(wsBs);
        buildReturns(wsRet);
        registerNamedRanges(wb, buildAllNamedRanges());
        wb.calcProperties.fullCalcOnLoad = true;
        return wb;
      }
      module.exports = { buildWorkbook };
    }
  });

  // browser-entry.js
  var require_browser_entry = __commonJS({
    "browser-entry.js"(exports, module) {
      var { buildWorkbook } = require_buildModel();
      var MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      function getExcelJS() {
        if (typeof window === "undefined" || !window.ExcelJS) {
          throw new Error(
            'ExcelJS not found. Load the ExcelJS browser build before lbo-model.js (e.g. <script src="exceljs.min.js"><\/script>).'
          );
        }
        return window.ExcelJS;
      }
      function build(inputs = {}) {
        return buildWorkbook(getExcelJS(), inputs);
      }
      async function buildBlob(inputs = {}) {
        const wb = build(inputs);
        const buffer = await wb.xlsx.writeBuffer();
        return new Blob([buffer], { type: MIME });
      }
      async function downloadModel(inputs = {}, filename = "lbo_model.xlsx") {
        const blob = await buildBlob(inputs);
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      }
      module.exports = { buildWorkbook: build, buildBlob, downloadModel };
    }
  });
  return require_browser_entry();
})();
