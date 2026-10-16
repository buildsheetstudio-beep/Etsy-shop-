'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Annual Summary'];
const S = "'Annual Summary'";
const SETUP = "'Fund Setup & Goals'";
const LOG   = "'Contribution Log'";

(async () => {
  const fmt = [];
  const vals = [];

  const YEARS = [2024, 2025, 2026];

  // ── Tab header ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['ANNUAL SINKING FUNDS SUMMARY']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A2`, values: [['Year-over-year contributions, withdrawals, and fund performance across all sinking funds']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.textMid) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // ── Year selector control (row 4) ─────────────────────────────────────────────
  vals.push({ range: `${S}!A4`, values: [['Selected Year']] });
  vals.push({ range: `${S}!B4`, values: [[2026]] });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9 } } }, fields: 'userEnteredFormat.textFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  fmt.push({
    setDataValidation: {
      range: gridRange(SID, 3, 4, 1, 2),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: YEARS.map(y => ({ userEnteredValue: String(y) })) },
        showCustomUi: true, strict: false,
      },
    },
  });

  // ── Section 1: Year-over-Year Summary Table ─────────────────────────────────
  // Rows 6-9: header + years
  fmt.push({ mergeCells: { range: gridRange(SID, 5, 6, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A6`, values: [['YEAR-OVER-YEAR SUMMARY']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const YOY_HEADERS = ['Year','Total Contributions','Total Withdrawals','Net Saved','Adjustments','Refunds','# Transactions','Avg Monthly Contrib','Funds Active','Funds Completed'];
  vals.push({ range: `${S}!A7`, values: [YOY_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, YOY_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  YEARS.forEach((yr, yi) => {
    const r = 8 + yi;
    const yearCond = `YEAR(${LOG}!$B$8:$B$5007)=${yr}`;
    const contrib    = `SUMPRODUCT((${yearCond})*(${LOG}!$H$8:$H$5007="Contribution")*${LOG}!$J$8:$J$5007)`;
    const withdraw   = `SUMPRODUCT((${yearCond})*(${LOG}!$H$8:$H$5007="Withdrawal")*${LOG}!$J$8:$J$5007)`;
    const adjInc     = `SUMPRODUCT((${yearCond})*(${LOG}!$H$8:$H$5007="Adjustment Increase")*${LOG}!$J$8:$J$5007)`;
    const adjDec     = `SUMPRODUCT((${yearCond})*(${LOG}!$H$8:$H$5007="Adjustment Decrease")*${LOG}!$J$8:$J$5007)`;
    const refund     = `SUMPRODUCT((${yearCond})*(${LOG}!$H$8:$H$5007="Refund")*${LOG}!$J$8:$J$5007)`;
    const txCount    = `SUMPRODUCT((${yearCond})*1)`;
    const netSaved   = `${contrib}+${adjInc}+${refund}-${withdraw}-${adjDec}`;
    const avgMonthly = `IFERROR((${contrib})/12,0)`;
    const fundsActive= `IFERROR(SUMPRODUCT((YEAR(${LOG}!$B$8:$B$5007)=${yr})*1/IFERROR(COUNTIF(${LOG}!$E$8:$E$5007,${LOG}!$E$8:$E$5007),1)),0)`;
    const fundsDone  = `IFERROR(COUNTIFS(${SETUP}!$V$8:$V$33,"Goal Reached"),0)`;

    vals.push({ range: `${S}!A${r}`, values: [[
      yr,
      `=IFERROR(${contrib},0)`,
      `=IFERROR(${withdraw},0)`,
      `=IFERROR(${netSaved},0)`,
      `=IFERROR(${adjInc}-${adjDec},0)`,
      `=IFERROR(${refund},0)`,
      `=IFERROR(${txCount},0)`,
      `=IFERROR(${avgMonthly},0)`,
      `=IFERROR(SUMPRODUCT((${yearCond})*(${LOG}!$H$8:$H$5007<>"")/IFERROR(COUNTIF(${LOG}!$E$8:$E$5007,${LOG}!$E$8:$E$5007),1)),0)`,
      `=IFERROR(${fundsDone},0)`,
    ]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 0, YOY_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(yi % 2 === 0 ? C.white : C.stripeBg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Currency for contrib/withdraw/net/adj/refund/avg cols
  [1,2,3,4,5,7].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 11, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Section 2: Monthly Breakdown for Selected Year ──────────────────────────
  const SEC2_ROW = 12;
  fmt.push({ mergeCells: { range: gridRange(SID, SEC2_ROW, SEC2_ROW+1, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!${colL(0)}${SEC2_ROW+1}`, values: [['MONTHLY BREAKDOWN — SELECTED YEAR']] });
  fmt.push({ repeatCell: { range: gridRange(SID, SEC2_ROW, SEC2_ROW+1, 0, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const MONTH_HEADERS = ['Month','Contributions','Withdrawals','Net','Transfer In','Transfer Out','# Transactions','Running Total'];
  vals.push({ range: `${S}!A${SEC2_ROW+2}`, values: [MONTH_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, SEC2_ROW+1, SEC2_ROW+2, 0, MONTH_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  MONTH_NAMES.forEach((mname, mi) => {
    const r = SEC2_ROW + 3 + mi;
    const m = mi + 1;
    const yrRef = `$B$4`;
    const dateCond = `(YEAR(${LOG}!$B$8:$B$5007)=${yrRef})*(MONTH(${LOG}!$B$8:$B$5007)=${m})`;
    const contrib  = `SUMPRODUCT((${dateCond})*(${LOG}!$H$8:$H$5007="Contribution")*${LOG}!$J$8:$J$5007)`;
    const withdraw = `SUMPRODUCT((${dateCond})*(${LOG}!$H$8:$H$5007="Withdrawal")*${LOG}!$J$8:$J$5007)`;
    const tIn      = `SUMPRODUCT((${dateCond})*(${LOG}!$H$8:$H$5007="Transfer In")*${LOG}!$J$8:$J$5007)`;
    const tOut     = `SUMPRODUCT((${dateCond})*(${LOG}!$H$8:$H$5007="Transfer Out")*${LOG}!$J$8:$J$5007)`;
    const txCnt    = `SUMPRODUCT((${dateCond})*1)`;
    const net      = `${contrib}-${withdraw}`;

    vals.push({ range: `${S}!A${r}`, values: [[
      mname,
      `=IFERROR(${contrib},0)`,
      `=IFERROR(${withdraw},0)`,
      `=IFERROR(${net},0)`,
      `=IFERROR(${tIn},0)`,
      `=IFERROR(${tOut},0)`,
      `=IFERROR(${txCnt},0)`,
      mi === 0
        ? `=IFERROR(B${r},0)`
        : `=IFERROR(H${r-1}+D${r},0)`,
    ]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 0, MONTH_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(mi % 2 === 0 ? C.white : C.stripeBg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Totals row
  const TOTAL_R = SEC2_ROW + 15;
  vals.push({ range: `${S}!A${TOTAL_R}`, values: [[
    'TOTAL',
    `=SUM(B${SEC2_ROW+3}:B${TOTAL_R-1})`,
    `=SUM(C${SEC2_ROW+3}:C${TOTAL_R-1})`,
    `=SUM(D${SEC2_ROW+3}:D${TOTAL_R-1})`,
    `=SUM(E${SEC2_ROW+3}:E${TOTAL_R-1})`,
    `=SUM(F${SEC2_ROW+3}:F${TOTAL_R-1})`,
    `=SUM(G${SEC2_ROW+3}:G${TOTAL_R-1})`,
    '',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID, TOTAL_R-1, TOTAL_R, 0, MONTH_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Currency formats for monthly section
  [1,2,3,4,5,7].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, SEC2_ROW+2, TOTAL_R, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Section 3: Per-Fund Annual Summary ──────────────────────────────────────
  const SEC3_ROW = TOTAL_R + 2;
  fmt.push({ mergeCells: { range: gridRange(SID, SEC3_ROW, SEC3_ROW+1, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!${colL(0)}${SEC3_ROW+1}`, values: [['PER-FUND ANNUAL BREAKDOWN — SELECTED YEAR']] });
  fmt.push({ repeatCell: { range: gridRange(SID, SEC3_ROW, SEC3_ROW+1, 0, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const FUND_HEADERS = ['Fund ID','Fund Name','Goal ($)','Current Balance ($)','% Funded','Annual Contributions','Annual Withdrawals','Net This Year','Status'];
  vals.push({ range: `${S}!A${SEC3_ROW+2}`, values: [FUND_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, SEC3_ROW+1, SEC3_ROW+2, 0, FUND_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  const FUND_IDS = Array.from({ length: 26 }, (_, i) => `FUND-${String(i+1).padStart(3,'0')}`);
  FUND_IDS.forEach((fid, fi) => {
    const r = SEC3_ROW + 3 + fi;
    const yrRef = `$B$4`;
    const fundCond = `(${LOG}!$E$8:$E$5007="${fid}")*(YEAR(${LOG}!$B$8:$B$5007)=${yrRef})`;
    const annualContrib  = `SUMPRODUCT((${fundCond})*(${LOG}!$H$8:$H$5007="Contribution")*${LOG}!$J$8:$J$5007)`;
    const annualWithdraw = `SUMPRODUCT((${fundCond})*(${LOG}!$H$8:$H$5007="Withdrawal")*${LOG}!$J$8:$J$5007)`;
    const goalAmt   = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$G$33,7,0),0)`;
    const currBal   = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$I$33,9,0),0)`;
    const pctFunded = `IFERROR(MIN(1,(${currBal})/(${goalAmt})),0)`;
    const statusV   = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$V$33,22,0),"")`;
    const nameV     = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$B$33,2,0),"")`;

    vals.push({ range: `${S}!A${r}`, values: [[
      fid,
      `=${nameV}`,
      `=${goalAmt}`,
      `=${currBal}`,
      `=${pctFunded}`,
      `=IFERROR(${annualContrib},0)`,
      `=IFERROR(${annualWithdraw},0)`,
      `=IFERROR(${annualContrib}-${annualWithdraw},0)`,
      `=${statusV}`,
    ]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 0, FUND_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(fi % 2 === 0 ? C.white : C.stripeBg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Currency/percent for fund section
  [2,3,5,6,7].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, SEC3_ROW+2, SEC3_ROW+29, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  fmt.push({ repeatCell: { range: gridRange(SID, SEC3_ROW+2, SEC3_ROW+29, 4, 5), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Column widths ────────────────────────────────────────────────────────────
  const colWidths = [80,160,90,110,80,110,110,90,110,90,90,90,90,90,90,90];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row heights ──────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Borders on main table ─────────────────────────────────────────────────────
  fmt.push({ updateBorders: { range: gridRange(SID, 6, 11, 0, YOY_HEADERS.length), innerHorizontal: { style: 'SOLID', color: hex(C.borderLight) }, innerVertical: { style: 'SOLID', color: hex(C.borderLight) }, bottom: { style: 'SOLID', color: hex(C.border) }, top: { style: 'SOLID', color: hex(C.border) }, left: { style: 'SOLID', color: hex(C.border) }, right: { style: 'SOLID', color: hex(C.border) } } });

  await batchUpdate(id, fmt, 'annual-fmt');
  await valuesBatchUpdate(id, vals, 'annual-vals');

  console.log('✓ Annual Summary complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
