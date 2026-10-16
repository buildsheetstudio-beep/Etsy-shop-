'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Monthly Funding Planner'];
const S = "'Monthly Funding Planner'";
const FS = "'Fund Setup & Goals'";
const CL = "'Contribution Log'";
const NC = 18; // A-R

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['MONTHLY FUNDING PLANNER']] });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.softPeach), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Plan how your available monthly savings will be distributed across active funds. Funding Order is a planning aid only — not a financial recommendation.']] });

  // ── Planning Controls (rows 3-8) ─────────────────────────────────────────
  const ctls = [
    ['Planning Month','August'],
    ['Planning Year',2026],
    ['Owner / Household','All Owners'],
    ['Available Monthly Savings',3000],
    ['Paycheck Frequency','Biweekly'],
    ['Expected Paychecks This Month',2],
  ];
  ctls.forEach(([lbl, val], i) => {
    const row = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID, row-1, row, 0, 3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row-1, row, 3, 8), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row-1, row, 0, 3), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row-1, row, 3, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    vals.push({ range: `${S}!A${row}`, values: [[lbl]] });
    vals.push({ range: `${S}!D${row}`, values: [[val]] });
  });
  // Named control references for formulas
  // D3=Month, D4=Year, D5=Owner, D6=AvailSavings, D7=PayFreq, D8=PaychecksThisMonth
  fmt.push({ setDataValidation: { range: gridRange(SID, 2, 3, 3, 4), rule: { condition: { type: 'ONE_OF_LIST', values: ['January','February','March','April','May','June','July','August','September','October','November','December'].map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });
  fmt.push({ setDataValidation: { range: gridRange(SID, 4, 5, 3, 4), rule: { condition: { type: 'ONE_OF_LIST', values: ['All Owners','Self','Person 1','Person 2','Joint / Household'].map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });
  fmt.push({ setDataValidation: { range: gridRange(SID, 6, 7, 3, 4), rule: { condition: { type: 'ONE_OF_LIST', values: ['Weekly','Biweekly','Semi-Monthly','Monthly'].map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });

  // ── Summary KPI Cards (cols J-R, rows 3-7) ────────────────────────────────
  const kpis = [
    { label: 'Available Monthly Savings', col: 9,  fml: `=$D$6` },
    { label: 'Total Planned Funding',     col: 12, fml: `=IFERROR(SUM($L$12:$L$511),0)` },
    { label: 'Funding Gap / Surplus',     col: 15, fml: `=IFERROR($D$6-SUM($L$12:$L$511),0)` },
  ];
  kpis.forEach(k => {
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 4, k.col, k.col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 4, 7, k.col, k.col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 2, 4, k.col, k.col+3), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 7, k.col, k.col+3), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 22, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${colL(k.col)}3`, values: [[k.label]] });
    vals.push({ range: `${S}!${colL(k.col)}5`, values: [[k.fml]] });
  });
  // Currency format on KPI cells
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 7, 9, 18), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Spacer row 9 ───────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 8, 9, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 6 }, fields: 'pixelSize' } });

  // ── Funding Order Note (row 10) ────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 9, 10, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 9, 10, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.info), textFormat: { fontSize: 9, fontFamily: 'Arial' } }, }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  vals.push({ range: `${S}!A10`, values: [['PLANNING NOTE: Funding Order is calculated from Priority (45%), Time Urgency (35%), and Funding Need (20%). It is a planning aid to help you decide how to allocate limited savings — it does not tell you what you should do. You decide.']] });

  // ── Table headers (row 11, 0-indexed 10) ──────────────────────────────────
  const THDRS = ['Fund ID','Fund Name','Owner','Priority','Funding Method','Current Balance','Amount Remaining','Target Date','Monthly Needed','Selected Planned Funding','Funding Order','User Planned Amount','Difference vs Plan','% of Available','Status','Days Until Goal','Monthly Needed (Alt.)','Notes'];
  fmt.push({ repeatCell: { range: gridRange(SID, 10, 11, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy,verticalAlignment)' } });
  vals.push({ range: `${S}!A11`, values: [THDRS] });

  // ── Allocation Table (rows 12:511) ─────────────────────────────────────────
  const NFUNDS = 500;
  // Lookup formulas from Fund Setup
  const aFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(IF(${FS}!$B$${8+i}="","",${FS}!$A$${8+i}),"")`]);
  const bFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$B$${8+i},"")`]);
  const cFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$D$${8+i},"")`]);
  const dFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$E$${8+i},"")`]);
  const eFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$F$${8+i},"")`]);
  const fFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$I$${8+i},0)`]);
  const gFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$J$${8+i},0)`]);
  const hFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$L$${8+i},"")`]);
  const iFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$O$${8+i},0)`]);
  const jFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$S$${8+i},0)`]);
  const kFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$W$${8+i},"")`]);
  // L = User Planned Amount (editable — pre-populate with suggested plan)
  const lFmls  = Array.from({ length: NFUNDS }, (_, i) => { const r=12+i; return [`=IFERROR(IF(A${r}="","",J${r}),0)`]; });
  const mFmls  = Array.from({ length: NFUNDS }, (_, i) => { const r=12+i; return [`=IFERROR(L${r}-J${r},0)`]; });
  const nFmls  = Array.from({ length: NFUNDS }, (_, i) => { const r=12+i; return [`=IFERROR(IF($D$6=0,0,L${r}/$D$6),0)`]; });
  const oFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$V$${8+i},"")`]);
  const pFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$U$${8+i},"")`]);
  const qFmls  = Array.from({ length: NFUNDS }, (_, i) => [`=IFERROR(${FS}!$O$${8+i},0)`]);

  vals.push({ range: `${S}!A12:A511`, values: aFmls });
  vals.push({ range: `${S}!B12:B511`, values: bFmls });
  vals.push({ range: `${S}!C12:C511`, values: cFmls });
  vals.push({ range: `${S}!D12:D511`, values: dFmls });
  vals.push({ range: `${S}!E12:E511`, values: eFmls });
  vals.push({ range: `${S}!F12:F511`, values: fFmls });
  vals.push({ range: `${S}!G12:G511`, values: gFmls });
  vals.push({ range: `${S}!H12:H511`, values: hFmls });
  vals.push({ range: `${S}!I12:I511`, values: iFmls });
  vals.push({ range: `${S}!J12:J511`, values: jFmls });
  vals.push({ range: `${S}!K12:K511`, values: kFmls });
  vals.push({ range: `${S}!L12:L511`, values: lFmls });
  vals.push({ range: `${S}!M12:M511`, values: mFmls });
  vals.push({ range: `${S}!N12:N511`, values: nFmls });
  vals.push({ range: `${S}!O12:O511`, values: oFmls });
  vals.push({ range: `${S}!P12:P511`, values: pFmls });
  vals.push({ range: `${S}!Q12:Q511`, values: qFmls });

  // Row styling
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 11+NFUNDS, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 8, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  for (let r = 0; r < NFUNDS; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, 11+r, 12+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Formula vs input tint
  [0,1,2,3,4,5,6,7,8,9,10,12,13,14,15,16].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 11, 11+NFUNDS, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });
  [11,17].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 11, 11+NFUNDS, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Number formats for table
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 11+NFUNDS, 5, 10),  cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 11+NFUNDS, 7, 8),   cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 11+NFUNDS, 11, 14), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 11+NFUNDS, 13, 14), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Paycheck Allocation Helper (row 515+) ─────────────────────────────────
  const pRow = 514;
  fmt.push({ mergeCells: { range: gridRange(SID, pRow, pRow+1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, pRow, pRow+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.dustyBlue), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial' } }, }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  vals.push({ range: `${S}!A${pRow+1}`, values: [['PAYCHECK ALLOCATION HELPER']] });

  const pHdrs = ['Paycheck #','Pay Date','Available Per Paycheck','Fund','Planned Amount','Remaining After'];
  fmt.push({ repeatCell: { range: gridRange(SID, pRow+1, pRow+2, 0, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A${pRow+2}`, values: [pHdrs] });

  // Two paycheck rows as examples
  const paycheckRows = [
    ['1','',`=IFERROR($D$6/MAX(1,$D$8),0)`,'','',`=IFERROR(C${pRow+3}-E${pRow+3},0)`],
    ['2','',`=IFERROR($D$6/MAX(1,$D$8),0)`,'','',`=IFERROR(C${pRow+4}-E${pRow+4},0)`],
  ];
  paycheckRows.forEach((row, i) => {
    vals.push({ range: `${S}!A${pRow+3+i}:F${pRow+3+i}`, values: [row] });
    fmt.push({ repeatCell: { range: gridRange(SID, pRow+2+i, pRow+3+i, 0, 6), cell: { userEnteredFormat: { backgroundColor: i%2===0 ? hex(C.panel) : hex(C.altRow), textFormat: { fontSize: 8 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  });
  fmt.push({ repeatCell: { range: gridRange(SID, pRow+2, pRow+5, 2, 3), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Freeze rows 1-6 ────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 11 } }, fields: 'gridProperties.frozenRowCount' } });

  // Row/col heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 11+NFUNDS }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });

  const WIDTHS = [70,160,90,70,110,80,80,90,80,80,100,80,80,60,90,60,80,160];
  WIDTHS.forEach((w,i) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } }));

  await batchUpdate(id, fmt, 'mp-fmt');
  await valuesBatchUpdate(id, vals, 'mp-vals');
  console.log('✓ Monthly Funding Planner complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
