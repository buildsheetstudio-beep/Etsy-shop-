'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Savings Forecast'];
const S = "'Savings Forecast'";
const SETUP = "'Fund Setup & Goals'";

(async () => {
  const fmt = [];
  const vals = [];

  // ── Tab header ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 18), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['SAVINGS FORECAST']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, 18), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 18), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A2`, values: [['Project when each fund reaches its goal based on your current saving rate']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, 18), cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.textMid) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // ── Scenario Controls (rows 4-6) ────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 3), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A4`, values: [['SCENARIO CONTROLS']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 3), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const controls = [
    ['Base Monthly Savings', 1200, 'D5'],
    ['Optimistic Boost (%)', 25, 'D6'],
    ['Conservative Cut (%)', 20, 'D7'],
  ];
  controls.forEach(([label, def, addr], i) => {
    vals.push({ range: `${S}!A${5+i}`, values: [[label]] });
    vals.push({ range: `${S}!${addr}`, values: [[def]] });
    fmt.push({ repeatCell: { range: gridRange(SID, 4+i, 5+i, 0, 3), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9 } } }, fields: 'userEnteredFormat.textFormat' } });
  });

  fmt.push({ repeatCell: { range: gridRange(SID, 4, 7, 3, 4), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), numberFormat: { type: 'NUMBER', pattern: '#,##0.##' } } }, fields: 'userEnteredFormat(backgroundColor,numberFormat)' } });

  // ── Column headers (row 9) ───────────────────────────────────────────────────
  const HEADERS = [
    'Fund ID','Fund Name','Category','Goal ($)','Current Balance ($)',
    'Remaining ($)','Monthly Rate (Base)','Months Left (Base)','Goal Date (Base)',
    'Monthly Rate (Opt.)','Months Left (Opt.)','Goal Date (Opt.)',
    'Monthly Rate (Cons.)','Months Left (Cons.)','Goal Date (Cons.)',
    '% Funded','Status','Notes',
  ];
  vals.push({ range: `${S}!A9`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, 8, 9, 0, HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  // ── Fund rows (rows 10–35, 26 funds) ────────────────────────────────────────
  const FUND_IDS = Array.from({ length: 26 }, (_, i) => `FUND-${String(i+1).padStart(3,'0')}`);
  const rowData = [];
  const fmtRows = [];

  FUND_IDS.forEach((fid, i) => {
    const r = 10 + i;
    // Pull data from Fund Setup via VLOOKUP (cols A=FundID, B=Name, C=Cat, D=Goal, I=Current Balance, V=Status, AC=Notes)
    // Fund Setup columns: A=ID, B=Name, C=Category, D=Owner, E=Priority, F=Method, G=GoalAmt, H=StartingBal, I=CurrentBal, J=Required Monthly, V=Status, Y=Notes
    const vlookup = (col) => `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$Z$33,${col},0),"")`;
    // col indices (1-based): A=1,B=2,C=3,D=4,E=5,F=6,G=7,H=8,I=9,J=10,...V=22,...Y=25
    const goalAmt = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$AB$33,7,0),0)`;
    const currBal = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$AB$33,9,0),0)`;
    const statusV = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$AB$33,22,0),"")`;
    const nameV   = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$AB$33,2,0),"")`;
    const catV    = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$AB$33,3,0),"")`;
    const notesV  = `IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$AB$33,25,0),"")`;

    // Remaining = Goal - CurrentBalance (floored at 0)
    const remaining = `MAX(0,${goalAmt}-(${currBal}))`;

    // Base monthly rate: allocate proportionally by remaining balance
    // Simplified: use D$5 * (remaining / total remaining across all funds)
    // For simplicity, use a direct formula: user's base savings * (remaining / SUM of all remainings)
    // But that's circular if all funds use same sheet... use fixed per-fund rate instead:
    // Base = $D$5 / 26 (equal split) — user can override
    const baseRate  = `IFERROR($D$5/COUNTIF(${SETUP}!$V$8:$V$33,"<>Goal Reached"),0)`;
    const optRate   = `IFERROR(${baseRate}*(1+$D$6/100),0)`;
    const consRate  = `IFERROR(${baseRate}*(1-$D$7/100),0)`;

    // Months left = CEILING(remaining / rate)
    const monthsBase  = `IFERROR(IF(${remaining}=0,0,CEILING(${remaining}/${baseRate},1)),999)`;
    const monthsOpt   = `IFERROR(IF(${remaining}=0,0,CEILING(${remaining}/${optRate},1)),999)`;
    const monthsCons  = `IFERROR(IF(${remaining}=0,0,CEILING(${remaining}/${consRate},1)),999)`;

    // Goal date = EDATE(TODAY(), months)
    const goalDateBase  = `IFERROR(IF(${remaining}=0,"Already Funded",EDATE(TODAY(),${monthsBase})),"")`;
    const goalDateOpt   = `IFERROR(IF(${remaining}=0,"Already Funded",EDATE(TODAY(),${monthsOpt})),"")`;
    const goalDateCons  = `IFERROR(IF(${remaining}=0,"Already Funded",EDATE(TODAY(),${monthsCons})),"")`;

    // % Funded
    const pctFunded = `IFERROR(IF(${goalAmt}=0,0,MIN(1,(${currBal})/${goalAmt})),0)`;

    rowData.push([
      `="${fid}"`,
      `=${nameV}`,
      `=${catV}`,
      `=${goalAmt}`,
      `=${currBal}`,
      `=${remaining}`,
      `=${baseRate}`,
      `=${monthsBase}`,
      `=${goalDateBase}`,
      `=${optRate}`,
      `=${monthsOpt}`,
      `=${goalDateOpt}`,
      `=${consRate}`,
      `=${monthsCons}`,
      `=${goalDateCons}`,
      `=${pctFunded}`,
      `=${statusV}`,
      `=${notesV}`,
    ]);
  });

  vals.push({ range: `${S}!A10`, values: rowData });

  // ── Number formats for data rows ────────────────────────────────────────────
  // Currency: D, E, F, G, J, M (goal, curr, remaining, rates)
  [3,4,5,6,9,12].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 9, 35, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Integer months: H, K, N
  [7,10,13].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 9, 35, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Date: I, L, O
  [8,11,14].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 9, 35, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Percent: P
  fmt.push({ repeatCell: { range: gridRange(SID, 9, 35, 15, 16), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Alternating row stripes ──────────────────────────────────────────────────
  for (let i = 0; i < 26; i++) {
    const bg = i % 2 === 0 ? C.white : C.stripeBg;
    fmt.push({ repeatCell: { range: gridRange(SID, 9+i, 10+i, 0, 18), cell: { userEnteredFormat: { backgroundColor: hex(bg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // ── Summary block (rows 38-45) ───────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 37, 38, 0, 6), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A38`, values: [['FORECAST SUMMARY']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 37, 38, 0, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const summaryRows = [
    ['Total Goal Amount', '=IFERROR(SUM(D10:D35),0)', '', 'Total Current Balance', '=IFERROR(SUM(E10:E35),0)', ''],
    ['Total Remaining', '=IFERROR(SUM(F10:F35),0)', '', 'Overall % Funded', '=IFERROR(SUM(E10:E35)/SUM(D10:D35),0)', ''],
    ['Funds Goal Reached', '=IFERROR(COUNTIF(Q10:Q35,"Goal Reached"),0)', '', 'Funds Active', '=IFERROR(COUNTIF(Q10:Q35,"Active"),0)', ''],
    ['Earliest Goal Date (Base)', '=IFERROR(MIN(IF(H10:H35<999,I10:I35)),"")', '', 'Latest Goal Date (Base)', '=IFERROR(MAX(IF(H10:H35<999,I10:I35)),"")', ''],
  ];
  vals.push({ range: `${S}!A39`, values: summaryRows });
  fmt.push({ repeatCell: { range: gridRange(SID, 38, 42, 0, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  [1,4].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 38, 40, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  fmt.push({ repeatCell: { range: gridRange(SID, 39, 40, 4, 5), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  [41, 42].forEach((row, i) => {
    [1, 4].forEach(col => {
      fmt.push({ repeatCell: { range: gridRange(SID, 38+i+2, 38+i+3, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
    });
  });
  // Bold labels
  [0,3].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 38, 42, col, col+1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat.textFormat' } });
  });

  // ── Column widths ────────────────────────────────────────────────────────────
  const colWidths = [80,160,110,90,105,90,100,85,95,100,85,95,100,85,95,70,110,140];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row heights ──────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // ── Freeze ───────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 9 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Borders on data table ────────────────────────────────────────────────────
  fmt.push({ updateBorders: { range: gridRange(SID, 8, 35, 0, 18), innerHorizontal: { style: 'SOLID', color: hex(C.borderLight) }, innerVertical: { style: 'SOLID', color: hex(C.borderLight) }, bottom: { style: 'SOLID', color: hex(C.border) }, top: { style: 'SOLID', color: hex(C.border) }, left: { style: 'SOLID', color: hex(C.border) }, right: { style: 'SOLID', color: hex(C.border) } } });

  await batchUpdate(id, fmt, 'forecast-fmt');
  await valuesBatchUpdate(id, vals, 'forecast-vals');

  console.log('✓ Savings Forecast complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
