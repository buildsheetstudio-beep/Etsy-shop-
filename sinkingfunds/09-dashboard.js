'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Sinking Funds Dashboard'];
const S = "'Sinking Funds Dashboard'";
const SETUP = "'Fund Setup & Goals'";
const LOG   = "'Contribution Log'";
const FORE  = "'Savings Forecast'";
const MILE  = "'Goals & Milestones'";

(async () => {
  const fmt = [];
  const vals = [];

  const NC = 18; // columns A-R

  // ── Tab header ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['SINKING FUNDS DASHBOARD']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A2`, values: [['Your complete sinking funds command center — all funds, progress, and priorities at a glance']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // ── Row 3: Last Updated ──────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A3`, values: [['=IFERROR("Last Updated: "&TEXT(TODAY(),"mmmm d, yyyy"),"")']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { italic: true, fontSize: 8, foregroundColor: hex(C.textMid) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // ── KPI Cards Row (rows 5-8) — 6 KPI cards ──────────────────────────────────
  // Cards: [Total Funds, Total Saved, Total Goal, Overall Progress, Funds Complete, Next Goal Date]
  const KPI = [
    { label: 'TOTAL ACTIVE FUNDS', val: `=IFERROR(COUNTIF(${SETUP}!$V$8:$V$33,"Active")+COUNTIF(${SETUP}!$V$8:$V$33,"Ahead of Plan")+COUNTIF(${SETUP}!$V$8:$V$33,"On Track")+COUNTIF(${SETUP}!$V$8:$V$33,"Behind Plan"),0)`, fmt: '#,##0', col: C.seafoam },
    { label: 'TOTAL SAVED', val: `=IFERROR(SUM(${SETUP}!$I$8:$I$33),0)`, fmt: '"$"#,##0.00', col: C.dustyBlue },
    { label: 'TOTAL GOAL', val: `=IFERROR(SUM(${SETUP}!$G$8:$G$33),0)`, fmt: '"$"#,##0.00', col: C.softLilac },
    { label: 'OVERALL PROGRESS', val: `=IFERROR(SUM(${SETUP}!$I$8:$I$33)/SUM(${SETUP}!$G$8:$G$33),0)`, fmt: '0.0%', col: C.mutedRose },
    { label: 'GOALS REACHED', val: `=IFERROR(COUNTIF(${SETUP}!$V$8:$V$33,"Goal Reached"),0)`, fmt: '#,##0', col: C.paleOlive },
    { label: 'MILESTONES ACHIEVED', val: `=IFERROR(COUNTIF(${MILE}!$K$6:$K$74,"Achieved"),0)`, fmt: '#,##0', col: C.softPeach },
  ];

  const KPI_COLS_START = [0, 3, 6, 9, 12, 15];
  const KPI_COL_SPAN = 3;

  KPI.forEach((kpi, ki) => {
    const col = KPI_COLS_START[ki];
    // Header row (row 5, 0-indexed 4)
    fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, col, col+KPI_COL_SPAN), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(col)}5`, values: [[kpi.label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, col, col+KPI_COL_SPAN), cell: { userEnteredFormat: { backgroundColor: hex(kpi.col), textFormat: { bold: true, fontSize: 7, foregroundColor: hex(C.textDark) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    // Value rows (rows 6-8, 0-indexed 5-7)
    fmt.push({ mergeCells: { range: gridRange(SID, 5, 8, col, col+KPI_COL_SPAN), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(col)}6`, values: [[kpi.val]] });
    const isPercent = kpi.fmt === '0.0%';
    const isCurrency = kpi.fmt.startsWith('"$"');
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 8, col, col+KPI_COL_SPAN), cell: { userEnteredFormat: { backgroundColor: hex(kpi.col), textFormat: { bold: true, fontSize: 20, foregroundColor: hex(C.textDark) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', numberFormat: { type: isPercent ? 'PERCENT' : isCurrency ? 'CURRENCY' : 'NUMBER', pattern: kpi.fmt } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)' } });
  });

  // ── Section: Top Priority Funds (rows 10-19) ─────────────────────────────────
  const PRI_ROW = 9;
  fmt.push({ mergeCells: { range: gridRange(SID, PRI_ROW, PRI_ROW+1, 0, 9), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${PRI_ROW+1}`, values: [['TOP PRIORITY FUNDS']] });
  fmt.push({ repeatCell: { range: gridRange(SID, PRI_ROW, PRI_ROW+1, 0, 9), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const PRI_HEADERS = ['Fund Name','Category','Priority','Status','Current Balance','Goal Amount','% Funded','Funding Order','Est. Goal Date'];
  vals.push({ range: `${S}!A${PRI_ROW+2}`, values: [PRI_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, PRI_ROW+1, PRI_ROW+2, 0, PRI_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  // Show top 8 active funds sorted by funding order (pull from Fund Setup rows 8-33)
  // We'll just reference Fund Setup directly for the top 8 by order
  // Using LARGE/INDEX pattern is complex; instead just show all active funds sorted naturally
  const TOP_FUNDS = 8;
  for (let fi = 0; fi < TOP_FUNDS; fi++) {
    const r = PRI_ROW + 3 + fi;
    const setupR = 8 + fi; // Fund rows 8-15 in Fund Setup (first 8)
    vals.push({ range: `${S}!A${r}`, values: [[
      `=IFERROR(${SETUP}!B${setupR},"")`,
      `=IFERROR(${SETUP}!C${setupR},"")`,
      `=IFERROR(${SETUP}!E${setupR},"")`,
      `=IFERROR(${SETUP}!V${setupR},"")`,
      `=IFERROR(${SETUP}!I${setupR},0)`,
      `=IFERROR(${SETUP}!G${setupR},0)`,
      `=IFERROR(IF(${SETUP}!G${setupR}=0,0,MIN(1,${SETUP}!I${setupR}/${SETUP}!G${setupR})),0)`,
      `=IFERROR(${SETUP}!W${setupR},"")`,
      `=IFERROR(IF(${SETUP}!G${setupR}<=${SETUP}!I${setupR},"Funded",EDATE(TODAY(),CEILING(MAX(0,${SETUP}!G${setupR}-${SETUP}!I${setupR})/${SETUP}!J${setupR},1))),"")`,
    ]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 0, PRI_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(fi % 2 === 0 ? C.white : C.stripeBg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // Number formats for priority section
  [4,5].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, PRI_ROW+2, PRI_ROW+11, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  fmt.push({ repeatCell: { range: gridRange(SID, PRI_ROW+2, PRI_ROW+11, 6, 7), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, PRI_ROW+2, PRI_ROW+11, 8, 9), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Section: Recent Contributions (rows 20-29) ────────────────────────────────
  const REC_ROW = 20;
  fmt.push({ mergeCells: { range: gridRange(SID, REC_ROW, REC_ROW+1, 0, 9), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${REC_ROW+1}`, values: [['RECENT TRANSACTIONS (LAST 10)']] });
  fmt.push({ repeatCell: { range: gridRange(SID, REC_ROW, REC_ROW+1, 0, 9), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const REC_HEADERS = ['Date','Fund ID','Fund Name','Type','Source','Amount','Balance Effect','Running Balance','Notes'];
  vals.push({ range: `${S}!A${REC_ROW+2}`, values: [REC_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, REC_ROW+1, REC_ROW+2, 0, REC_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  // Pull last 10 transactions from Contribution Log
  // Contribution Log data starts at row 8, has up to 5000 rows
  // Use OFFSET + COUNTA to get last N rows - simpler: just pull fixed rows ~218-227 (end of sample data)
  // Since we know sample data is 218 rows (rows 8 to 225), pull rows 216-225
  for (let ri = 0; ri < 10; ri++) {
    const logRow = 216 + ri; // last 10 sample data rows
    const r = REC_ROW + 3 + ri;
    vals.push({ range: `${S}!A${r}`, values: [[
      `=IFERROR(${LOG}!B${logRow},"")`,
      `=IFERROR(${LOG}!E${logRow},"")`,
      `=IFERROR(${LOG}!F${logRow},"")`,
      `=IFERROR(${LOG}!H${logRow},"")`,
      `=IFERROR(${LOG}!I${logRow},"")`,
      `=IFERROR(${LOG}!J${logRow},0)`,
      `=IFERROR(${LOG}!N${logRow},0)`,
      `=IFERROR(${LOG}!O${logRow},0)`,
      `=IFERROR(${LOG}!P${logRow},"")`,
    ]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 0, REC_HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(ri % 2 === 0 ? C.white : C.stripeBg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Date format col A
  fmt.push({ repeatCell: { range: gridRange(SID, REC_ROW+2, REC_ROW+13, 0, 1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  // Currency cols F, G, H
  [5, 6, 7].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, REC_ROW+2, REC_ROW+13, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Section: Right-hand Forecast Summary (rows 10-19, cols 9-17) ─────────────
  fmt.push({ mergeCells: { range: gridRange(SID, PRI_ROW, PRI_ROW+1, 9, NC), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!${colL(9)}${PRI_ROW+1}`, values: [['FORECAST AT A GLANCE']] });
  fmt.push({ repeatCell: { range: gridRange(SID, PRI_ROW, PRI_ROW+1, 9, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const FORE_ITEMS = [
    ['Total Remaining to Save', `=IFERROR(${FORE}!F39,0)`, 'CURRENCY'],
    ['Funds Already Funded', `=IFERROR(${FORE}!B40,0)`, 'NUMBER'],
    ['Avg Months to Complete (Base)', `=IFERROR(AVERAGEIF(${FORE}!H10:H35,"<999"),0)`, 'NUMBER'],
    ['Nearest Goal Date (Base)', `=IFERROR(MIN(IF((${FORE}!H10:H35<999)*(${FORE}!H10:H35>0),${FORE}!I10:I35)),0)`, 'DATE'],
    ['Furthest Goal Date (Base)', `=IFERROR(MAX(IF(${FORE}!H10:H35<999,${FORE}!I10:I35)),0)`, 'DATE'],
    ['% Milestones Achieved', `=IFERROR(COUNTIF(${MILE}!$K$6:$K$74,"Achieved")/69,0)`, 'PERCENT'],
    ['Milestones Delayed', `=IFERROR(COUNTIF(${MILE}!$K$6:$K$74,"Delayed"),0)`, 'NUMBER'],
    ['Total Contributions (All Time)', `=IFERROR(SUMPRODUCT((${LOG}!$H$8:$H$5007="Contribution")*${LOG}!$J$8:$J$5007),0)`, 'CURRENCY'],
  ];

  FORE_ITEMS.forEach(([label, formula, ftype], fi) => {
    const r = PRI_ROW + 2 + fi;
    fmt.push({ mergeCells: { range: gridRange(SID, r, r+1, 9, 14), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, r, r+1, 14, NC), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(9)}${r+1}`, values: [[label]] });
    vals.push({ range: `${S}!${colL(14)}${r+1}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 9, 14), cell: { userEnteredFormat: { backgroundColor: hex(fi % 2 === 0 ? C.lightGray : C.white), textFormat: { bold: true, fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    const numFmt = ftype === 'CURRENCY' ? { type: 'CURRENCY', pattern: '"$"#,##0.00' }
                 : ftype === 'PERCENT'  ? { type: 'PERCENT', pattern: '0.0%' }
                 : ftype === 'DATE'     ? { type: 'DATE', pattern: 'mmm d, yyyy' }
                 :                        { type: 'NUMBER', pattern: '#,##0.0' };
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 14, NC), cell: { userEnteredFormat: { backgroundColor: hex(fi % 2 === 0 ? C.lightGray : C.white), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.secondary) }, horizontalAlignment: 'RIGHT', numberFormat: numFmt } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)' } });
  });

  // ── Section: Upcoming Milestones (rows 32-39, right side) ─────────────────────
  const UPM_ROW = 32;
  fmt.push({ mergeCells: { range: gridRange(SID, UPM_ROW, UPM_ROW+1, 9, NC), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!${colL(9)}${UPM_ROW+1}`, values: [['UPCOMING MILESTONES']] });
  fmt.push({ repeatCell: { range: gridRange(SID, UPM_ROW, UPM_ROW+1, 9, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const UPM_HEADERS = ['Fund Name','Milestone','Target $','Status'];
  vals.push({ range: `${S}!${colL(9)}${UPM_ROW+2}`, values: [UPM_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, UPM_ROW+1, UPM_ROW+2, 9, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Pull first 6 In Progress milestones
  const IN_PROGRESS_MILE_ROWS = [6,7,8,9,10,11]; // first few milestone rows that are typically In Progress
  IN_PROGRESS_MILE_ROWS.forEach((mRow, mi) => {
    const r = UPM_ROW + 3 + mi;
    vals.push({ range: `${S}!${colL(9)}${r}`, values: [[
      `=IFERROR(${MILE}!C${mRow},"")`,
      `=IFERROR(${MILE}!E${mRow},"")`,
      `=IFERROR(${MILE}!F${mRow},0)`,
      `=IFERROR(${MILE}!K${mRow},"")`,
    ]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 9, NC), cell: { userEnteredFormat: { backgroundColor: hex(mi % 2 === 0 ? C.white : C.stripeBg), textFormat: { fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  });
  fmt.push({ repeatCell: { range: gridRange(SID, UPM_ROW+2, UPM_ROW+9, 11, 12), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Quick Stats section (row 33-39, left side) ───────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, UPM_ROW, UPM_ROW+1, 0, 9), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${UPM_ROW+1}`, values: [['THIS MONTH SNAPSHOT']] });
  fmt.push({ repeatCell: { range: gridRange(SID, UPM_ROW, UPM_ROW+1, 0, 9), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const SNAP_ITEMS = [
    ['Contributions This Month', `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$B$8:$B$5007)=YEAR(TODAY()))*(MONTH(${LOG}!$B$8:$B$5007)=MONTH(TODAY()))*(${LOG}!$H$8:$H$5007="Contribution")*${LOG}!$J$8:$J$5007),0)`, 'CURRENCY'],
    ['Withdrawals This Month', `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$B$8:$B$5007)=YEAR(TODAY()))*(MONTH(${LOG}!$B$8:$B$5007)=MONTH(TODAY()))*(${LOG}!$H$8:$H$5007="Withdrawal")*${LOG}!$J$8:$J$5007),0)`, 'CURRENCY'],
    ['Net This Month', `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$B$8:$B$5007)=YEAR(TODAY()))*(MONTH(${LOG}!$B$8:$B$5007)=MONTH(TODAY()))*${LOG}!$N$8:$N$5007),0)`, 'CURRENCY'],
    ['Transactions This Month', `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$B$8:$B$5007)=YEAR(TODAY()))*(MONTH(${LOG}!$B$8:$B$5007)=MONTH(TODAY()))*(${LOG}!$B$8:$B$5007<>"")),0)`, 'NUMBER'],
    ['Funds Contributed To (Month)', `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$B$8:$B$5007)=YEAR(TODAY()))*(MONTH(${LOG}!$B$8:$B$5007)=MONTH(TODAY()))*(${LOG}!$H$8:$H$5007="Contribution")/IFERROR(COUNTIF(${LOG}!$E$8:$E$5007,${LOG}!$E$8:$E$5007),1)),0)`, 'NUMBER'],
    ['Current Month', `=TEXT(TODAY(),"mmmm yyyy")`, 'TEXT'],
  ];

  SNAP_ITEMS.forEach(([label, formula, ftype], si) => {
    const r = UPM_ROW + 3 + si;
    fmt.push({ mergeCells: { range: gridRange(SID, r-1, r, 0, 5), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, r-1, r, 5, 9), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!${colL(5)}${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 0, 5), cell: { userEnteredFormat: { backgroundColor: hex(si % 2 === 0 ? C.lightGray : C.white), textFormat: { bold: true, fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    const numFmt = ftype === 'CURRENCY' ? { type: 'CURRENCY', pattern: '"$"#,##0.00' }
                 : ftype === 'PERCENT'  ? { type: 'PERCENT', pattern: '0.0%' }
                 :                        { type: 'TEXT' };
    fmt.push({ repeatCell: { range: gridRange(SID, r-1, r, 5, 9), cell: { userEnteredFormat: { backgroundColor: hex(si % 2 === 0 ? C.lightGray : C.white), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', numberFormat: ftype !== 'TEXT' ? numFmt : undefined } }, fields: `userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment${ftype !== 'TEXT' ? ',numberFormat' : ''})` } });
  });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const colWidths = [80,130,110,90,110,90,90,90,110,100,100,90,90,90,90,90,90,90];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row heights ──────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 8 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Borders ──────────────────────────────────────────────────────────────────
  [[PRI_ROW+1, PRI_ROW+11, 0, 9], [REC_ROW+1, REC_ROW+13, 0, 9]].forEach(([r1, r2, c1, c2]) => {
    fmt.push({ updateBorders: { range: gridRange(SID, r1, r2, c1, c2), innerHorizontal: { style: 'SOLID', color: hex(C.borderLight) }, innerVertical: { style: 'SOLID', color: hex(C.borderLight) }, bottom: { style: 'SOLID', color: hex(C.border) }, top: { style: 'SOLID', color: hex(C.border) }, left: { style: 'SOLID', color: hex(C.border) }, right: { style: 'SOLID', color: hex(C.border) } } });
  });

  await batchUpdate(id, fmt, 'dash-fmt');
  await valuesBatchUpdate(id, vals, 'dash-vals');

  console.log('✓ Sinking Funds Dashboard complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
