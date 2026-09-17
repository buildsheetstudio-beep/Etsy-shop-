'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Seasonal Planner'];
const S = "'Seasonal Planner'";
const TASK = "'Maintenance Task Log'";
const SETUP = "'Home & Property Setup'";
const NC = 14; // A-N

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title row ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 13, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['Seasonal Planner']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });

  // ── Subtitle row ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Your seasonal task checklist — live view filtered from the Task Log by season and status']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // ── Season filter row (row 2, 0-indexed) ─────────────────────────────────────
  // 4 season tabs as clickable visual — user selects via Dashboard Controls (C45 in Home Setup)
  const SEASONS = ['Winter','Spring','Summer','Fall'];
  const seasonColors = {
    'Winter': C.info,
    'Spring': C.success,
    'Summer': C.warning,
    'Fall':   C.secondary,
  };

  // Season summary cards (row 2)
  const sColStart = [0, 3, 7, 10];
  const sColEnd   = [3, 7, 10, 14];
  SEASONS.forEach((season, si) => {
    const c1 = sColStart[si];
    const c2 = sColEnd[si];
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, c1, c2), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, c1, c2), cell: { userEnteredFormat: { backgroundColor: hex(seasonColors[season]), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    // Show count of tasks for this season from Task Log
    const seasonFormula = `="${season} (" & SUMPRODUCT((${TASK}!I7:I1000="${season}")*1) & " tasks)"`;
    vals.push({ range: `${S}!${String.fromCharCode(65+c1)}3`, values: [[seasonFormula]] });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Divider (row 3) ──────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.wheat) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ── Active Season filter display (row 4) ─────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, 0, 4), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, 4), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A5`, values: [[`="Showing: " & IF(${SETUP}!C45="","All Seasons",${SETUP}!C45) & " — " & IF(${SETUP}!C43="","All Years",TEXT(${SETUP}!C43,"0"))`]] });

  fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, 4, 9), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 4, 9), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!E5`, values: [['← Tip: Change season filter in Home & Property Setup (C45)']] });

  // Status legend
  fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, 9, 14), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 9, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!J5`, values: [['Status: Not Started | Scheduled | In Progress | Complete']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // ── Header row (row 5, 0-indexed) ────────────────────────────────────────────
  const HEADERS = ['#','Task Name','Category','Season','Priority','Status','Due Date','Completed','DIY/Pro','Est. Cost ($)','Actual Cost ($)','Provider','Asset ID','Notes'];
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A6:N6`, values: [HEADERS] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Live data rows: FILTER from Task Log by selected season ──────────────────
  // Rows 6-155 (0-indexed), up to 150 tasks shown
  // Use IFERROR(INDEX(FILTER(...),n),"") pattern for each row
  // Filter condition: season matches C45 (or all if C45=""), AND property matches C42
  const NROWS = 150;
  const DATA_START = 6; // 0-indexed

  // Build the FILTER base — filter Task Log rows by season (C45) or all seasons
  // Columns mapped from Task Log: A=TaskLog!C(name), B=TaskLog!D(cat), C=TaskLog!I(season),
  //   D=TaskLog!G(priority), E=TaskLog!H(status), F=TaskLog!K(due), G=TaskLog!L(completed),
  //   H=TaskLog!M(DIY/Pro), I=TaskLog!Q(estcost), J=TaskLog!R(actualcost), K=TaskLog!N(provider),
  //   L=TaskLog!F(assetId), M=TaskLog!S(notes)

  const filterRange = `${TASK}!A7:X1000`;
  // Season filter: if C45="" then all rows, else filter by season
  const seasonFilter = `IF(OR(${SETUP}!C45="",${SETUP}!C45="All"),${TASK}!I7:I1000<>"",${TASK}!I7:I1000=${SETUP}!C45)`;
  const propFilter   = `${TASK}!B7:B1000=${SETUP}!C42`;

  const colMap = [
    { col: 'A', src: `${TASK}!C7:C1000` }, // task name
    { col: 'B', src: `${TASK}!D7:D1000` }, // category
    { col: 'C', src: `${TASK}!I7:I1000` }, // season
    { col: 'D', src: `${TASK}!G7:G1000` }, // priority
    { col: 'E', src: `${TASK}!H7:H1000` }, // status
    { col: 'F', src: `${TASK}!K7:K1000` }, // due date
    { col: 'G', src: `${TASK}!L7:L1000` }, // completed date
    { col: 'H', src: `${TASK}!M7:M1000` }, // diy/pro
    { col: 'I', src: `${TASK}!Q7:Q1000` }, // est cost
    { col: 'J', src: `${TASK}!R7:R1000` }, // actual cost
    { col: 'K', src: `${TASK}!N7:N1000` }, // provider
    { col: 'L', src: `${TASK}!F7:F1000` }, // asset id
    { col: 'M', src: `${TASK}!S7:S1000` }, // notes
  ];

  const filterCond = `(${propFilter})*(${seasonFilter})`;

  const dataVals = [];
  for (let ri = 0; ri < NROWS; ri++) {
    const row = [];
    const n = ri + 1;
    const rowArr = Array(NC).fill('');

    // col A(0) = row number
    rowArr[0] = `=IFERROR(IF(IFERROR(INDEX(FILTER(${TASK}!C7:C1000,${filterCond}),${n}),"")="","",${n}),"")`;

    colMap.forEach(({ col, src }, ci) => {
      rowArr[ci + 1] = `=IFERROR(INDEX(FILTER(${src},${filterCond}),${n}),"")`;
    });

    dataVals.push(rowArr);
  }

  vals.push({ range: `${S}!A7:N${6+NROWS}`, values: dataVals });

  // ── Data formatting ────────────────────────────────────────────────────────────
  for (let ri = 0; ri < NROWS; ri++) {
    const rowIdx = DATA_START + ri;
    const bg = ri % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Formula background for all data cells
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+NROWS, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Date formats: F (col 5) = Due Date, G (col 6) = Completed Date
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+NROWS, 5, 7), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Currency: I (col 8) = Est Cost, J (col 9) = Actual Cost
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+NROWS, 8, 10), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Wrap notes column (M = col 13)
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+NROWS, 13, 14), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat.wrapStrategy' } });

  // ── Summary row at bottom ─────────────────────────────────────────────────────
  const sumRow = DATA_START + NROWS; // 0-indexed
  fmt.push({ mergeCells: { range: gridRange(SID, sumRow, sumRow+1, 0, 4), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, sumRow, sumRow+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  vals.push({ range: `${S}!A${sumRow+1}`, values: [[`=SUMPRODUCT(${filterCond}*1) & " tasks matched"`]] });
  // Total actual cost of filtered tasks
  vals.push({ range: `${S}!J${sumRow+1}`, values: [[`=IFERROR(SUMPRODUCT((${filterCond})*ISNUMBER(${TASK}!R7:R1000)*${TASK}!R7:R1000),0)`]] });
  fmt.push({ repeatCell: { range: gridRange(SID, sumRow, sumRow+1, 9, 10), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: sumRow, endIndex: sumRow+1 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // ── Row heights (data) ────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DATA_START, endIndex: DATA_START+NROWS }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const colWidths = [35,240,120,80,70,90,90,90,80,85,90,140,80,180];
  colWidths.forEach((px, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'seasonal-fmt');
  await valuesBatchUpdate(id, vals, 'seasonal-vals');

  console.log(`✓ Seasonal Planner complete — ${NROWS} live-filter rows written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
