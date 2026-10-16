'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reading Dashboard'];
const S = "'Reading Dashboard'";
const LIB = "'Master Book Library'";
const INS = "'Reading Insights'";
const WL  = "'Wishlist'";
const REV = "'Book Review & Notes'";
const GOA = "'Goals & Challenges'";

(async () => {
  const fmt  = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,150,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Title ──────────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['📚 Reading Dashboard']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 18, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' }});

  // Subtitle / Last updated
  vals.push({ range: `${S}!A2`, values: [[`=IFERROR("Ultimate Book Tracker & Digital Reading Journal  |  Last Updated: "&TEXT(TODAY(),"mmmm d, yyyy"),"")`]] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── Row 3: 6 Key KPI Cards ────────────────────────────────────────────────
  const KPI_LABELS = ['📖 Total Books','✅ Finished','📄 Pages Read','⭐ Avg Rating','💝 Favorites','🔖 Reading Now'];
  const KPI_VALS = [
    `=COUNTA(${LIB}!$B$8:$B$1008)`,
    `=SUMPRODUCT((${LIB}!$M$8:$M$1008="Finished")*1)`,
    `=SUMPRODUCT((${LIB}!$M$8:$M$1008="Finished")*${LIB}!$Q$8:$Q$1008)`,
    `=IFERROR(ROUND(AVERAGEIF(${LIB}!$M$8:$M$1008,"Finished",${LIB}!$P$8:$P$1008),2),"—")`,
    `=SUMPRODUCT((${LIB}!$U$8:$U$1008=TRUE)*1)`,
    `=SUMPRODUCT((${LIB}!$M$8:$M$1008="Reading")*1)`,
  ];
  const KPI_COLORS = [C.primary, C.secondary, C.accent, C.secondary, C.primary, C.accent];
  const KPI_TINTS  = [C.wineTint, C.goldTint, C.greenTint, C.goldTint, C.wineTint, C.greenTint];

  KPI_LABELS.forEach((lbl, ki) => {
    const c1 = ki * 3;
    vals.push({ range: `${S}!${String.fromCharCode(65+c1)}3`, values: [[lbl]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c1)}4`, values: [[KPI_VALS[ki]]] });
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c1+3), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c1+3), cell: { userEnteredFormat: {
      backgroundColor: hex(KPI_COLORS[ki]),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c1,c1+3), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c1,c1+3), cell: { userEnteredFormat: {
      backgroundColor: hex(KPI_TINTS[ki]),
      textFormat: { bold: true, fontSize: 20, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'NUMBER', pattern: '#,##0' },
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ── Row 5: This Year's KPIs ───────────────────────────────────────────────
  const spYr = (col, extra = '') =>
    `SUMPRODUCT((YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=YEAR(TODAY()))*(${LIB}!$M$8:$M$1008="Finished")${extra})`;

  const YR_LABELS = ['📅 Read This Year','🎯 Annual Goal','📈 On Track?','📖 Pages This Year','⭐ Best Rating','📊 Goal Progress'];
  const YR_VALS = [
    `=${spYr('')}`,
    `=IFERROR(INDEX(FILTER('Goals & Challenges'!$C$6:$C$13,('Goals & Challenges'!$A$6:$A$13)=YEAR(TODAY())),1),"Set goal in Goals tab")`,
    `=IFERROR(IF(${spYr('')}>=INDEX(FILTER('Goals & Challenges'!$C$6:$C$13,('Goals & Challenges'!$A$6:$A$13)=YEAR(TODAY())),1)/12*(MONTH(TODAY())),"✅ On Track","⚠️ Behind"),"—")`,
    `=${spYr('','*\'Master Book Library\'!$Q$8:$Q$1008')}`,
    `=IFERROR(MAX(IF((YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=YEAR(TODAY()))*(${LIB}!$M$8:$M$1008="Finished"),${LIB}!$P$8:$P$1008)),"—")`,
    `=IFERROR(${spYr('')}&"/"&INDEX(FILTER('Goals & Challenges'!$C$6:$C$13,('Goals & Challenges'!$A$6:$A$13)=YEAR(TODAY())),1)&" books","—")`,
  ];
  const YR_COLORS = [C.info, C.warning, C.success, C.info, C.warning, C.success];
  const YR_TINTS  = ['#EAF3F9','#FBF0E0','#E3EDE7','#EAF3F9','#FBF0E0','#E3EDE7'];

  YR_LABELS.forEach((lbl, ki) => {
    const c1 = ki * 3;
    vals.push({ range: `${S}!${String.fromCharCode(65+c1)}6`, values: [[lbl]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c1)}7`, values: [[YR_VALS[ki]]] });
    fmt.push({ mergeCells: { range: gridRange(SID,5,6,c1,c1+3), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,5,6,c1,c1+3), cell: { userEnteredFormat: {
      backgroundColor: hex(YR_COLORS[ki]),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,6,7,c1,c1+3), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,6,7,c1,c1+3), cell: { userEnteredFormat: {
      backgroundColor: hex(YR_TINTS[ki]),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 38 }, fields: 'pixelSize' }});

  // ── Section: Top Reads This Year ──────────────────────────────────────────
  vals.push({ range: `${S}!A9`, values: [['⭐ TOP READS THIS YEAR (5-Star Books)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,8,9,0,9), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,8,9,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Top reads header
  vals.push({ range: `${S}!A10`, values: [['Title','Author','Genre','Date Finished','Rating','Pages']] });
  fmt.push({ repeatCell: { range: gridRange(SID,9,10,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.wineTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Top reads data — FILTER for 5-star books finished this year, constrained to 8 rows
  const topCond = `(${LIB}!$P$8:$P$1008=5)*(${LIB}!$M$8:$M$1008="Finished")*(YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=YEAR(TODAY()))`;
  const topFilter = (col) => `=IFERROR(ARRAY_CONSTRAIN(FILTER(${LIB}!${col}$8:${col}$1008,${topCond}),8,1),"")`;

  vals.push({ range: `${S}!A11`, values: [[topFilter('$B')]] });
  vals.push({ range: `${S}!B11`, values: [[topFilter('$C')]] });
  vals.push({ range: `${S}!C11`, values: [[topFilter('$D')]] });
  vals.push({ range: `${S}!D11`, values: [[topFilter('$O')]] });
  vals.push({ range: `${S}!E11`, values: [[topFilter('$P')]] });
  vals.push({ range: `${S}!F11`, values: [[topFilter('$Q')]] });

  for (let i = 0; i < 8; i++) {
    fmt.push({ repeatCell: { range: gridRange(SID,10+i,11+i,0,6), cell: { userEnteredFormat: {
      backgroundColor: hex(i % 2 === 0 ? C.white : C.altRow),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10+i, endIndex: 11+i }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  }

  // ── Section: Currently Reading ────────────────────────────────────────────
  vals.push({ range: `${S}!K9`, values: [['📖 CURRENTLY READING']] });
  fmt.push({ mergeCells: { range: gridRange(SID,8,9,10,18), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,8,9,10,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  vals.push({ range: `${S}!K10`, values: [['Title','Author','Genre','Pages Read','Total Pages','Progress %']] });
  fmt.push({ repeatCell: { range: gridRange(SID,9,10,10,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.greenTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  const rdCond = `${LIB}!$M$8:$M$1008="Reading"`;
  const rdFilter = (col) => `=IFERROR(ARRAY_CONSTRAIN(FILTER(${LIB}!${col}$8:${col}$1008,${rdCond}),8,1),"")`;
  vals.push({ range: `${S}!K11`, values: [[rdFilter('$B')]] });
  vals.push({ range: `${S}!L11`, values: [[rdFilter('$C')]] });
  vals.push({ range: `${S}!M11`, values: [[rdFilter('$D')]] });
  vals.push({ range: `${S}!N11`, values: [[rdFilter('$R')]] });
  vals.push({ range: `${S}!O11`, values: [[rdFilter('$Q')]] });
  vals.push({ range: `${S}!P11`, values: [[rdFilter('$S')]] });

  for (let i = 0; i < 8; i++) {
    fmt.push({ repeatCell: { range: gridRange(SID,10+i,11+i,10,16), cell: { userEnteredFormat: {
      backgroundColor: hex(i % 2 === 0 ? C.white : C.altRow),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  // Progress % format
  fmt.push({ repeatCell: { range: gridRange(SID,10,19,15,16), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // ── Section: Recent Reads (last 10 finished) ──────────────────────────────
  vals.push({ range: `${S}!A21`, values: [['📋 RECENT READS (Last 10 Finished)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,20,21,0,9), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,20,21,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 20, endIndex: 21 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A22`, values: [['Book ID','Title','Author','Genre','Date Finished','Rating','Pages','Format','Favorite?']] });
  fmt.push({ repeatCell: { range: gridRange(SID,21,22,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.wineTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 21, endIndex: 22 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Last 10 finished: sort by Date Finished descending, take top 10
  // Using SORT + FILTER then ARRAY_CONSTRAIN
  const sortedFinCond = `${LIB}!$M$8:$M$1008="Finished"`;
  const recFilter = (col) => `=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER(${LIB}!${col}$8:${col}$1008,${sortedFinCond}),SORT(FILTER(${LIB}!$O$8:$O$1008,${sortedFinCond}),1,-1),1,-1),10,1),"")`;
  // Simpler: just use FILTER + SORT
  // Actually the simpler version: filter and sort by col O (Date Finished) descending
  const recentCols = ['$A','$B','$C','$D','$O','$P','$Q','$F','$U'];
  const colLetters = ['A','B','C','D','E','F','G','H','I'];
  recentCols.forEach((libCol, ci) => {
    const cellLetter = colLetters[ci];
    vals.push({ range: `${S}!${cellLetter}23`, values: [[`=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER(${LIB}!${libCol}$8:${libCol}$1008,${sortedFinCond}),SORT(FILTER(${LIB}!$O$8:$O$1008,${sortedFinCond}),1,FALSE),1,FALSE),10,1),"")`]] });
  });

  for (let i = 0; i < 10; i++) {
    fmt.push({ repeatCell: { range: gridRange(SID,22+i,23+i,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(i % 2 === 0 ? C.white : C.altRow),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 22+i, endIndex: 23+i }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  }

  // ── Section: Top Wishlist Priorities ─────────────────────────────────────
  vals.push({ range: `${S}!K21`, values: [['🌟 TOP WISHLIST (High Priority)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,20,21,10,18), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,20,21,10,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  vals.push({ range: `${S}!K22`, values: [['Title','Author','Genre','Priority','Source','Est. Cost']] });
  fmt.push({ repeatCell: { range: gridRange(SID,21,22,10,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  const wlCond = `${WL}!$K$6:$K$506="High"`;
  const wlFilter = (col) => `=IFERROR(ARRAY_CONSTRAIN(FILTER(${WL}!${col}$6:${col}$506,${wlCond}),10,1),"")`;
  const wlCols = ['$B','$C','$D','$K','$L','$R'];
  const wlColLetters = ['K','L','M','N','O','P'];
  wlCols.forEach((wlCol, ci) => {
    vals.push({ range: `${S}!${wlColLetters[ci]}23`, values: [[wlFilter(wlCol)]] });
  });

  for (let i = 0; i < 10; i++) {
    fmt.push({ repeatCell: { range: gridRange(SID,22+i,23+i,10,16), cell: { userEnteredFormat: {
      backgroundColor: hex(i % 2 === 0 ? C.white : C.altRow),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  // Currency format for Est. Cost col
  fmt.push({ repeatCell: { range: gridRange(SID,22,32,15,16), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // ── Column widths ──────────────────────────────────────────────────────────
  const widths = [90,160,130,100,80,60, 80,80,60, 60,90,160,130,80,80,70, 80,80,60,60];
  widths.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze top rows
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 8 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '10-dashboard format');
  await valuesBatchUpdate(id, vals, '10-dashboard values');
  console.log('✅  Reading Dashboard done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
