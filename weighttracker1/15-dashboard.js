'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weight Dashboard'];
const S = "'Weight Dashboard'";
const WP = "'Weight Progress'";
const WT = "'Workout Tracker'";
const NL = "'Nutrition Log'";
const CI = "'Weekly Check-In'";
const GM = "'Goals & Milestones'";
const HT = "'Habit Tracker'";
const NSV = "'Non-Scale Victories'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title (merged)
  vals.push({ range: `${S}!A1`, values: [['WEIGHT MANAGEMENT & LIFESTYLE DASHBOARD']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 18, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Personal tracking summary — all numbers pull from your other tabs. Change year in B3 to view historical data.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Row 3: Year filter input
  vals.push({ range: `${S}!A3`, values: [['Year:', 2026]] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,1,3), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 13, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Spacer row 4
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ── KPI BOXES: Rows 5-7 (3 rows of KPI pairs) ─────────────────────────────
  // Each KPI box = 3 cols wide: label (merged 3 cols, ri=labelRi), value (merged 3 cols, ri=labelRi+1)
  // 4 boxes per row across 12 cols (cols 0-2, 3-5, 6-8, 9-11)
  // Row 1 of KPIs (ri=4,5)
  const KPI_ROW1 = [
    { label: 'Entries Logged (All Time)', formula: `=IFERROR(COUNTA(${WP}!B8:B5000),"")` },
    { label: 'Starting Weight (lbs)',     formula: `=IFERROR(INDEX(${WP}!C8:C5000,MATCH(TRUE,${WP}!C8:C5000<>"",0)),"")` },
    { label: 'Most Recent Weight (lbs)',  formula: `=IFERROR(LOOKUP(2,1/(${WP}!C8:C5000<>""),${WP}!C8:C5000),"")` },
    { label: '7-Day Moving Avg (lbs)',   formula: `=IFERROR(LOOKUP(2,1/(${WP}!E8:E5000<>""),${WP}!E8:E5000),"")` },
  ];
  const KPI_ROW2 = [
    { label: 'Workout Sessions This Year', formula: `=IFERROR(COUNTIFS(${WT}!$B$8:$B$5000,">="&DATE($B$3,1,1),${WT}!$B$8:$B$5000,"<="&DATE($B$3,12,31)),"")` },
    { label: 'Nutrition Days Logged',      formula: `=IFERROR(SUMPRODUCT((${NL}!B8:B5000<>"")*IFERROR(1/COUNTIF(${NL}!B8:B5000,${NL}!B8:B5000),0)),"")` },
    { label: 'Active Goals',               formula: `=IFERROR(COUNTIF(${GM}!D8:D500,"Active"),"")` },
    { label: 'Active Habits',              formula: `=IFERROR(COUNTIF(${HT}!F4:F13,TRUE),"")` },
  ];
  const KPI_ROW3 = [
    { label: 'Weekly Check-Ins Logged', formula: `=IFERROR(COUNTA(${CI}!B8:B500),"")` },
    { label: 'Non-Scale Victories',     formula: `=IFERROR(COUNTA(${NSV}!B8:B500),"")` },
    { label: 'Habit Log Entries',       formula: `=IFERROR(COUNTA(${HT}!B20:B10000),"")` },
    { label: 'Meals Planned',           formula: `=IFERROR(COUNTA('Weekly Meal Planner'!D8:D5000),"")` },
  ];

  [KPI_ROW1, KPI_ROW2, KPI_ROW3].forEach((kpiRow, rowIdx) => {
    const labelRi = 4 + rowIdx * 3;
    const valueRi = labelRi + 1;
    kpiRow.forEach(({ label, formula }, boxIdx) => {
      const c1 = boxIdx * 3;
      const c2 = c1 + 3;
      fmt.push({ mergeCells: { range: gridRange(SID,labelRi,labelRi+1,c1,c2), mergeType: 'MERGE_ALL' }});
      fmt.push({ mergeCells: { range: gridRange(SID,valueRi,valueRi+1,c1,c2), mergeType: 'MERGE_ALL' }});
      vals.push({ range: `${S}!${String.fromCharCode(65+c1)}${labelRi+1}`, values: [[label]] });
      vals.push({ range: `${S}!${String.fromCharCode(65+c1)}${valueRi+1}`, values: [[formula]] });
      fmt.push({ repeatCell: { range: gridRange(SID,labelRi,labelRi+1,c1,c2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      }}, fields: 'userEnteredFormat' }});
      fmt.push({ repeatCell: { range: gridRange(SID,valueRi,valueRi+1,c1,c2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.white), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat' }});
    });
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: labelRi, endIndex: labelRi+1 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: valueRi, endIndex: valueRi+1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});
    // Spacer row between KPI groups
    if (rowIdx < 2) {
      fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: valueRi+1, endIndex: valueRi+2 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});
    }
  });

  // Spacer before recent entries
  const RECENT_SPACER_RI = 13;
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: RECENT_SPACER_RI, endIndex: RECENT_SPACER_RI+1 }, properties: { pixelSize: 10 }, fields: 'pixelSize' }});

  // ── RECENT WEIGHT ENTRIES (rows 15-16 header, 17-26 = last 10 entries) ─────
  const RECENT_HEADER_RI = 14;
  vals.push({ range: `${S}!A${RECENT_HEADER_RI+1}`, values: [['RECENT WEIGHT ENTRIES (LAST 10)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,RECENT_HEADER_RI,RECENT_HEADER_RI+1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,RECENT_HEADER_RI,RECENT_HEADER_RI+1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: RECENT_HEADER_RI, endIndex: RECENT_HEADER_RI+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Column headers for recent entries
  const REC_COL_RI = RECENT_HEADER_RI + 1;
  vals.push({ range: `${S}!A${REC_COL_RI+1}`, values: [['Date','Weight (lbs)','Weight (kg)','7-Day Avg','30-Day Avg','Change vs Prior','Change vs Start','Notes','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,REC_COL_RI,REC_COL_RI+1,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: REC_COL_RI, endIndex: REC_COL_RI+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // 10 recent entry rows (ri = REC_COL_RI+1 through REC_COL_RI+10)
  const totalCol = `COUNTA(${WP}!$B$8:$B$5000)`;
  for (let i = 0; i < 10; i++) {
    const ri = REC_COL_RI + 1 + i;
    const idx = `MAX(${totalCol}-${i},1)`;
    vals.push({ range: `${S}!A${ri+1}`, values: [[
      `=IFERROR(INDEX(${WP}!$B$8:$B$5000,${idx}),"")`,
      `=IFERROR(INDEX(${WP}!$C$8:$C$5000,${idx}),"")`,
      `=IFERROR(INDEX(${WP}!$D$8:$D$5000,${idx}),"")`,
      `=IFERROR(INDEX(${WP}!$E$8:$E$5000,${idx}),"")`,
      `=IFERROR(INDEX(${WP}!$F$8:$F$5000,${idx}),"")`,
      `=IFERROR(INDEX(${WP}!$G$8:$G$5000,${idx}),"")`,
      `=IFERROR(INDEX(${WP}!$H$8:$H$5000,${idx}),"")`,
      `=IFERROR(INDEX(${WP}!$I$8:$I$5000,${idx}),"")`,
    ]] });
  }

  // Format recent rows
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,REC_COL_RI+1,REC_COL_RI+11,0,8),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});
  fmt.push({ repeatCell: { range: gridRange(SID,REC_COL_RI+1,REC_COL_RI+11,0,8), cell: { userEnteredFormat: {
    textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(textFormat,verticalAlignment,horizontalAlignment)' }});
  fmt.push({ repeatCell: { range: gridRange(SID,REC_COL_RI+1,REC_COL_RI+11,0,1), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: REC_COL_RI+1, endIndex: REC_COL_RI+11 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Spacer before disclaimer
  const DISC_RI = REC_COL_RI + 12;
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DISC_RI - 1, endIndex: DISC_RI }, properties: { pixelSize: 12 }, fields: 'pixelSize' }});

  // ── DISCLAIMER BOX ────────────────────────────────────────────────────────
  const DISCLAIMER = 'This workbook is for personal organization and self-tracking only. It is not medical advice, nutritional guidance, or a substitute for professional support. Consult a qualified healthcare provider before making changes to your diet, exercise routine, or health goals.';
  vals.push({ range: `${S}!A${DISC_RI+1}`, values: [[DISCLAIMER]] });
  fmt.push({ mergeCells: { range: gridRange(SID,DISC_RI,DISC_RI+3,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,DISC_RI,DISC_RI+3,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.neutral),
    textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DISC_RI, endIndex: DISC_RI+3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Column widths
  const WIDTHS = [120, 110, 110, 110, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-3 only (NO frozenColumnCount — title merged across all 14 cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '15-dashboard values');
  await batchUpdate(id, fmt, '15-dashboard format');
  console.log('✅  Weight Dashboard done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
