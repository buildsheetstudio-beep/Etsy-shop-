'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Progress & Trends'];
const S = "'Progress & Trends'";
const WP = "'Weight Progress'";
const NL = "'Nutrition Log'";
const WT = "'Workout Tracker'";
const HT = "'Habit Tracker'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['PROGRESS & TRENDS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle + Year input
  vals.push({ range: `${S}!A2`, values: [['Change the year in cell B2 to view that year\'s data. All summary formulas update automatically.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Row 3: Year filter input
  vals.push({ range: `${S}!A3`, values: [['Year:', 2026]] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Spacer row 4
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ── SECTION 1: Monthly Weight Summary (rows 5-6 header, 7-18 = Jan-Dec) ─────
  // Section header row 5 (ri=4)
  vals.push({ range: `${S}!A5`, values: [['MONTHLY WEIGHT SUMMARY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,4,5,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Row 6 (ri=5): column headers for monthly summary
  vals.push({ range: `${S}!A6`, values: [['Month','Weight Entries','Avg Weight (lbs)','Min Weight (lbs)','Max Weight (lbs)','Total Cal Logged','Workout Sessions','Check-Ins']] });
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Rows 7-18 (ri=6-17): Jan-Dec formulas
  // Month = ROW()-6, so row 7 = month 1 (January)
  // Use AVERAGEIFS/COUNTIFS/SUMIFS with date range bounds — no YEAR(range) in criteria
  const YEAR_CELL = '$B$3';
  for (let m = 1; m <= 12; m++) {
    const ri = 5 + m;  // ri=6 for Jan, ri=17 for Dec
    const dateStart = `DATE(${YEAR_CELL},${m},1)`;
    const dateEnd = `DATE(${YEAR_CELL},${m}+1,0)`;
    vals.push({ range: `${S}!A${ri+1}`, values: [[`=TEXT(DATE(${YEAR_CELL},${m},1),"MMMM YYYY")`]] });
    vals.push({ range: `${S}!B${ri+1}`, values: [[`=IFERROR(COUNTIFS(${WP}!$B$8:$B$5000,">="&${dateStart},${WP}!$B$8:$B$5000,"<="&${dateEnd}),"")` ]] });
    vals.push({ range: `${S}!C${ri+1}`, values: [[`=IFERROR(ROUND(AVERAGEIFS(${WP}!$C$8:$C$5000,${WP}!$B$8:$B$5000,">="&${dateStart},${WP}!$B$8:$B$5000,"<="&${dateEnd}),1),"")` ]] });
    vals.push({ range: `${S}!D${ri+1}`, values: [[`=IFERROR(MINIFS(${WP}!$C$8:$C$5000,${WP}!$B$8:$B$5000,">="&${dateStart},${WP}!$B$8:$B$5000,"<="&${dateEnd},${WP}!$C$8:$C$5000,"<>"),"")` ]] });
    vals.push({ range: `${S}!E${ri+1}`, values: [[`=IFERROR(MAXIFS(${WP}!$C$8:$C$5000,${WP}!$B$8:$B$5000,">="&${dateStart},${WP}!$B$8:$B$5000,"<="&${dateEnd},${WP}!$C$8:$C$5000,"<>"),"")` ]] });
    vals.push({ range: `${S}!F${ri+1}`, values: [[`=IFERROR(SUMIFS(${NL}!$G$8:$G$10000,${NL}!$B$8:$B$10000,">="&${dateStart},${NL}!$B$8:$B$10000,"<="&${dateEnd}),"")` ]] });
    vals.push({ range: `${S}!G${ri+1}`, values: [[`=IFERROR(COUNTIFS(${WT}!$B$8:$B$5000,">="&${dateStart},${WT}!$B$8:$B$5000,"<="&${dateEnd}),"")` ]] });
    vals.push({ range: `${S}!H${ri+1}`, values: [[`=IFERROR(COUNTIFS('Weekly Check-In'!$B$8:$B$500,">="&${dateStart},'Weekly Check-In'!$B$8:$B$500,"<="&${dateEnd}),"")` ]] });
  }

  // Format monthly rows
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,6,18,0,8),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Spacer after section 1
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 18, endIndex: 19 }, properties: { pixelSize: 12 }, fields: 'pixelSize' }});

  // ── SECTION 2: Workout Type Breakdown (rows 20-21 header, 22-29 data) ─────────
  vals.push({ range: `${S}!A20`, values: [['WORKOUT TYPE BREAKDOWN (SELECTED YEAR)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,19,20,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,19,20,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 19, endIndex: 20 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Row 21 (ri=20): headers
  vals.push({ range: `${S}!A21`, values: [['Workout Type','Sessions This Year','Avg Duration (min)','Total Duration (min)']] });
  fmt.push({ repeatCell: { range: gridRange(SID,20,21,0,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 20, endIndex: 21 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Workout type rows 22-29 (ri=21-28)
  const WTYPES = ['Strength','Cardio','Yoga/Flexibility','Walking','Cycling','Swimming','Sports','Other'];
  const YEAR_D_START = `DATE(${YEAR_CELL},1,1)`;
  const YEAR_D_END = `DATE(${YEAR_CELL},12,31)`;
  WTYPES.forEach((wtype, i) => {
    const ri = 21 + i;
    vals.push({ range: `${S}!A${ri+1}`, values: [[wtype]] });
    vals.push({ range: `${S}!B${ri+1}`, values: [[`=IFERROR(COUNTIFS(${WT}!$C$8:$C$5000,"${wtype}",${WT}!$B$8:$B$5000,">="&${YEAR_D_START},${WT}!$B$8:$B$5000,"<="&${YEAR_D_END}),"")`]] });
    vals.push({ range: `${S}!C${ri+1}`, values: [[`=IFERROR(ROUND(AVERAGEIFS(${WT}!$E$8:$E$5000,${WT}!$C$8:$C$5000,"${wtype}",${WT}!$B$8:$B$5000,">="&${YEAR_D_START},${WT}!$B$8:$B$5000,"<="&${YEAR_D_END}),0),"")`]] });
    vals.push({ range: `${S}!D${ri+1}`, values: [[`=IFERROR(SUMIFS(${WT}!$E$8:$E$5000,${WT}!$C$8:$C$5000,"${wtype}",${WT}!$B$8:$B$5000,">="&${YEAR_D_START},${WT}!$B$8:$B$5000,"<="&${YEAR_D_END}),"")`]] });
  });

  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,21,29,0,4),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 29, endIndex: 30 }, properties: { pixelSize: 12 }, fields: 'pixelSize' }});

  // ── SECTION 3: Habit Compliance Summary (rows 31-32 header, 33-43 data) ──────
  vals.push({ range: `${S}!A31`, values: [['HABIT COMPLIANCE SUMMARY (ALL TIME)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,30,31,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,30,31,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 30, endIndex: 31 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Row 32 (ri=31): headers
  vals.push({ range: `${S}!A32`, values: [['Habit Name','Total Done','Total Logged','Completion %']] });
  fmt.push({ repeatCell: { range: gridRange(SID,31,32,0,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 31, endIndex: 32 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Habit rows 33-42 (ri=32-41): pull from habit definitions (B4:B13)
  for (let h = 0; h < 10; h++) {
    const ri = 32 + h;
    const hRow = 4 + h;  // habit definition rows 4-13
    vals.push({ range: `${S}!A${ri+1}`, values: [[`=IFERROR(${HT}!B${hRow},"")`]] });
    vals.push({ range: `${S}!B${ri+1}`, values: [[`=IFERROR(COUNTIFS(${HT}!$B$20:$B$10000,${HT}!$A${hRow},${HT}!$D$20:$D$10000,TRUE),"")`]] });
    vals.push({ range: `${S}!C${ri+1}`, values: [[`=IFERROR(COUNTIF(${HT}!$B$20:$B$10000,${HT}!$A${hRow}),"")`]] });
    vals.push({ range: `${S}!D${ri+1}`, values: [[`=IFERROR(IF(C${ri+1}=0,"",TEXT(B${ri+1}/C${ri+1},"0%")),"")`]] });
  }

  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,32,42,0,4),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Column widths
  const WIDTHS = [160, 130, 130, 130, 130, 120, 110, 110, 110, 110, 100, 100];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-3
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '14-trends values');
  await batchUpdate(id, fmt, '14-trends format');
  console.log('✅  Progress & Trends done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
