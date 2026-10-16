'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Macro Tracker'];
const S = "'Macro Tracker'";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function fmtDate(d) { return d.toISOString().split('T')[0]; }

// Same 103-day range as nutrition log: Jun 1 - Sep 11, 2026
// One row per day; user enters date, formulas aggregate from Nutrition Log
const START = new Date('2026-06-01');
const dateRows = [];
for (let day = 0; day < 103; day++) {
  dateRows.push([fmtDate(addDays(START, day)), '', '', '', '', '', '', '', '', 2000, 130, 220, 65, '']);
  // B=date, C-I=aggregation formulas (empty here, set via repeatCell), J=cal target, K=protein target, L=carb target, M=fat target, N=notes
}

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,10100,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['DAILY MACRO TRACKER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,15), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Daily totals auto-sum from Nutrition Log. Set your personal targets in columns J–M. Targets are for reference — adjust them to match what works for you.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,15), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Days Logged',         `=COUNTA(A8:A5000)`],
    ['Avg Daily Calories',  `=IFERROR(ROUND(AVERAGEIF(B8:B5000,"<>",B8:B5000),0),"")`],
    ['Avg Daily Protein(g)',`=IFERROR(ROUND(AVERAGEIF(C8:C5000,"<>",C8:C5000),0),"")`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,4), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,4,9), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!E${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Date','Calories','Protein (g)','Carbs (g)','Fat (g)','Fiber (g)','Sugar (g)','Sodium (mg)','Water (ml)','Cal Target','Protein Target','Carb Target','Fat Target','Cal vs Target','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Formula columns B-I (aggregating from Nutrition Log)
  const NL = "'Nutrition Log'";
  // B: Total Calories
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,1,2), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$G$8:$G$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // C: Protein
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,2,3), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$H$8:$H$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // D: Carbs
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,3,4), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$I$8:$I$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // E: Fat
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,4,5), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$J$8:$J$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // F: Fiber
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,5,6), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$K$8:$K$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // G: Sugar
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,6,7), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$L$8:$L$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // H: Sodium
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,7,8), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$M$8:$M$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // I: Water
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,8,9), cell: {
    userEnteredValue: { formulaValue: `=IF(A8="","",IFERROR(SUMIFS(${NL}!$N$8:$N$10000,${NL}!$B$8:$B$10000,A8),0))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});
  // N: Cal vs Target
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,13,14), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(IF(OR(B8="",J8=""),"",B8-J8),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '+#,##0;-#,##0;0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push date + target rows (col A = date, J-M = targets, O = notes)
  // dateRows: [date, '', '', '', '', '', '', '', '', cal_target, protein_target, carb_target, fat_target, notes]
  vals.push({ range: `${S}!A8`, values: dateRows.map(r => [r[0]]) });                       // col A: dates
  vals.push({ range: `${S}!J8`, values: dateRows.map(r => [r[9], r[10], r[11], r[12]]) }); // cols J-M: targets

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,5100,0,15),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col A
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,0,1), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Highlight target columns J-M
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,9,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.warmTint), textFormat: { fontFamily: 'Arial', italic: true },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  // Column widths
  const WIDTHS = [100, 80, 80, 80, 60, 60, 70, 80, 75, 80, 90, 80, 80, 90, 180];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '06-macros values');
  await batchUpdate(id, fmt, '06-macros format');
  console.log(`✅  Macro Tracker done — ${dateRows.length} daily rows.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
