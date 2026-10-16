'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Body Measurements'];
const S = "'Body Measurements'";

// 24 monthly measurements: Oct 2024 - Sep 2026
function addMonths(d, n) {
  const r = new Date(d);
  r.setMonth(r.getMonth() + n);
  return r;
}
const START = new Date('2024-10-01');
// [neck, chest, waist, hips, l_arm, r_arm, l_thigh, r_thigh, l_calf, r_calf, body_fat]
// Gentle random-ish fluctuation with slight downward drift on waist/hips
const BASE = [14.5, 40.0, 34.5, 41.0, 13.5, 13.5, 23.0, 23.0, 14.5, 14.5, 22.5];
const DRIFT = [0, -0.05, -0.08, -0.06, 0, 0, -0.04, -0.04, 0, 0, -0.08];
const NOISE_M = [
  [0.2,-0.3,0.1,-0.2,0.1,0.1,-0.1,0.0,0.1,0.0,0.2],
  [-0.1,0.2,-0.3,0.1,-0.1,-0.1,0.2,-0.1,-0.1,0.1,-0.3],
  [0.1,-0.1,0.2,-0.3,0.0,0.0,-0.2,0.1,0.0,-0.1,0.1],
  [-0.2,0.3,-0.1,0.2,0.1,0.1,0.1,-0.2,0.1,0.0,-0.2],
  [0.1,-0.2,0.3,-0.1,-0.1,-0.1,-0.1,0.1,-0.1,0.1,0.3],
  [-0.1,0.1,-0.2,0.3,0.0,0.0,0.2,-0.1,0.0,-0.1,-0.1],
  [0.2,-0.3,0.1,-0.2,0.1,0.1,-0.1,0.0,0.1,0.1,0.2],
  [-0.1,0.2,-0.3,0.1,-0.1,-0.1,0.2,-0.1,-0.1,0.0,-0.3],
  [0.1,-0.1,0.2,-0.3,0.0,0.0,-0.2,0.1,0.0,0.1,0.1],
  [-0.2,0.3,-0.1,0.2,0.1,0.1,0.1,-0.2,0.1,0.0,-0.2],
  [0.1,-0.2,0.3,-0.1,-0.1,-0.1,-0.1,0.1,-0.1,0.1,0.3],
  [-0.1,0.1,-0.2,0.3,0.0,0.0,0.2,-0.1,0.0,-0.1,-0.1],
  [0.2,-0.3,0.1,-0.2,0.1,0.1,-0.1,0.0,0.1,0.1,0.2],
  [-0.1,0.2,-0.3,0.1,-0.1,-0.1,0.2,-0.1,-0.1,0.0,-0.3],
  [0.1,-0.1,0.2,-0.3,0.0,0.0,-0.2,0.1,0.0,0.1,0.1],
  [-0.2,0.3,-0.1,0.2,0.1,0.1,0.1,-0.2,0.1,0.0,-0.2],
  [0.1,-0.2,0.3,-0.1,-0.1,-0.1,-0.1,0.1,-0.1,0.1,0.3],
  [-0.1,0.1,-0.2,0.3,0.0,0.0,0.2,-0.1,0.0,-0.1,-0.1],
  [0.2,-0.3,0.1,-0.2,0.1,0.1,-0.1,0.0,0.1,0.1,0.2],
  [-0.1,0.2,-0.3,0.1,-0.1,-0.1,0.2,-0.1,-0.1,0.0,-0.3],
  [0.1,-0.1,0.2,-0.3,0.0,0.0,-0.2,0.1,0.0,0.1,0.1],
  [-0.2,0.3,-0.1,0.2,0.1,0.1,0.1,-0.2,0.1,0.0,-0.2],
  [0.1,-0.2,0.3,-0.1,-0.1,-0.1,-0.1,0.1,-0.1,0.1,0.3],
  [-0.1,0.1,-0.2,0.3,0.0,0.0,0.2,-0.1,0.0,-0.1,-0.1],
];
const current = [...BASE];
const measureRows = [];
for (let i = 0; i < 24; i++) {
  const date = addMonths(START, i).toISOString().split('T')[0];
  const row = current.map((v, j) => {
    current[j] = Math.round((current[j] + DRIFT[j] + NOISE_M[i][j]) * 10) / 10;
    return current[j];
  });
  measureRows.push([date, ...row, '']);  // date + 11 measurements + notes
}

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,10100,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['BODY MEASUREMENTS TRACKER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Enter measurements in inches. Monthly tracking recommended. All measurement columns are user-entered — there are no right or wrong numbers.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Entries Logged',   `=COUNTA(B8:B5000)`],
    ['Date Range',       `=IFERROR(TEXT(MIN(B8:B5000),"mmm yyyy")&" – "&TEXT(MAX(B8:B5000),"mmm yyyy"),"")`],
    ['Latest Waist (in)',`=IFERROR(LOOKUP(2,1/(E8:E5000<>""),E8:E5000),"")`],
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
  vals.push({ range: `${S}!A7`, values: [['Entry ID','Date','Neck (in)','Chest (in)','Waist (in)','Hips (in)','L Arm (in)','R Arm (in)','L Thigh (in)','R Thigh (in)','L Calf (in)','R Calf (in)','Body Fat %','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // A: Entry ID formula
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","MS-"&TEXT(ROW()-7,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Sample data (B=date, C-M=measurements, N=notes)
  vals.push({ range: `${S}!B8`, values: measureRows });

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,10100,0,14),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col B
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [75, 100, 75, 75, 75, 75, 70, 70, 80, 80, 70, 70, 75, 180];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '04-measurements values');
  await batchUpdate(id, fmt, '04-measurements format');
  console.log(`✅  Body Measurements done — ${measureRows.length} entries.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
