'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weight Progress'];
const S = "'Weight Progress'";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function fmtDate(d) { return d.toISOString().split('T')[0]; }

// 100 weekly entries: 2024-10-07 through 2026-09-21
const NOISE = [0.8,-0.5,1.2,-0.9,0.3,-1.1,0.7,-0.4,1.0,-0.8,0.5,-1.3,0.9,-0.6,1.4,-0.7,
               0.2,-1.0,0.6,-0.3,1.1,-0.9,0.4,-0.7,0.8,-1.2,0.5,-0.4,0.9,-0.6,1.0,-0.8,
               0.3,-1.1,0.7,-0.5,1.2,-0.9,0.4,-0.7,0.6,-1.0,0.8,-0.4,1.1,-0.8,0.5,-0.6,
               0.9,-1.2,0.4,-0.7,0.8,-0.5,1.0,-0.9,0.3,-1.1,0.6,-0.4,0.9,-0.7,1.2,-0.8,
               0.5,-1.0,0.7,-0.5,1.1,-0.9,0.4,-0.6,0.8,-1.2,0.5,-0.8,0.9,-0.5,1.0,-0.7,
               0.4,-1.1,0.6,-0.9,0.8,-0.4,1.2,-0.7,0.5,-1.0,0.7,-0.8,0.3,-0.6,1.1,-0.9,
               0.5,-0.5,0.8,-1.0];
const START = new Date('2024-10-07');
let w = 185.2;
const weightRows = [];
const NOTES = {
  0: 'Starting point',
  12: 'Feeling more energetic this month',
  26: 'Started new morning routine',
  52: 'Halfway through the year!',
  75: 'Completed 10K walk challenge',
  99: 'One year of consistent tracking',
};
for (let i = 0; i < 100; i++) {
  w = Math.round((w - 0.05 + NOISE[i]) * 10) / 10;
  weightRows.push([fmtDate(addDays(START, i * 7)), w, '', '', '', '', '', NOTES[i] || '']);
}

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,10100,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['WEIGHT PROGRESS TRACKER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,9), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Log your weight. Formula columns auto-fill. Track trends over time — not day-to-day fluctuations.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,9), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Quick stats
  const STATS = [
    ['Entries Logged',    `=COUNTA(B8:B10000)`],
    ['Days Tracked',      `=IFERROR(MAX(B8:B10000)-MIN(B8:B10000),"")`],
    ['Most Recent (lbs)', `=IFERROR(LOOKUP(2,1/(C8:C5000<>""),C8:C5000),"")`],
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
  vals.push({ range: `${S}!A7`, values: [['Entry ID','Date','Weight (lbs)','Weight (kg)','7-Day Avg (lbs)','30-Day Avg (lbs)','Change vs Prior','Change vs Start','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Formula columns (data rows start at ri=7 = spreadsheet row 8)
  // A: Entry ID
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","WT-"&TEXT(ROW()-7,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // D: Weight kg
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,3,4), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(ROUND(C8/2.20462,1),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // E: 7-day moving average
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,4,5), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(ROUND(AVERAGE(OFFSET(C8,-(MIN(7,ROW()-7)-1),0,MIN(7,ROW()-7),1)),1),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // F: 30-day moving average
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,5,6), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(ROUND(AVERAGEIFS(C$8:C8,B$8:B8,">="&(B8-29),B$8:B8,"<="&B8),1),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // G: Change vs prior entry
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,6,7), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(IF(ROW()=8,"",ROUND(C8-C7,1)),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '+0.0;-0.0;0.0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // H: Change vs start
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,7,8), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(ROUND(C8-$C$8,1),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '+0.0;-0.0;0.0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push sample data (B=date, C=weight, D-H=formulas, I=notes)
  for (let i = 0; i < weightRows.length; i += 200) {
    vals.push({ range: `${S}!B${8+i}`, values: weightRows.slice(i, i+200) });
  }

  // Alternating row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,10100,0,9),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col B
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths: A B C D E F G H I
  const WIDTHS = [80, 100, 90, 90, 110, 110, 110, 110, 220];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7 (no frozenColumnCount — title merged across all 9 cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '03-weight values');
  await batchUpdate(id, fmt, '03-weight format');
  console.log(`✅  Weight Progress done — ${weightRows.length} entries.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
