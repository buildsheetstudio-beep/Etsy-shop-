'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WKL = sheetMap['📅 Weekly Check-In & Reflection'];

// Layout (GSheets 1-indexed):
// Row 1: title banner (A standalone, B:K merged)
// Row 2: column headers
// Rows 3-54: week 1-52 data
// Cols M-N rows 2-4: sparkline/trend block

const NCOLS = 11; // A-K (indices 0-10)

// Weeks 1-6 fully filled, 7-26 scores only, 27-52 blank
const WEEKLY_FULL = [
  [1,'Goal 1, Goal 3','Started leadership course and savings transfer','Busy with work','Week 2 of course',4,5,4],
  [2,'Goal 1, Goal 2','Completed first 5K run attempt','Tired from new exercise routine','Keep running, module 2',3,4,3],
  [3,'Goal 2, Goal 4','Duolingo streak at 14 days','Struggled to find time for Spanish','Spanish 30 mins daily',3,3,3],
  [4,'Goal 1, Goal 2, Goal 3','Finished leadership module 3, cancelled subscriptions','Nothing major','Leadership module 4',4,4,4],
  [5,'Goal 4, Goal 6','Registered blog domain!','Hard to write consistently','Draft first blog post',3,4,3],
  [6,'Goal 1','Completed entire leadership course!','Feeling tired this week','Apply learnings at work',2,3,3],
];
// Weeks 7-26: just scores
const WEEKLY_SCORES = [
  [7,4,4,4],[8,3,4,3],[9,5,5,5],[10,4,4,4],[11,3,3,3],[12,4,4,4],
  [13,3,5,4],[14,4,4,4],[15,2,3,2],[16,4,4,4],[17,5,5,5],[18,4,4,4],
  [19,3,3,3],[20,4,4,4],[21,3,4,3],[22,5,5,5],[23,4,4,4],[24,3,3,3],
  [25,4,4,4],[26,4,5,4],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'📅 Weekly Check-In & Reflection'";

  // Column widths A-K
  [60,110,110,200,220,200,200,80,90,90,80].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: WKL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  // Cols M-N for trend block
  [130,260].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: WKL, dimension: 'COLUMNS', startIndex: 12+i, endIndex: 13+i },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner (A standalone, B:K merged)
  reqs.push({ repeatCell: {
    range: gridRange(WKL,0,1,0,1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.midnightBlue) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(WKL,0,1,1,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(WKL,0,1,1,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midnightBlue),
      textFormat: { foregroundColor: hex(C.electricGold), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WKL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['  📅  Weekly Check-In & Reflection — 2025']] });

  // Trend sparkline block: M-N cols, rows 1-3 (0-indexed)
  reqs.push({ repeatCell: {
    range: gridRange(WKL,0,4,12,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WKL, dimension: 'ROWS', startIndex: 1, endIndex: 4 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!M1`, values: [['ENERGY TREND', `=SPARKLINE(H3:H54,{"charttype","line";"color","#C9A800";"lineWidth",2})`]] });
  vals.push({ range: `${TAB}!M2`, values: [['MOTIVATION TREND', `=SPARKLINE(I3:I54,{"charttype","line";"color","#F4D03F";"lineWidth",2})`]] });
  vals.push({ range: `${TAB}!M3`, values: [['AVG WEEK RATING', `=IFERROR(TEXT(AVERAGE(J3:J54),"0.0"),"—")`]] });
  reqs.push({ repeatCell: {
    range: gridRange(WKL,3,4,13,14),
    cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.electricGold) }}},
    fields: 'userEnteredFormat.textFormat',
  }});

  // Row 1 (0-indexed): column headers (GSheets row 2)
  const hdrs = ['WEEK #','WEEK STARTS','WEEK ENDS','GOALS WORKED ON','WINS THIS WEEK','CHALLENGES','NEXT WEEK\'S FOCUS','ENERGY (1–5)','MOTIVATION (1–5)','WEEK RATING (1–5)','QUARTER'];
  reqs.push({ repeatCell: {
    range: gridRange(WKL,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WKL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [hdrs] });

  // 52 weekly data rows (0-indexed 2-53 = GSheets rows 3-54)
  reqs.push({ repeatCell: {
    range: gridRange(WKL,2,54,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.white),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Alternating rows
  for (let i = 0; i < 52; i++) {
    if (i % 2 === 1) {
      reqs.push({ repeatCell: {
        range: gridRange(WKL, 2+i, 3+i, 0, NCOLS),
        cell: { userEnteredFormat: { backgroundColor: hex(C.paleGold) }},
        fields: 'userEnteredFormat.backgroundColor',
      }});
    }
  }
  // Taller rows for reflection
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WKL, dimension: 'ROWS', startIndex: 2, endIndex: 54 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Date format for B and C cols
  reqs.push({ repeatCell: {
    range: gridRange(WKL,2,54,1,3),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }}},
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Center H, I, J, K cols
  reqs.push({ repeatCell: {
    range: gridRange(WKL,2,54,7,11),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(WKL,1,54,0,NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Pre-fill week numbers, formulas for B (week start) and C (week end), K (quarter)
  for (let w = 1; w <= 52; w++) {
    const row = w + 2; // GSheets rows 3-54
    vals.push({ range: `${TAB}!A${row}`, values: [[w]] });
    vals.push({ range: `${TAB}!B${row}`, values: [[`=DATE(2025,1,1)+(A${row}-1)*7`]] });
    vals.push({ range: `${TAB}!C${row}`, values: [[`=B${row}+6`]] });
    vals.push({ range: `${TAB}!K${row}`, values: [[`=IF(A${row}<=13,"Q1",IF(A${row}<=26,"Q2",IF(A${row}<=39,"Q3","Q4")))`]] });
  }

  // Pre-fill weeks 1-6 (fully)
  for (const [wk, goals, wins, challenges, focus, energy, motivation, rating] of WEEKLY_FULL) {
    const row = wk + 2; // GSheets row
    vals.push({ range: `${TAB}!D${row}`, values: [[goals]] });
    vals.push({ range: `${TAB}!E${row}`, values: [[wins]] });
    vals.push({ range: `${TAB}!F${row}`, values: [[challenges]] });
    vals.push({ range: `${TAB}!G${row}`, values: [[focus]] });
    vals.push({ range: `${TAB}!H${row}`, values: [[energy]] });
    vals.push({ range: `${TAB}!I${row}`, values: [[motivation]] });
    vals.push({ range: `${TAB}!J${row}`, values: [[rating]] });
  }

  // Pre-fill weeks 7-26 (scores only)
  for (const [wk, energy, motivation, rating] of WEEKLY_SCORES) {
    const row = wk + 2;
    vals.push({ range: `${TAB}!H${row}`, values: [[energy]] });
    vals.push({ range: `${TAB}!I${row}`, values: [[motivation]] });
    vals.push({ range: `${TAB}!J${row}`, values: [[rating]] });
  }

  await batchUpdate(id, reqs, 'weekly-format');
  await valuesBatchUpdate(id, vals, 'weekly-values');
  console.log('Weekly Check-In & Reflection complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
