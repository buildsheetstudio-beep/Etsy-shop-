'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];

// Layout (1-indexed rows):
//  1    : Block A column headers
//  2-7  : 6 sample chores
//  8-9  : empty input rows
// 10    : Block B section label (merged)
// 11    : Block B column headers
// 12-21 : 10 family members
// 22    : blank separator
// 23    : Block C section label (merged)
// 24    : Block C column headers
// 25-64 : 40 log entries (data)

const MEMBERS = ['Emma','Liam','Sofia','Noah','Olivia','Ethan','Ava','Lucas','Mia','Jackson'];

const CHORES = [
  ['Feed the Pet','Pets','Hourly',3.50],
  ['Mow the Lawn','Yard Work','Flat Rate',8.00],
  ['Wash Dishes','Kitchen','Flat Rate',2.00],
  ['Do Laundry','Cleaning','Hourly',1.50],
  ['Homework Check-in','Homework','Hourly',0.75],
  ['Make Bed','Cleaning','Flat Rate',1.00],
];

const LOG = [
  ['2025-06-01','Emma','Make Bed',1,true],
  ['2025-06-01','Liam','Wash Dishes',1,true],
  ['2025-06-02','Sofia','Feed the Pet',1,true],
  ['2025-06-02','Noah','Homework Check-in',2,true],
  ['2025-06-03','Olivia','Do Laundry',1.5,true],
  ['2025-06-03','Ethan','Mow the Lawn',1,true],
  ['2025-06-04','Ava','Make Bed',1,false],
  ['2025-06-04','Lucas','Wash Dishes',1,false],
  ['2025-06-05','Mia','Feed the Pet',1,false],
  ['2025-06-05','Jackson','Homework Check-in',1,false],
  ['2025-06-06','Emma','Do Laundry',2,true],
  ['2025-06-06','Liam','Mow the Lawn',1,true],
  ['2025-06-07','Sofia','Make Bed',1,true],
  ['2025-06-07','Noah','Wash Dishes',1,true],
  ['2025-06-08','Olivia','Feed the Pet',1.5,true],
  ['2025-06-09','Ethan','Homework Check-in',1,false],
  ['2025-06-10','Ava','Do Laundry',1,false],
  ['2025-06-10','Lucas','Mow the Lawn',1,false],
  ['2025-06-11','Mia','Make Bed',1,true],
  ['2025-06-11','Jackson','Wash Dishes',1,true],
  ['2025-06-12','Emma','Feed the Pet',2,true],
  ['2025-06-13','Liam','Homework Check-in',2,true],
  ['2025-06-14','Sofia','Do Laundry',2,false],
  ['2025-06-15','Noah','Mow the Lawn',1,false],
  ['2025-06-16','Olivia','Make Bed',1,false],
  ['2025-06-17','Ethan','Wash Dishes',1,true],
  ['2025-06-18','Ava','Feed the Pet',1,true],
  ['2025-06-19','Lucas','Homework Check-in',1.5,true],
  ['2025-06-20','Mia','Do Laundry',1,true],
  ['2025-06-20','Jackson','Mow the Lawn',1,false],
  ['2025-06-21','Emma','Wash Dishes',1,false],
  ['2025-06-22','Liam','Make Bed',1,false],
  ['2025-06-23','Sofia','Mow the Lawn',1,true],
  ['2025-06-24','Noah','Feed the Pet',1,true],
  ['2025-06-25','Olivia','Homework Check-in',2,true],
  ['2025-06-26','Ethan','Do Laundry',1.5,true],
  ['2025-06-27','Ava','Wash Dishes',1,false],
  ['2025-06-28','Lucas','Make Bed',1,false],
  ['2025-06-29','Mia','Homework Check-in',1,true],
  ['2025-06-30','Jackson','Feed the Pet',2,true],
];

(async () => {
  const data = [];

  // Block A header
  data.push({ range: "'Chore & Allowance Tracker'!A1:D1", values: [['Chore Name','Category','Type','Rate ($/hr or flat)']] });

  // Block A: 6 chores + 2 empty input rows
  const choreRows = [...CHORES, ['','','',''], ['','','','']];
  data.push({ range: "'Chore & Allowance Tracker'!A2:D9", values: choreRows });

  // Block B: section label (row 10)
  data.push({ range: "'Chore & Allowance Tracker'!A10", values: [['📊  Weekly Summary — How much has each family member earned this week?']] });

  // Block B: column headers (row 11)
  data.push({ range: "'Chore & Allowance Tracker'!A11:D11", values: [['Family Member','Total Hours This Week','Allowance Earned This Week','Total Earned (All Time)']] });

  // Block B: family members + formulas (rows 12-21)
  const ws = '(TODAY()-WEEKDAY(TODAY(),2)+1)';
  const we = '(TODAY()-WEEKDAY(TODAY(),2)+7)';
  MEMBERS.forEach((name, i) => {
    const row = 12 + i;
    const ref = `$A$${row}`;
    data.push({
      range: `'Chore & Allowance Tracker'!A${row}:D${row}`,
      values: [[
        name,
        `=IFERROR(SUMIFS($D$25:$D$1000,$B$25:$B$1000,${ref},$A$25:$A$1000,">="&${ws},$A$25:$A$1000,"<="&${we}),0)`,
        `=IFERROR(SUMIFS($E$25:$E$1000,$B$25:$B$1000,${ref},$A$25:$A$1000,">="&${ws},$A$25:$A$1000,"<="&${we}),0)`,
        `=IFERROR(SUMIFS($E$25:$E$1000,$B$25:$B$1000,${ref}),0)`
      ]]
    });
  });

  // Block C: section label (row 23)
  data.push({ range: "'Chore & Allowance Tracker'!A23", values: [['📝  Chore Log — Add completed chores below. Allowance calculates automatically.']] });

  // Block C: column headers (row 24)
  data.push({ range: "'Chore & Allowance Tracker'!A24:F24", values: [['Date','Family Member','Chore','Hours / Qty','Allowance Earned','Paid Out']] });

  // Block C: 40 log rows (rows 25-64)
  LOG.forEach(([date, member, chore, hrs, paid], i) => {
    const row = 25 + i;
    data.push({
      range: `'Chore & Allowance Tracker'!A${row}:F${row}`,
      values: [[
        date, member, chore, hrs,
        `=IFERROR(IF(VLOOKUP(C${row},$A$2:$D$9,3,FALSE)="Hourly",D${row}*VLOOKUP(C${row},$A$2:$D$9,4,FALSE),VLOOKUP(C${row},$A$2:$D$9,4,FALSE)),0)`,
        paid
      ]]
    });
  });

  await valuesBatchUpdate(id, data, 'tracker-values');

  // ---- FORMATTING ----
  const reqs = [];

  // Block A background: distinct light turquoise tint for rows 1-9 (0-indexed 0-8)
  reqs.push({ repeatCell: { range: gridRange(TAB,0,9,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.blockATint) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Block A header (row 1, 0-indexed 0)
  reqs.push({ repeatCell: { range: gridRange(TAB,0,1,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.turquoise), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Block A data rows (0-indexed 1-8): input bg for A-C, currency for D
  reqs.push({ repeatCell: { range: gridRange(TAB,1,9,0,3), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(TAB,1,9,3,4), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 }, numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 1, endIndex: 9 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Block B section label (row 10, 0-indexed 9): merged A-D
  reqs.push({ mergeCells: { range: gridRange(TAB,9,10,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(TAB,9,10,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.turquoise), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Block B column headers (row 11, 0-indexed 10)
  reqs.push({ repeatCell: { range: gridRange(TAB,10,11,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.marigold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Block B data rows (0-indexed 11-20)
  for (let i = 0; i < 10; i++) {
    const ri = 11 + i;
    const bg = hex(i%2===0 ? C.altRow : C.white);
    reqs.push({ repeatCell: { range: gridRange(TAB,ri,ri+1,0,1), cell: { userEnteredFormat: { backgroundColor: bg, textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(TAB,ri,ri+1,1,2), cell: { userEnteredFormat: { backgroundColor: bg, textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 }, numberFormat: { type: 'NUMBER', pattern: '0.0' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,horizontalAlignment)' } });
    reqs.push({ repeatCell: { range: gridRange(TAB,ri,ri+1,2,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 }, numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,horizontalAlignment)' } });
  }
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 11, endIndex: 21 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Blank row 22 (0-indexed 21)
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 21, endIndex: 22 }, properties: { pixelSize: 12 }, fields: 'pixelSize' } });

  // Block C section label (row 23, 0-indexed 22): merged A-F
  reqs.push({ mergeCells: { range: gridRange(TAB,22,23,0,6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(TAB,22,23,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.turquoise), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 22, endIndex: 23 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Block C column headers (row 24, 0-indexed 23)
  reqs.push({ repeatCell: { range: gridRange(TAB,23,24,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.marigold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 23, endIndex: 24 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Block C data rows (0-indexed 24-63)
  for (let i = 0; i < 40; i++) {
    const ri = 24 + i;
    const bg = hex(i%2===0 ? C.altRow : C.white);
    reqs.push({ repeatCell: { range: gridRange(TAB,ri,ri+1,0,2), cell: { userEnteredFormat: { backgroundColor: bg, textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(TAB,ri,ri+1,2,4), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(TAB,ri,ri+1,4,5), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(TAB,ri,ri+1,5,6), cell: { userEnteredFormat: { backgroundColor: bg, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)' } });
  }
  reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'ROWS', startIndex: 24, endIndex: 65 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Date format: Block C col A (0-indexed 0), rows 25-64 (0-indexed 24-63)
  reqs.push({ repeatCell: { range: gridRange(TAB,24,65,0,1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Hours format: Block C col D (0-indexed 3)
  reqs.push({ repeatCell: { range: gridRange(TAB,24,65,3,4), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths
  [[0,160],[1,150],[2,160],[3,120],[4,140],[5,80]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: TAB, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'tracker-format');
  console.log('Chore & Allowance Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
