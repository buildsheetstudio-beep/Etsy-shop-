'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BDY = sheetMap['📏 Body Measurements & Progress Photos'];

// Columns: A=Date, B=Weight, C=Body Fat %, D=Chest, E=Waist, F=Hips, G=Arms, H=Thighs, I=Photo Link (formula), J=Notes
// Photo link formula: =IFERROR(IF(K2="","",HYPERLINK(K2,"View Photo")),"") with K = raw URL (hidden-ish)
const HEADERS = ['Date', 'Weight (kg)', 'Body Fat % (opt.)', 'Chest (cm)', 'Waist (cm)', 'Hips (cm)', 'Arms (cm)', 'Thighs (cm)', 'View Photo', 'Photo URL (paste here)', 'Notes'];

const ROWS = [
  ['=DATE(2025,1,1)',  78.5, 18.2, 96, 82, 94, 33, 57, '', '', 'Starting measurements — New Year'],
  ['=DATE(2025,2,1)',  77.8, 17.9, 95, 81, 93, 33, 57, '', '', 'Down 0.7kg — slow and steady'],
  ['=DATE(2025,3,1)',  77.2, 17.5, 95, 80, 93, 34, 57, '', '', 'Arms growing — waist dropping'],
  ['=DATE(2025,4,1)',  76.5, 17.1, 96, 79, 92, 34, 56, '', '', 'Body recomposition in progress'],
  ['=DATE(2025,5,1)',  76.0, 16.8, 96, 78, 92, 35, 56, '', '', 'Best measurements yet — keep it up'],
  ['=DATE(2025,6,1)',  75.5, 16.4, 97, 77, 91, 35, 55, '', '', 'Visible progress in the mirror!'],
];

(async () => {
  const dataRows = ROWS.map((row, i) => {
    const r = i + 2;
    return [row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7],
      `=IFERROR(IF(J${r}="","",HYPERLINK(J${r},"View Photo")),"")`,
      row[9], row[10]];
  });

  await valuesBatchUpdate(id, [
    { range: "'📏 Body Measurements & Progress Photos'!A1:K1", values: [HEADERS] },
    { range: "'📏 Body Measurements & Progress Photos'!A2:K7", values: dataRows },
  ], 'bdy-values');

  const reqs = [];

  // Base — neutral, no green/red CF
  reqs.push({ repeatCell: { range: gridRange(BDY,0,20,0,11), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Header
  reqs.push({ repeatCell: { range: gridRange(BDY,0,1,0,11), cell: { userEnteredFormat: { backgroundColor: hex('#7B5EA7'), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });

  // Alternating rows
  for (let i = 1; i < 7; i += 2) {
    reqs.push({ repeatCell: { range: gridRange(BDY,i,i+1,0,11), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  }

  // Date — date format
  reqs.push({ repeatCell: { range: gridRange(BDY,1,7,0,1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Measurement columns — centered
  for (let ci = 1; ci <= 7; ci++) {
    reqs.push({ repeatCell: { range: gridRange(BDY,1,7,ci,ci+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0' } } }, fields: 'userEnteredFormat(horizontalAlignment,numberFormat)' } });
  }

  // Photo link col (idx 8) — teal link style
  reqs.push({ repeatCell: { range: gridRange(BDY,1,7,8,9), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex('#3F8CB5'), bold: true }, horizontalAlignment: 'CENTER', backgroundColor: hex(C.formulaBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,backgroundColor)' } });

  // Photo URL col (idx 9) — small gray text
  reqs.push({ repeatCell: { range: gridRange(BDY,1,7,9,10), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.cobalt), fontSize: 9 }, backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,backgroundColor)' } });

  // Input bg for data cols
  reqs.push({ repeatCell: { range: gridRange(BDY,1,7,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(BDY,1,7,10,11), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: BDY, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BDY, dimension: 'ROWS', startIndex: 1, endIndex: 7 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Column widths
  [[0,110],[1,100],[2,110],[3,90],[4,90],[5,90],[6,90],[7,90],[8,100],[9,160],[10,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: BDY, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'bdy-format');
  console.log('Body Measurements complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
