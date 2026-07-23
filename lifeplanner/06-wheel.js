'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WHL = sheetMap['⚖️ Life Balance Wheel'];

const HEADERS = ['Life Category','Current Satisfaction (1-10)','Goal Satisfaction (1-10)','Notes'];
const ROWS = [
  ['Personal Growth',7,9,'Reading more, learning Python. Want to add mentorship.'],
  ['Health & Wellness',5,8,'Sleep and exercise inconsistent. Clear priority gap.'],
  ['Career & Business',7,8,'Stable job, exploring side income. Getting there.'],
  ['Home & Household',6,7,'Functional but cluttered. Want more calm environment.'],
  ['Relationships & Social',6,8,'Good friendships, could be more intentional.'],
  ['Finance',5,9,'Budget improving but emergency fund a long way off.'],
];

(async () => {
  if (ROWS.length !== 6) throw new Error(`Expected 6 rows, got ${ROWS.length}`);
  const data = [
    { range: "'⚖️ Life Balance Wheel'!A1:D1", values: [HEADERS] },
    { range: "'⚖️ Life Balance Wheel'!A2:D7", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'whl-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(WHL,0,1,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: WHL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  for (let r = 1; r <= 6; r++) {
    reqs.push({ repeatCell: { range: gridRange(WHL,r,r+1,0,4), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }
  reqs.push({ repeatCell: { range: gridRange(WHL,1,7,1,3), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true, fontSize: 14 } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(WHL,1,7,3,4), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });
  [[0,170],[1,180],[2,180],[3,320]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: WHL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: WHL, dimension: 'ROWS', startIndex: 1, endIndex: 7 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  await batchUpdate(id, reqs, 'whl-format');
  console.log('Life Balance Wheel complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
