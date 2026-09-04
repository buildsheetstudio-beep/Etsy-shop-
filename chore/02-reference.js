'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['Reference Data'];

(async () => {
  const data = [
    { range: "'Reference Data'!A1:B1", values: [['Chore Type', 'Chore Category']] },
    { range: "'Reference Data'!A2:A3", values: [['Hourly'], ['Flat Rate']] },
    { range: "'Reference Data'!B2:B7", values: [['Cleaning'], ['Pets'], ['Homework'], ['Yard Work'], ['Kitchen'], ['Other']] },
  ];
  await valuesBatchUpdate(id, data, 'ref-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(REF,0,1,0,2), cell: { userEnteredFormat: { backgroundColor: hex(C.turquoise), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  for (let r = 1; r <= 6; r++) {
    reqs.push({ repeatCell: { range: gridRange(REF,r,r+1,0,2), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }
  [[0,100],[1,120]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
