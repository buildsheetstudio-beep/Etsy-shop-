'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['Reference Data'];

(async () => {
  const data = [
    { range: "'Reference Data'!A1:B1", values: [['Chore Type','Slot Color']] },
    { range: "'Reference Data'!A2:A3", values: [['Hourly'],['Flat Rate']] },
    {
      range: "'Reference Data'!B2:B11",
      values: C.slot.map((col, i) => [`Slot ${i+1}: ${col}`])
    },
  ];
  await valuesBatchUpdate(id, data, 'ref-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(REF,0,1,0,2), cell: { userEnteredFormat: { backgroundColor: hex(C.chrome), textFormat: { foregroundColor: hex(C.charcoal), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  for (let r = 1; r <= 10; r++) {
    reqs.push({ repeatCell: { range: gridRange(REF,r,r+1,0,2), cell: { userEnteredFormat: { backgroundColor: r%2===1 ? hex(C.chrome) : hex(C.white), textFormat: { foregroundColor: hex(C.charcoal), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    // Color the slot color cell for row 2-11 in col B
    if (r >= 1 && r <= 10) {
      reqs.push({ repeatCell: { range: gridRange(REF,r,r+1,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.slot[r-1]) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }
  [[0,80],[1,140]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
