'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// Columns:
//   A: Relationship/Group (5 items)
//   B: Delivery Method (3 items)
//   C: Gift Idea Priority (3 items)
//   D: Holiday Card Status (3 items)
//   E: Return/Exchange Status (4 items)
//   F: Store Name  G: Return Window (days)  (6 stores)

const data = [
  ['Relationship/Group', 'Delivery Method', 'Priority',  'Card Status',    'Return Status',   'Store Name', 'Return Window (days)'],
  ['Family',             'In Person',        'High',      'Not Sent',       'N/A',             'Amazon',      30],
  ['Extended Family',    'Mail',             'Medium',    'Sent',           'Pending Return',  'Target',      90],
  ['Friends',            'Pickup',           'Low',       'Received',       'Returned',        'Best Buy',    15],
  ['Coworkers',          '',                 '',          '',               'Exchanged',       'Walmart',     90],
  ['Other',              '',                 '',          '',               '',                'Etsy',        14],
  ['',                   '',                 '',          '',               '',                'Other',       30],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📋 Reference Data'!A1", values: data },
  ], 'ref-values');

  const reqs = [];

  // Header row
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepCranberry),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // Data rows
  for (let r = 1; r < data.length; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(REF, r, r + 1, 0, 7),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 1 ? C.snowWhite : C.altRow),
            textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Return window column G — number format
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 7, 6, 7),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  [130, 120, 80, 100, 120, 110, 160].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
