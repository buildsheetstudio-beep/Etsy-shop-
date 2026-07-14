'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// 7 columns:
//   A: Relationship  B: Delivery Method  C: Priority  D: Card Status
//   E: Return Status  F: Store Name  G: Return Window (days)

const relationships = ['Spouse/Partner','Parent','Child','Sibling','Grandparent','Friend','Coworker','Neighbor','Other'];
const delivery      = ['Ship to Home','In Person','Store Pickup','Digital/Download','Mail'];
const priority      = ['High','Medium','Low'];
const cardStatus    = ['Not Sent','Sent','Received'];
const returnStatus  = ['Pending','Initiated','In Progress','Completed','Denied'];
const stores        = ['Amazon','Target','Walmart','Best Buy','Etsy','Macy\'s','Nordstrom','HomeGoods','Other'];
const windows       = [30, 30, 90, 15, 30, 60, 60, 30, 30];

(async () => {
  // ── Values ────────────────────────────────────────────────────────────────
  const maxRows = Math.max(relationships.length, delivery.length, priority.length,
    cardStatus.length, returnStatus.length, stores.length);

  const rows = [];
  // Header
  rows.push(['Relationship','Delivery Method','Priority','Card Status','Return Status','Store Name','Return Window (days)']);
  for (let i = 0; i < maxRows; i++) {
    rows.push([
      relationships[i] || '',
      delivery[i]      || '',
      priority[i]      || '',
      cardStatus[i]    || '',
      returnStatus[i]  || '',
      stores[i]        || '',
      windows[i]       || '',
    ]);
  }

  await valuesBatchUpdate(id, [
    { range: "'📋 Reference Data'!A1", values: rows },
  ], 'ref-values');

  // ── Formatting ────────────────────────────────────────────────────────────
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
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Data rows alternating
  for (let r = 1; r <= maxRows; r++) {
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

  // Column widths
  [130, 130, 80, 100, 110, 110, 160].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
