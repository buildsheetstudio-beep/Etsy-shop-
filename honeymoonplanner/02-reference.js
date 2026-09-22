'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

(async () => {
  // ── Headers + Data ────────────────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    // Col A: Currency codes
    { range: "'📋 Reference Data'!A1", values: [
      ['Currency'],
      ['USD'], ['EUR'], ['GBP'], ['JPY'], ['AUD'], ['CAD'], ['CHF'], ['MXN'], ['THB'], ['GRC'],
    ]},
    // Col B: Checklist Category
    { range: "'📋 Reference Data'!B1", values: [
      ['Checklist Category'],
      ['Flights & Transport'], ['Accommodation'], ['Activities & Excursions'],
      ['Documents & Visas'], ['Budget & Finance'], ['Health & Safety'], ['Packing'], ['General'],
    ]},
    // Col C: Document Type
    { range: "'📋 Reference Data'!C1", values: [
      ['Document Type'],
      ['Flight'], ['Hotel'], ['Car Rental'], ['Tour/Activity'], ['Restaurant'],
      ['Travel Insurance'], ['Visa'], ['Passport'], ['Other'],
    ]},
    // Col D: Booking Status
    { range: "'📋 Reference Data'!D1", values: [
      ['Booking Status'],
      ['Confirmed'], ['Researching'], ['Pending'], ['Cancelled'], ['Not Needed'],
    ]},
    // Col E: Partner
    { range: "'📋 Reference Data'!E1", values: [
      ['Partner'],
      ['Partner 1'], ['Partner 2'], ['Both'],
    ]},
    // Col F: Packing Category
    { range: "'📋 Reference Data'!F1", values: [
      ['Packing Category'],
      ['Clothing'], ['Toiletries & Beauty'], ['Documents & Money'], ['Electronics & Gadgets'],
      ['Health & Medical'], ['Beach & Pool'], ['Entertainment'], ['Miscellaneous'],
    ]},
    // Col G: Activity Category
    { range: "'📋 Reference Data'!G1", values: [
      ['Activity Category'],
      ['Sightseeing'], ['Dining & Food'], ['Spa & Wellness'], ['Adventure & Outdoors'],
      ['Beach & Water'], ['Cultural & Arts'], ['Nightlife'], ['Shopping'], ['Day Trip'],
    ]},
    // Col H: Budget Category
    { range: "'📋 Reference Data'!H1", values: [
      ['Budget Category'],
      ['Flights'], ['Accommodation'], ['Food & Dining'], ['Activities & Tours'],
      ['Transport'], ['Shopping'], ['Health & Beauty'], ['Miscellaneous'],
    ]},
  ], 'ref-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Header row formatting
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.slateTeal),
          textFormat: { foregroundColor: hex(C.warmIvory), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // Data rows
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 20, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warmIvory),
          textFormat: { foregroundColor: hex(C.warmCharcoal), fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // Column widths
  [[0,100],[1,160],[2,140],[3,120],[4,100],[5,160],[6,160],[7,140]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
