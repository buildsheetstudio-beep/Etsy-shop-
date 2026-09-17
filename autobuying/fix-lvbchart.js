'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const LVB = sheetMap['Lease vs Buy'];
const S   = "'Lease vs Buy'";

// A 3-row helper lookup table tucked in columns F:G, rows 24-26 (1-indexed).
// These cells pull from the comparison metrics and feed the chart cleanly.
//   F24: "Category"          G24: "Total Cash Outflow"   ← header row (headerCount: 1)
//   F25: "Lease"             G25: =C36                   ← lease total cash outflow
//   F26: "Buy"               G26: =D36                   ← buy total cash outflow
//
// Chart domain: F24:F26 (0-indexed rows 23-26, col 5)
// Chart series: G24:G26 (0-indexed rows 23-26, col 6)

(async () => {
  // 1. Write helper lookup table (rows 24-26, cols F-G)
  await valuesBatchUpdate(id, [
    { range: `${S}!F24`, values: [['Category']] },
    { range: `${S}!G24`, values: [['Total Cash Outflow']] },
    { range: `${S}!F25`, values: [['Lease']] },
    { range: `${S}!G25`, values: [['=C36']] },
    { range: `${S}!F26`, values: [['Buy']] },
    { range: `${S}!G26`, values: [['=D36']] },
  ], 'lvb-chart-helper');

  // 2. Add chart — domain F24:F26, series G24:G26, headerCount: 1
  //    Produces 2 bars: X-axis "Lease" / "Buy", Y-axis = total cash outflow
  await batchUpdate(id, [{
    addChart: {
      chart: {
        spec: {
          title: 'Lease vs Buy — Total Cash Outflow',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS' },
              { position: 'LEFT_AXIS', title: 'Total Cost ($)' },
            ],
            domains: [{
              domain: { sourceRange: { sources: [{
                sheetId: LVB,
                startRowIndex: 23, endRowIndex: 26,
                startColumnIndex: 5, endColumnIndex: 6,
              }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{
                sheetId: LVB,
                startRowIndex: 23, endRowIndex: 26,
                startColumnIndex: 6, endColumnIndex: 7,
              }] } },
              type: 'COLUMN',
              color: hex(C.secondary),
            }],
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: LVB, rowIndex: 44, columnIndex: 0 },
            widthPixels: 480,
            heightPixels: 280,
          },
        },
      },
    },
  }], 'lvb-chart');

  console.log('✓ Lease vs Buy — Total Cash Outflow chart fixed');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
