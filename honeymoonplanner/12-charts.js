'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const OVW = sheetMap['🌅 Honeymoon Overview'];
const BGT = sheetMap['💰 Budget & Expense Tracker'];

function srcRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function overlayPos(sheetId, row, col, w, h) {
  return {
    overlayPosition: {
      anchorCell: { sheetId, rowIndex: row, columnIndex: col },
      offsetXPixels: 4,
      offsetYPixels: 4,
      widthPixels: w,
      heightPixels: h,
    },
  };
}

(async () => {
  const reqs = [
    // Chart 1: Budget Breakdown Donut — spending by category
    // Data: OVW rows 32-38 (idx 32-38), cols A-B (category + total actual)
    {
      addChart: {
        chart: {
          spec: {
            title: '💰 Spending by Category',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.warmCharcoal) },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND',
              pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(OVW, 32, 39, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(OVW, 32, 39, 1, 2)] } },
            },
            backgroundColor: hex(C.warmIvory),
          },
          position: overlayPos(OVW, 18, 0, 380, 300),
        },
      },
    },
    // Chart 2: Budget vs Actual column chart — per category
    {
      addChart: {
        chart: {
          spec: {
            title: '📊 Budget vs. Actual by Category',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.warmCharcoal) },
            basicChart: {
              chartType: 'COLUMN',
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: 'Category' },
                { position: 'LEFT_AXIS', title: 'Amount ($)' },
              ],
              domains: [{ domain: { sourceRange: { sources: [srcRange(BGT, 0, 26, 0, 1)] } } }],
              series: [
                {
                  series: { sourceRange: { sources: [srcRange(BGT, 0, 26, 2, 3)] } },
                  targetAxis: 'LEFT_AXIS',
                  color: hex(C.amber),
                },
                {
                  series: { sourceRange: { sources: [srcRange(BGT, 0, 26, 3, 4)] } },
                  targetAxis: 'LEFT_AXIS',
                  color: hex(C.dustyPeach),
                },
              ],
              stackedType: 'NOT_STACKED',
              headerCount: 1,
            },
            backgroundColor: hex(C.warmIvory),
          },
          position: overlayPos(OVW, 18, 4, 550, 300),
        },
      },
    },
  ];

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
