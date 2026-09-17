'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BE = sheetMap['Budget & Expenses'];

(async () => {
  const chartReqs = [];

  // Chart 1: Donut — Actual Spending by Category
  // H col = index 7 (category label), I col = index 8 (amount) — rows 11-26 (index 10-26)
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Actual Spending by Category',
      pieChart: {
        legendPosition:'RIGHT_LEGEND',
        pieHole:0.4,
        domain: { sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:7, endColumnIndex:8 }]}},
        series: { sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:8, endColumnIndex:9 }]}}
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:4, columnIndex:9 }, widthPixels:420, heightPixels:250 } }
  }}});

  // Chart 2: Grouped Column — Planned vs Actual (A-C cols 0-2, rows 10-26)
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Planned vs. Actual by Category',
      basicChart: {
        chartType:'COLUMN',
        legendPosition:'RIGHT_LEGEND',
        domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:0, endColumnIndex:1 }]}}}],
        series:[
          { series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:1, endColumnIndex:2 }]}}, targetAxis:'LEFT_AXIS' },
          { series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:2, endColumnIndex:3 }]}}, targetAxis:'LEFT_AXIS' },
        ],
        headerCount:1
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:4, columnIndex:9 }, widthPixels:450, heightPixels:250, offsetXPixels:430 } }
  }}});

  // Chart 3: Bar — Payment Status (K-L cols 10-11, rows 11-14 index 11-14)
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Expenses by Payment Status',
      basicChart: {
        chartType:'BAR',
        legendPosition:'BOTTOM_LEGEND',
        domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:11, endRowIndex:15, startColumnIndex:10, endColumnIndex:11 }]}}}],
        series:[{ series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:11, endRowIndex:15, startColumnIndex:11, endColumnIndex:12 }]}}, targetAxis:'BOTTOM_AXIS' }],
        headerCount:1
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:9, columnIndex:9 }, widthPixels:400, heightPixels:220 } }
  }}});

  await batchUpdate(id, chartReqs);
  console.log('Budget charts complete');
})();
