'use strict';
const { batchUpdate, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];
const GRT = sheetMap['🎁 Gift Recipient Tracker'];

// Chart data table positions on Dashboard (0-based row indices):
//   Bought donut:  rows 31-32 in sheet → idx 30-31
//   Wrapped donut: rows 33-34 in sheet → idx 32-33
//   Sent donut:    rows 35-36 in sheet → idx 34-35
// Budget vs Actual: references GRT directly (rows 2-26, cols A,E,F = idx 0,4,5)

function srcRange(sheetId, startRow, endRow, startCol, endCol) {
  return { sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol };
}

function overlayPos(sheetId, row, col, widthPx, heightPx) {
  return {
    overlayPosition: {
      anchorCell: { sheetId, rowIndex: row, columnIndex: col },
      offsetXPixels: 4,
      offsetYPixels: 4,
      widthPixels: widthPx,
      heightPixels: heightPx,
    },
  };
}

function donutSpec(title, domainSrc, seriesSrc) {
  return {
    title,
    titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
    pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.5,
      domain: { sourceRange: { sources: [domainSrc] } },
      series: { sourceRange: { sources: [seriesSrc] } },
    },
    backgroundColor: hex(C.snowWhite),
  };
}

(async () => {
  const reqs = [];

  // ── Chart 1: DONUT — Gifts Bought ────────────────────────────────────────
  // Data: DSH idx rows 30-31 (idx 30, endRow 32), cols 0-1
  reqs.push({
    addChart: {
      chart: {
        spec: donutSpec(
          '🎁 Gifts Bought',
          srcRange(DSH, 30, 32, 0, 1),
          srcRange(DSH, 30, 32, 1, 2),
        ),
        position: overlayPos(DSH, 13, 0, 350, 270),
      },
    },
  });

  // ── Chart 2: DONUT — Gifts Wrapped ───────────────────────────────────────
  // Data: DSH idx rows 32-33 (idx 32, endRow 34), cols 0-1
  reqs.push({
    addChart: {
      chart: {
        spec: donutSpec(
          '🎀 Gifts Wrapped',
          srcRange(DSH, 32, 34, 0, 1),
          srcRange(DSH, 32, 34, 1, 2),
        ),
        position: overlayPos(DSH, 13, 4, 350, 270),
      },
    },
  });

  // ── Chart 3: DONUT — Gifts Sent/Delivered ────────────────────────────────
  // Data: DSH idx rows 34-35 (idx 34, endRow 36), cols 0-1
  reqs.push({
    addChart: {
      chart: {
        spec: donutSpec(
          '📦 Gifts Delivered',
          srcRange(DSH, 34, 36, 0, 1),
          srcRange(DSH, 34, 36, 1, 2),
        ),
        position: overlayPos(DSH, 20, 0, 350, 270),
      },
    },
  });

  // ── Chart 4: COLUMN — Budget vs. Actual by Recipient ─────────────────────
  // Domains: GRT col A rows 2-26 (idx 1-25, endRow 26)
  // Series 1 (Budget): GRT col E (idx 4)
  // Series 2 (Actual): GRT col F (idx 5)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '💰 Budget vs. Actual Spend by Recipient',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Recipient' },
              { position: 'LEFT_AXIS', title: 'Amount ($)' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(GRT, 1, 26, 0, 1)] } } }],
            series: [
              {
                series: { sourceRange: { sources: [srcRange(GRT, 1, 26, 4, 5)] } },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.warmGold),
              },
              {
                series: { sourceRange: { sources: [srcRange(GRT, 1, 26, 5, 6)] } },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.deepCranberry),
              },
            ],
            stackedType: 'NOT_STACKED',
            headerCount: 1,
          },
          backgroundColor: hex(C.snowWhite),
        },
        position: overlayPos(DSH, 20, 4, 550, 290),
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
