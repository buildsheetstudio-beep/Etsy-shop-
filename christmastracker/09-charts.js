'use strict';
const { batchUpdate, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Chart data table positions on Dashboard (0-based row indices):
//   Bought donut:  rows 23-25 (idx 22-24)  label col 0, count col 1
//   Wrapped donut: rows 28-29 (idx 27-28)  label col 0, count col 1
//   Sent donut:    rows 32-33 (idx 31-32)  label col 0, count col 1
//   Card status column: rows 36-38 (idx 35-37)  label col 0, count col 1

function srcRange(sheetId, startRow, endRow, startCol, endCol) {
  return { sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol };
}

function overlayPos(row, col, widthPx, heightPx) {
  return {
    overlayPosition: {
      anchorCell: { sheetId: DSH, rowIndex: row, columnIndex: col },
      offsetXPixels: 4,
      offsetYPixels: 4,
      widthPixels: widthPx,
      heightPixels: heightPx,
    },
  };
}

(async () => {
  const reqs = [];

  // ── Chart 1: DONUT — Gift Bought Status ──────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '🎁 Gifts Bought',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.45,
            domain: { sourceRange: { sources: [srcRange(DSH, 23, 25, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(DSH, 23, 25, 1, 2)] } },
          },
          backgroundColor: hex(C.snowWhite),
        },
        position: overlayPos(14, 0, 340, 250),
      },
    },
  });

  // ── Chart 2: DONUT — Gift Wrapped Status ─────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '🎀 Gifts Wrapped',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.45,
            domain: { sourceRange: { sources: [srcRange(DSH, 27, 29, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(DSH, 27, 29, 1, 2)] } },
          },
          backgroundColor: hex(C.snowWhite),
        },
        position: overlayPos(14, 4, 340, 250),
      },
    },
  });

  // ── Chart 3: DONUT — Gift Delivered Status ───────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '📦 Gifts Delivered',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.45,
            domain: { sourceRange: { sources: [srcRange(DSH, 31, 33, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(DSH, 31, 33, 1, 2)] } },
          },
          backgroundColor: hex(C.snowWhite),
        },
        position: overlayPos(20, 0, 340, 250),
      },
    },
  });

  // ── Chart 4: COLUMN — Holiday Card Status ────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '💌 Holiday Card Status',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Status' },
              { position: 'LEFT_AXIS', title: 'Count' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 35, 38, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(DSH, 35, 38, 1, 2)] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.trueRed),
            }],
            stackedType: 'NOT_STACKED',
            headerCount: 0,
          },
          backgroundColor: hex(C.snowWhite),
        },
        position: overlayPos(20, 4, 340, 250),
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
