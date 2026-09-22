'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Dashboard data table positions (0-based row indices, sheet rows = +1):
// By Pillar: rows 10-19 (cols A=0 label, B=1 count)
// By Status: rows 22-27 (cols A=0 label, B=1 count)
// Top 10 Downloads: rows 30-39 (col A=0 ep#, C=2 title, D=3 downloads)
// By Month: rows 42-53 (cols A=0 month, B=1 revenue)

function srcRange(startRow, endRow, startCol, endCol) {
  return { sheetId: DSH, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol };
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

  // ── Chart 1: Donut — Episodes by Content Pillar ──────────────────────────
  // Data rows 10-19 (0-based), label col 0, value col 1
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '🎯 Episodes by Content Pillar',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepCharcoal) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.5,
            domain: { sourceRange: { sources: [srcRange(10, 20, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(10, 20, 1, 2)] } },
          },
          backgroundColor: hex(C.warmCream),
        },
        position: overlayPos(8, 4, 450, 300),
      },
    },
  });

  // ── Chart 2: Column — Downloads by Episode (Top 10) ────────────────────
  // Data rows 30-39 (0-based), ep# in col 0, downloads in col 3
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '📥 Total Downloads — Top 10 Episodes',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepCharcoal) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Episode #' },
              { position: 'LEFT_AXIS', title: 'Total Downloads' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(30, 40, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(30, 40, 3, 4)] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.burntOrange),
            }],
            stackedType: 'NOT_STACKED',
          },
          backgroundColor: hex(C.warmCream),
        },
        position: overlayPos(21, 4, 500, 290),
      },
    },
  });

  // ── Chart 3: Column — Sponsorship Revenue by Month ──────────────────────
  // Data rows 42-53 (0-based), month in col 0, revenue in col 1
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '💰 Sponsorship Revenue by Month (2024)',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepCharcoal) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Revenue ($)' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(42, 54, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(42, 54, 1, 2)] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.mutedSage),
            }],
            stackedType: 'NOT_STACKED',
          },
          backgroundColor: hex(C.warmCream),
        },
        position: overlayPos(41, 4, 450, 280),
      },
    },
  });

  // ── Chart 4: Donut — Episode Status Breakdown ───────────────────────────
  // Data rows 22-27 (0-based), label col 0, count col 1
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '📌 Episode Status Breakdown',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepCharcoal) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.5,
            domain: { sourceRange: { sources: [srcRange(22, 28, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(22, 28, 1, 2)] } },
          },
          backgroundColor: hex(C.warmCream),
        },
        position: overlayPos(29, 4, 450, 280),
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
