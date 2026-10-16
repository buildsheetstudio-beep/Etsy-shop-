'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🎉 Dashboard'];

// Chart 1: Budget by Category — Donut chart on Dashboard
// Data: Dashboard rows 12-21 (0-indexed), col A (categories) and col B (budgeted amounts)
// 1-indexed: rows 13-22, cols A=0, B=1
// Row 12 (0-indexed) = header row: "Category" | "Budgeted" | ...
// Data: rows 12-21 (0-indexed, 10 cats), cols 0-1 (A=category, B=budgeted)

// Chart 2: RSVP Status Counts — Column chart on Dashboard
// Data: Dashboard rows 24-28 (0-indexed), col A (status) and col B (count)
// Row 24 (0-indexed) = RSVP col headers row
// Data rows: 25-28 (0-indexed, 4 statuses)

function src(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Budget Donut ──
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Budget by Category',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.357, green: 0.165, blue: 0.302 } },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.45,
        domain: { sourceRange: { sources: [src(DSH, 11, 22, 0, 1)] } },
        series: { sourceRange: { sources: [src(DSH, 11, 22, 1, 2)] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DSH, rowIndex: 2, columnIndex: 9 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 400, heightPixels: 280,
      },
    },
  }}});

  // ── Chart 2: RSVP Status Column ──
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'RSVP Status Breakdown',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.357, green: 0.165, blue: 0.302 } },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'RSVP Status' },
          { position: 'LEFT_AXIS', title: 'Number of Guests' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(DSH, 24, 29, 0, 1)] } } }],
        series: [{
          series: { sourceRange: { sources: [src(DSH, 24, 29, 1, 2)] } },
          targetAxis: 'LEFT_AXIS',
          color: { red: 0.753, green: 0.553, blue: 0.420 },
        }],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DSH, rowIndex: 13, columnIndex: 9 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 400, heightPixels: 260,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
