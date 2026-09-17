'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Tab IDs for reference
const DASH = sheetMap['Dashboard'];
const INC  = sheetMap['Income & Orders'];
const EXP  = sheetMap['Expense Log'];
const MO   = sheetMap['Monthly Dashboard'];
const PROD = sheetMap['Product Profitability'];

function sourceRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Revenue vs Expenses by Month (Combo chart on Dashboard) ──────
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Monthly Revenue vs Expenses — 2026',
      titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'COMBO',
        legendPosition: 'BOTTOM_LEGEND',
        series: [
          {
            series: { sourceRange: { sources: [sourceRange(DASH, 23, 37, 1, 2)] } },
            type: 'COLUMN',
            color: hex(C.success),
          },
          {
            series: { sourceRange: { sources: [sourceRange(DASH, 23, 37, 2, 3)] } },
            type: 'COLUMN',
            color: hex(C.attention),
          },
          {
            series: { sourceRange: { sources: [sourceRange(DASH, 23, 37, 3, 4)] } },
            type: 'LINE',
            color: hex(C.primary),
          },
        ],
        domains: [{ domain: { sourceRange: { sources: [sourceRange(DASH, 23, 37, 0, 1)] } } }],
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DASH, rowIndex: 38, columnIndex: 0 },
        widthPixels: 700, heightPixels: 360,
        offsetXPixels: 0, offsetYPixels: 10,
      },
    },
  }}});

  // ── Chart 2: Expense Category Donut on Dashboard ─────────────────────────
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Expenses by Category — YTD',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.4,
        series: { sourceRange: { sources: [sourceRange(DASH, 23, 37, 2, 3)] } },
        domain: { sourceRange: { sources: [sourceRange(DASH, 23, 37, 0, 1)] } },
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DASH, rowIndex: 38, columnIndex: 7 },
        widthPixels: 400, heightPixels: 360,
        offsetXPixels: 0, offsetYPixels: 10,
      },
    },
  }}});

  // ── Chart 3: Income by Source Donut on Income & Orders ──────────────────
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Revenue by Source — 2026',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.4,
        series: { sourceRange: { sources: [sourceRange(INC, 4, 500, 7, 8)] } },
        domain: { sourceRange: { sources: [sourceRange(INC, 4, 500, 5, 6)] } },
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: INC, rowIndex: 4, columnIndex: 19 },
        widthPixels: 420, heightPixels: 320,
        offsetXPixels: 0, offsetYPixels: 10,
      },
    },
  }}});

  // ── Chart 4: Product Revenue Bar on Product Profitability ────────────────
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Gross Revenue by Listing',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'BOTTOM_LEGEND',
        series: [
          { series: { sourceRange: { sources: [sourceRange(PROD, 3, 23, 6, 7)] } }, color: hex(C.success) },
          { series: { sourceRange: { sources: [sourceRange(PROD, 3, 23, 10, 11)] } }, color: hex(C.primary) },
        ],
        domains: [{ domain: { sourceRange: { sources: [sourceRange(PROD, 3, 23, 2, 3)] } } }],
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: PROD, rowIndex: 3, columnIndex: 13 },
        widthPixels: 560, heightPixels: 400,
        offsetXPixels: 0, offsetYPixels: 10,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'finalize-charts');

  // ── Final: move Dashboard to front (index 0) ────────────────────────────
  const moveReqs = [];
  moveReqs.push({ updateSheetProperties: { properties: { sheetId: DASH, index: 0 }, fields: 'index' } });
  await batchUpdate(id, moveReqs, 'finalize-reorder');

  console.log('Finalize complete');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
