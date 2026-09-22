'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB   = sheetMap['📋 Dashboard'];
const CK   = sheetMap['✅ Executor Checklist'];
const AD   = sheetMap['💼 Assets & Debts'];
const DIST = sheetMap['🏦 Estate Distribution'];

(async () => {
  const reqs = [];

  // ── Chart 1: Task Status by Phase — Stacked Column on Dashboard ──────────
  // Data: Phase progress table rows 14-19 cols A (phase) and B,C (total, completed)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Executor Task Progress by Phase',
      basicChart: {
        chartType: 'COLUMN',
        stackedType: 'STACKED',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Phase' },
          { position: 'LEFT_AXIS', title: 'Tasks' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: DB, startRowIndex: 13, endRowIndex: 20, startColumnIndex: 0, endColumnIndex: 1 }] } } }],
        series: [
          {
            series: { sourceRange: { sources: [{ sheetId: DB, startRowIndex: 13, endRowIndex: 20, startColumnIndex: 2, endColumnIndex: 3 }] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.mutedSage),
          },
          {
            series: { sourceRange: { sources: [{ sheetId: DB, startRowIndex: 13, endRowIndex: 20, startColumnIndex: 1, endColumnIndex: 2 }] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.stoneBorder),
          },
        ],
        headerCount: 0,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DB, rowIndex: 20, columnIndex: 0 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 600, heightPixels: 320,
      },
    },
  }}});

  // ── Chart 2: Estate Asset Breakdown — Pie on Dashboard ───────────────────
  // Data: Asset descriptions (B5:B16) and values (F5:F16)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Asset Portfolio Breakdown',
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        domain: { sourceRange: { sources: [{ sheetId: AD, startRowIndex: 4, endRowIndex: 16, startColumnIndex: 1, endColumnIndex: 2 }] } },
        series: { sourceRange: { sources: [{ sheetId: AD, startRowIndex: 4, endRowIndex: 16, startColumnIndex: 5, endColumnIndex: 6 }] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DB, rowIndex: 20, columnIndex: 5 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 560, heightPixels: 320,
      },
    },
  }}});

  // ── Chart 3: Distribution by Beneficiary — Column on Distribution tab ────
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Estate Distribution by Beneficiary',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Beneficiary' },
          { position: 'LEFT_AXIS', title: 'Amount ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: DIST, startRowIndex: 11, endRowIndex: 15, startColumnIndex: 0, endColumnIndex: 1 }] } } }],
        series: [
          {
            series: { sourceRange: { sources: [{ sheetId: DIST, startRowIndex: 11, endRowIndex: 15, startColumnIndex: 2, endColumnIndex: 3 }] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.mutedSage),
          },
          {
            series: { sourceRange: { sources: [{ sheetId: DIST, startRowIndex: 11, endRowIndex: 15, startColumnIndex: 3, endColumnIndex: 4 }] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.warmTaupe),
          },
        ],
        headerCount: 0,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DIST, rowIndex: 16, columnIndex: 0 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 560, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 4: Professional Fees — Donut on Professional Services tab ──────
  // Use Pie chart (API doesn't support donut natively, use pieChart with hole)
  const PS = sheetMap['🏛️ Professional Services & Fees'];
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Professional Fee Summary',
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        domain: { sourceRange: { sources: [{ sheetId: PS, startRowIndex: 3, endRowIndex: 8, startColumnIndex: 1, endColumnIndex: 2 }] } },
        series: { sourceRange: { sources: [{ sheetId: PS, startRowIndex: 3, endRowIndex: 8, startColumnIndex: 8, endColumnIndex: 9 }] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: PS, rowIndex: 10, columnIndex: 0 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 500, heightPixels: 280,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
