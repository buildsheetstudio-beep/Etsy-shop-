'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🎓 Dashboard'];
const TRK = sheetMap['📋 Application Tracker'];
const RES = sheetMap['🏆 Results & Awards Log'];

function src(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Application Status Breakdown — Donut (Dashboard) ──
  // Data: Dashboard rows 28-36 (0-indexed), cols B-C (indices 1-2)
  // statusChartRi = 28 (header), data starts at 29
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Applications by Status',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.102, green: 0.227, blue: 0.545 } },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.5,
        domain: { sourceRange: { sources: [src(DSH, 28, 37, 1, 2)] } },
        series: { sourceRange: { sources: [src(DSH, 28, 37, 2, 3)] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DSH, rowIndex: 2, columnIndex: 6 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 380, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 2: Scholarship Amounts — Column Chart (Application Tracker) ──
  // Data: Tracker rows 3-27 (0-indexed), col A (names, index 0), col D (amounts, index 3)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Scholarship Amounts by Application',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.102, green: 0.227, blue: 0.545 } },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Scholarship' },
          { position: 'LEFT_AXIS',   title: 'Amount ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(TRK, 2, 28, 0, 1)] } } }],
        series: [{
          series: { sourceRange: { sources: [src(TRK, 2, 28, 3, 4)] } },
          targetAxis: 'LEFT_AXIS',
          color: { red: 0.788, green: 0.659, blue: 0.000 }, // #C9A800 gold
        }],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: TRK, rowIndex: 29, columnIndex: 16 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 600, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 3: Awards by Type — Column Chart (Results tab) ──
  // Data: Results rows 29-39 (0-indexed type breakdown), cols A-B (index 0-1 for type/total)
  // typeHdrRi = 29, data rows 30-39 (10 types)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Total Awarded by Scholarship Type',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.102, green: 0.227, blue: 0.545 } },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Type' },
          { position: 'LEFT_AXIS',   title: 'Total Awarded ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(RES, 29, 40, 0, 1)] } } }],
        series: [{
          series: { sourceRange: { sources: [src(RES, 29, 40, 2, 3)] } },
          targetAxis: 'LEFT_AXIS',
          color: { red: 0.788, green: 0.659, blue: 0.000 },
        }],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: RES, rowIndex: 22, columnIndex: 5 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 440, heightPixels: 260,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
