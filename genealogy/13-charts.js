'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

function srcRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function overlayPos(sheetId, row, col, w, h) {
  return {
    overlayPosition: {
      anchorCell: { sheetId, rowIndex: row, columnIndex: col },
      offsetXPixels: 4, offsetYPixels: 4,
      widthPixels: w, heightPixels: h,
    },
  };
}

(async () => {
  const reqs = [
    // Chart 1: Column — Ancestors by Era of Birth
    // Data: DSH rows 16-22 (idx 16,22) cols A-B
    {
      addChart: {
        chart: {
          spec: {
            title: '🌳 Ancestors by Era of Birth',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.sepia) },
            basicChart: {
              chartType: 'COLUMN',
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: 'Era' },
                { position: 'LEFT_AXIS', title: 'Number of Ancestors' },
              ],
              domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 16, 23, 0, 1)] } } }],
              series: [{
                series: { sourceRange: { sources: [srcRange(DSH, 16, 23, 1, 2)] } },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.sepia),
              }],
              stackedType: 'NOT_STACKED',
              headerCount: 1,
            },
            backgroundColor: hex(C.parchment),
          },
          position: overlayPos(DSH, 46, 0, 500, 300),
        },
      },
    },
    // Chart 2: Donut — Ancestors by Birth Location
    // Data: DSH rows 26-32 (idx 26,32) cols A-B
    {
      addChart: {
        chart: {
          spec: {
            title: '🗺️ Ancestors by Birth Country',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.sepia) },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND',
              pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(DSH, 26, 33, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(DSH, 26, 33, 1, 2)] } },
            },
            backgroundColor: hex(C.parchment),
          },
          position: overlayPos(DSH, 46, 4, 450, 300),
        },
      },
    },
    // Chart 3: Donut — DNA/Ethnicity Breakdown
    // Data: DSH rows 36-45 (idx 36,45) cols A-B
    {
      addChart: {
        chart: {
          spec: {
            title: '🧬 DNA Ethnicity Breakdown',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.sepia) },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND',
              pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(DSH, 36, 46, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(DSH, 36, 46, 1, 2)] } },
            },
            backgroundColor: hex(C.parchment),
          },
          position: overlayPos(DSH, 56, 0, 500, 300),
        },
      },
    },
  ];

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
