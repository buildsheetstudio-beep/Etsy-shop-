'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// 4 Charts on Dashboard:
// 1. Donut — Deductions by Category (Dashboard A10:B22)
// 2. Donut — Documentation Status: Documented vs Not Documented (Dashboard A25:B26)
// 3. Column — Monthly Deduction Totals (Dashboard A29:B40)
// 4. Line — Multi-Year Comparison (Dashboard A43:B45)

function srcRange(sheetId, r1, r2, c1, c2) {
  return {
    sheetId,
    startRowIndex: r1, endRowIndex: r2,
    startColumnIndex: c1, endColumnIndex: c2,
  };
}

function addChart(spec, anchorRow, anchorCol, offsetRow, offsetCol, width, height) {
  return {
    addChart: {
      chart: {
        spec,
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DSH, rowIndex: anchorRow, columnIndex: anchorCol },
            offsetXPixels: offsetRow,
            offsetYPixels: offsetCol,
            widthPixels: width,
            heightPixels: height,
          },
        },
      },
    },
  };
}

(async () => {
  const reqs = [];

  // Chart 1: Donut — Deductions by Category
  reqs.push(addChart(
    {
      title: '2024 Deductions by Category',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepSapphire) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.45,
        series: { sourceRange: { sources: [srcRange(DSH, 9, 22, 1, 2)] } },
        domain: { sourceRange: { sources: [srcRange(DSH, 9, 22, 0, 1)] } },
      },
      backgroundColor: hex(C.warmIvory),
    },
    3, 3, 4, 4, 380, 280
  ));

  // Chart 2: Donut — Documentation Status
  reqs.push(addChart(
    {
      title: 'Documentation Status',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepSapphire) },
      pieChart: {
        legendPosition: 'BOTTOM_LEGEND',
        pieHole: 0.45,
        series: { sourceRange: { sources: [srcRange(DSH, 24, 26, 1, 2)] } },
        domain: { sourceRange: { sources: [srcRange(DSH, 24, 26, 0, 1)] } },
      },
      backgroundColor: hex(C.warmIvory),
    },
    3, 6, 4, 4, 280, 280
  ));

  // Chart 3: Column — Monthly Deduction Totals
  reqs.push(addChart(
    {
      title: '2024 Monthly Deduction Totals',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepSapphire) },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Month' },
          { position: 'LEFT_AXIS', title: 'Amount ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 28, 40, 0, 1)] } } }],
        series: [{
          series: { sourceRange: { sources: [srcRange(DSH, 28, 40, 1, 2)] } },
          targetAxis: 'LEFT_AXIS',
          color: hex(C.warmGold),
        }],
        headerCount: 0,
      },
      backgroundColor: hex(C.warmIvory),
    },
    27, 3, 4, 4, 480, 260
  ));

  // Chart 4: Line — Multi-Year Comparison
  reqs.push(addChart(
    {
      title: 'Multi-Year Deduction Trend',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.deepSapphire) },
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Tax Year' },
          { position: 'LEFT_AXIS', title: 'Total Deductions ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 42, 45, 0, 1)] } } }],
        series: [{
          series: { sourceRange: { sources: [srcRange(DSH, 42, 45, 1, 2)] } },
          targetAxis: 'LEFT_AXIS',
          color: hex(C.deepSapphire),
        }],
        headerCount: 0,
        lineSmoothing: true,
      },
      backgroundColor: hex(C.warmIvory),
    },
    41, 3, 4, 4, 380, 240
  ));

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
