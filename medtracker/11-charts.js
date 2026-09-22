'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Dashboard data table row positions (0-based):
// Category data:    rows 9-21  (labels in col A=0, amounts in col B=1)
// Status data:      rows 24-29 (labels in col A=0, amounts in col B=1)
// Monthly data:     rows 32-43 (labels in col A=0, amounts in col B=1)
// Participant data: rows 46-49 (labels in col A=0, amounts in col B=1)

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

  // ── Chart 1: Donut — Expenses by Category ───────────────────────────────
  // Data rows 9-21 (includes header row 9 which has "Category"/"Amount" labels)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '🗂 Expenses by Category',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.deepTeal) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.5,
            domain: { sourceRange: { sources: [srcRange(9, 22, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(9, 22, 1, 2)] } },
          },
          backgroundColor: hex(C.softIvory),
        },
        position: overlayPos(8, 4, 420, 310),
      },
    },
  });

  // ── Chart 2: Donut — Reimbursement Status ───────────────────────────────
  // Data rows 24-29 (includes header row 24)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '📋 Reimbursement Status',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.deepTeal) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.5,
            domain: { sourceRange: { sources: [srcRange(24, 30, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(24, 30, 1, 2)] } },
          },
          backgroundColor: hex(C.softIvory),
        },
        position: overlayPos(23, 4, 420, 270),
      },
    },
  });

  // ── Chart 3: Column — Monthly Totals ────────────────────────────────────
  // Data rows 32-43 (includes header row 32)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '📅 Monthly Medical Expenses',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.deepTeal) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Amount ($)' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(32, 44, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(32, 44, 1, 2)] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.warmCoral),
            }],
            stackedType: 'NOT_STACKED',
          },
          backgroundColor: hex(C.softIvory),
        },
        position: overlayPos(31, 4, 500, 270),
      },
    },
  });

  // ── Chart 4: Column — By Participant ────────────────────────────────────
  // Data rows 46-49 (includes header row 46)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '👥 Expenses by Participant',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.deepTeal) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Participant' },
              { position: 'LEFT_AXIS', title: 'Amount ($)' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(46, 50, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(46, 50, 1, 2)] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.mutedSage),
            }],
            stackedType: 'NOT_STACKED',
          },
          backgroundColor: hex(C.softIvory),
        },
        position: overlayPos(45, 4, 380, 250),
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
