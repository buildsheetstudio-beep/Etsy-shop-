'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH  = sheetMap['Dashboard'];
const HARV  = sheetMap['Harvest & Yield Tracker'];
const BUD   = sheetMap['Garden Budget'];
const PEST  = sheetMap['Pest & Disease Log'];

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

function hexColor(h) {
  const c = h.replace('#','');
  return { red: parseInt(c.slice(0,2),16)/255, green: parseInt(c.slice(2,4),16)/255, blue: parseInt(c.slice(4,6),16)/255 };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Plant Status Donut (Dashboard) ────────────────────────────────
  // Data in Dashboard rows 23-30 (indices 22-29), cols A-B (0-1)
  // Chart helper starts at chartHelperRow + 2 = (14+9)+2 = 25 → row indices 25-32
  // Actually let me recalculate: uptRow=14, harvRow=14+9=23, chartHelperRow=23+9=32
  // chartHelperRow + 2 = 34 (header), data rows 35-41
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Plant Status Distribution',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hexColor(C.soil) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            threeDimensional: false,
            domain: {
              sourceRange: { sources: [gridRange(DASH, 34, 41, 0, 1)] },
            },
            series: {
              sourceRange: { sources: [gridRange(DASH, 34, 41, 1, 2)] },
            },
          },
          backgroundColor: hexColor(C.cream2),
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DASH, rowIndex: 6, columnIndex: 8 },
            offsetXPixels: 0, offsetYPixels: 0,
            widthPixels: 400, heightPixels: 260,
          },
        },
      },
    },
  });

  // ── Chart 2: Harvest by Plant Column (Harvest tab) ─────────────────────────
  // Monthly harvest data in Harvest tab: chartRow+2 header, chartRow+3 data
  // chartRow = sumRow + 5, sumRow = 2 + 12 + 2 = 16, so chartRow = 21
  // Header at row 22 (index 22), data rows 23-28 (indices 23-28)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Monthly Harvest Totals (lbs)',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hexColor(C.soil) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Pounds Harvested' },
            ],
            domains: [{
              domain: { sourceRange: { sources: [gridRange(HARV, 23, 29, 0, 1)] } },
            }],
            series: [{
              series: { sourceRange: { sources: [gridRange(HARV, 23, 29, 2, 3)] } },
              targetAxis: 'LEFT_AXIS',
              color: hexColor(C.oliveLight),
            }],
            headerCount: 0,
          },
          backgroundColor: hexColor(C.cream2),
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: HARV, rowIndex: 2, columnIndex: 7 },
            offsetXPixels: 0, offsetYPixels: 0,
            widthPixels: 420, heightPixels: 260,
          },
        },
      },
    },
  });

  // ── Chart 3: Budget vs Actual Grouped Column (Budget tab) ─────────────────
  // Category summary in Budget tab:
  // catRow = 3 + 10 + 2 = 15, header at row 16 (index 16), data rows 17-24 (indices 17-24)
  // catRow + 1 (header) = 16, data = 17-24
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Budget vs Actual by Category',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hexColor(C.soil) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'RIGHT_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Category' },
              { position: 'LEFT_AXIS', title: 'Amount ($)' },
            ],
            domains: [{
              domain: { sourceRange: { sources: [gridRange(BUD, 17, 25, 0, 1)] } },
            }],
            series: [
              {
                series: { sourceRange: { sources: [gridRange(BUD, 17, 25, 1, 2)] } },
                targetAxis: 'LEFT_AXIS',
                color: hexColor(C.sage),
              },
              {
                series: { sourceRange: { sources: [gridRange(BUD, 17, 25, 2, 3)] } },
                targetAxis: 'LEFT_AXIS',
                color: hexColor(C.terracotta),
              },
            ],
            headerCount: 0,
          },
          backgroundColor: hexColor(C.cream2),
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: BUD, rowIndex: 2, columnIndex: 9 },
            offsetXPixels: 0, offsetYPixels: 0,
            widthPixels: 440, heightPixels: 280,
          },
        },
      },
    },
  });

  // ── Chart 4: Pest by Type Donut (Pest tab) ────────────────────────────────
  // Chart helper at: chartRow = sumRow + 5, sumRow = 2 + 6 + 2 = 10, chartRow = 15
  // Header at row 16 (index 16), data rows 17-22 (indices 17-22)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Pest & Disease by Type',
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hexColor(C.soil) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            threeDimensional: false,
            domain: {
              sourceRange: { sources: [gridRange(PEST, 17, 23, 0, 1)] },
            },
            series: {
              sourceRange: { sources: [gridRange(PEST, 17, 23, 1, 2)] },
            },
          },
          backgroundColor: hexColor(C.cream2),
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: PEST, rowIndex: 2, columnIndex: 9 },
            offsetXPixels: 0, offsetYPixels: 0,
            widthPixels: 380, heightPixels: 250,
          },
        },
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete — 4 charts (2 donuts, 2 column)');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
