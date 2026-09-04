'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🏠 Dashboard'];
const TT   = sheetMap['✅ Task Tracker'];
const BUD  = sheetMap['💰 Moving Budget'];
const PT   = sheetMap['📦 Packing Tracker'];

(async () => {
  const reqs = [];

  // ── Chart 1: Budget vs Actual — Column chart on Moving Budget tab ──────────
  // Data: categories in B6:B15, Budgeted D6:D15, Actual E6:E15
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Budget vs Actual by Category',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Category' },
              { position: 'LEFT_AXIS',   title: 'Amount ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: BUD, startRowIndex: 5, endRowIndex: 15, startColumnIndex: 1, endColumnIndex: 2 }],
                },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: BUD, startRowIndex: 5, endRowIndex: 15, startColumnIndex: 3, endColumnIndex: 4 }],
                  },
                },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.deepForest),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: BUD, startRowIndex: 5, endRowIndex: 15, startColumnIndex: 4, endColumnIndex: 5 }],
                  },
                },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.antiqueBrass),
              },
            ],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: BUD, rowIndex: 17, columnIndex: 0 },
            offsetXPixels: 0, offsetYPixels: 0,
            widthPixels: 600, heightPixels: 320,
          },
        },
      },
    },
  });

  // ── Chart 2: Task Status — Pie/Donut on Task Tracker tab ─────────────────
  // Data from summary row 3: B3(Done), C3(In Progress), D3(Not Started), E3(Blocked)
  // Use fixed data labels since source is row values
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Task Status Overview',
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.4,
            domain: {
              sourceRange: {
                sources: [{ sheetId: TT, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 6 }],
              },
            },
            series: {
              sourceRange: {
                sources: [{ sheetId: TT, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 6 }],
              },
            },
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: TT, rowIndex: 45, columnIndex: 0 },
            offsetXPixels: 0, offsetYPixels: 10,
            widthPixels: 480, heightPixels: 300,
          },
        },
      },
    },
  });

  // ── Chart 3: Packing Progress by Room — Stacked Bar on Packing Tracker ────
  // Room summary data in K3:P10 — rooms in col K, Packed/Loaded/Arrived/Unpacked in M/N/O/P
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Packing Progress by Room',
          basicChart: {
            chartType: 'BAR',
            stackedType: 'STACKED',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Boxes' },
              { position: 'LEFT_AXIS',   title: 'Room' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: PT, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 10, endColumnIndex: 11 }],
                },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: PT, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 12, endColumnIndex: 13 }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hex(C.amber),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: PT, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 13, endColumnIndex: 14 }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hex(C.forestGreen),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: PT, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 14, endColumnIndex: 15 }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hex(C.deepForest),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: PT, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 15, endColumnIndex: 16 }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hex(C.mossGreen),
              },
            ],
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: PT, rowIndex: 22, columnIndex: 0 },
            offsetXPixels: 0, offsetYPixels: 10,
            widthPixels: 580, heightPixels: 320,
          },
        },
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
