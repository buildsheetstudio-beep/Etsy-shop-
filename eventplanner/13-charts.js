'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🎉 Dashboard'];
const BUD  = sheetMap['💰 Event Budget'];
const GL   = sheetMap['👥 Guest List & RSVP'];
const ROS  = sheetMap['⏱️ Run-of-Show'];

(async () => {
  const reqs = [];

  // ── Chart 1: Events by Status — Donut on Dashboard ─────────────────────────
  // Data: A29:B34 (Status, Count) on Dashboard
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Events by Status',
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.4,
            domain: {
              sourceRange: {
                sources: [{ sheetId: DASH, startRowIndex: 28, endRowIndex: 35, startColumnIndex: 0, endColumnIndex: 1 }],
              },
            },
            series: {
              sourceRange: {
                sources: [{ sheetId: DASH, startRowIndex: 28, endRowIndex: 35, startColumnIndex: 1, endColumnIndex: 2 }],
              },
            },
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DASH, rowIndex: 35, columnIndex: 0 },
            offsetXPixels: 0, offsetYPixels: 10,
            widthPixels: 480, heightPixels: 300,
          },
        },
      },
    },
  });

  // ── Chart 2: Budget vs Actual — Column chart on Budget tab ─────────────────
  // Data: B4:B15 (categories), E4:E15 (budgeted), F4:F15 (actual)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Budget vs Actual Spend by Category',
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
                  sources: [{ sheetId: BUD, startRowIndex: 3, endRowIndex: 15, startColumnIndex: 2, endColumnIndex: 3 }],
                },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: BUD, startRowIndex: 3, endRowIndex: 15, startColumnIndex: 4, endColumnIndex: 5 }],
                  },
                },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.sageGreen),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: BUD, startRowIndex: 3, endRowIndex: 15, startColumnIndex: 5, endColumnIndex: 6 }],
                  },
                },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.deepCrimson),
              },
            ],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: BUD, rowIndex: 38, columnIndex: 0 },
            offsetXPixels: 0, offsetYPixels: 10,
            widthPixels: 580, heightPixels: 320,
          },
        },
      },
    },
  });

  // ── Chart 3: RSVP Status — Donut on Guest List ─────────────────────────────
  // Data: R2:V2 (headers) + R3:V8 (event RSVP summary) — use Confirmed column
  // Simpler: use summary totals from row 3 of guest list (summary row)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'RSVP Status Overview',
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.4,
            domain: {
              sourceRange: {
                sources: [{ sheetId: GL, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 6 }],
              },
            },
            series: {
              sourceRange: {
                sources: [{ sheetId: GL, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 6 }],
              },
            },
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: GL, rowIndex: 30, columnIndex: 17 },
            offsetXPixels: 0, offsetYPixels: 10,
            widthPixels: 480, heightPixels: 300,
          },
        },
      },
    },
  });

  // ── Chart 4: Run-of-Show duration — Bar chart on Run-of-Show ───────────────
  // Data: B5:B18 (segment names), G5:G18 (duration) — but duration is text
  // Use start time E and end time F as bar chart (gantt-style)
  // Simpler: just segment name + duration as a bar
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Tech Summit 2025 — Run-of-Show Timeline',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Time' },
              { position: 'LEFT_AXIS',   title: 'Segment' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: ROS, startRowIndex: 4, endRowIndex: 18, startColumnIndex: 1, endColumnIndex: 2 }],
                },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: ROS, startRowIndex: 4, endRowIndex: 18, startColumnIndex: 4, endColumnIndex: 5 }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hex(C.slateBlue),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{ sheetId: ROS, startRowIndex: 4, endRowIndex: 18, startColumnIndex: 5, endColumnIndex: 6 }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hex(C.deepBlush),
              },
            ],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: ROS, rowIndex: 19, columnIndex: 0 },
            offsetXPixels: 0, offsetYPixels: 10,
            widthPixels: 580, heightPixels: 360,
          },
        },
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
