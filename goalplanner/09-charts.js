'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['Life Dashboard'];      // 0
const GOALS = sheetMap['Annual Goals'];       // 1
const QPLAN = sheetMap['Quarterly Planner'];  // 2
const MONTHLY = sheetMap['Monthly Milestones']; // 3
const HABITS = sheetMap['Habit Tracker'];     // 4

function anchor(sheetId, row, col) {
  return { sheetId, rowIndex: row, columnIndex: col };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Wheel of Life — bar chart on Life Dashboard ─────────────────
  // Data: Dashboard H9:I14 (area name, score)
  // Anchor: Dashboard H8 (row 7, col 7)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Wheel of Life — 2025 Scores',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Score (1–10)' },
              { position: 'LEFT_AXIS', title: '' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: DASH, startRowIndex: 8, endRowIndex: 14, startColumnIndex: 7, endColumnIndex: 8 }],
                },
              },
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{ sheetId: DASH, startRowIndex: 8, endRowIndex: 14, startColumnIndex: 8, endColumnIndex: 9 }],
                },
              },
              targetAxis: 'BOTTOM_AXIS',
              color: hex(C.dustyLavender),
            }],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(DASH, 7, 7),
            offsetXPixels: 0,
            offsetYPixels: 0,
            widthPixels: 420,
            heightPixels: 260,
          },
        },
      },
    },
  });

  // ── Chart 2: Goal Status — pie chart on Annual Goals ─────────────────────
  // Data: Annual Goals N5:O9 (status, count)
  // Anchor: Annual Goals N4 (row 3, col 13)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '2025 Goal Status',
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            domain: {
              sourceRange: {
                sources: [{ sheetId: GOALS, startRowIndex: 4, endRowIndex: 9, startColumnIndex: 13, endColumnIndex: 14 }],
              },
            },
            series: {
              sourceRange: {
                sources: [{ sheetId: GOALS, startRowIndex: 4, endRowIndex: 9, startColumnIndex: 14, endColumnIndex: 15 }],
              },
            },
            threeDimensional: false,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(GOALS, 3, 13),
            offsetXPixels: 0,
            offsetYPixels: 0,
            widthPixels: 380,
            heightPixels: 260,
          },
        },
      },
    },
  });

  // ── Chart 3: Q2 Weekly Completion — line chart on Quarterly Planner ──────
  // Data: Quarterly M12:N25 (week, completion %)
  // Anchor: Quarterly M11 (row 10, col 12)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Q2 Weekly Completion %',
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Week' },
              { position: 'LEFT_AXIS', title: 'Completion %' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: QPLAN, startRowIndex: 12, endRowIndex: 25, startColumnIndex: 12, endColumnIndex: 13 }],
                },
              },
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{ sheetId: QPLAN, startRowIndex: 12, endRowIndex: 25, startColumnIndex: 13, endColumnIndex: 14 }],
                },
              },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.skyBlue),
            }],
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(QPLAN, 10, 12),
            offsetXPixels: 0,
            offsetYPixels: 0,
            widthPixels: 400,
            heightPixels: 260,
          },
        },
      },
    },
  });

  // ── Chart 4: Monthly June Goal Progress — bar on Monthly Milestones ───────
  // Data: Monthly A43:B48 (goal, progress)
  // Anchor: Monthly E41 (row 40, col 4)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'June Goal Progress',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: '% Complete' },
              { position: 'LEFT_AXIS', title: '' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: MONTHLY, startRowIndex: 42, endRowIndex: 48, startColumnIndex: 0, endColumnIndex: 1 }],
                },
              },
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{ sheetId: MONTHLY, startRowIndex: 42, endRowIndex: 48, startColumnIndex: 1, endColumnIndex: 2 }],
                },
              },
              targetAxis: 'BOTTOM_AXIS',
              color: hex(C.blush),
            }],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(MONTHLY, 40, 4),
            offsetXPixels: 0,
            offsetYPixels: 0,
            widthPixels: 400,
            heightPixels: 260,
          },
        },
      },
    },
  });

  // ── Chart 5: Habit Completion % by Day — column on Habit Tracker ─────────
  // Data: Habit Tracker I17:W17 (days 1-15 completion %)
  // Anchor: Habit Tracker A19 (row 18, col 0)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Daily Habit Completion — June 1–15',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Day' },
              { position: 'LEFT_AXIS', title: '% of Habits Done' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: HABITS, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 8, endColumnIndex: 23 }],
                },
              },
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{ sheetId: HABITS, startRowIndex: 16, endRowIndex: 17, startColumnIndex: 8, endColumnIndex: 23 }],
                },
              },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.dustyLavender),
            }],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(HABITS, 18, 0),
            offsetXPixels: 0,
            offsetYPixels: 0,
            widthPixels: 500,
            heightPixels: 240,
          },
        },
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');

  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
