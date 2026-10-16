'use strict';
const { sheets, hex, batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH_SID = sheetMap['Home Maintenance Dashboard'];
const TASK_SID = sheetMap['Maintenance Task Log'];
const COST_SID = sheetMap['Maintenance Costs'];
const REP_SID  = sheetMap['Home Repair Log'];

// Helper: color object from hex string
function hx(colorStr) {
  const c = colorStr.replace('#','');
  return { red: parseInt(c.slice(0,2),16)/255, green: parseInt(c.slice(2,4),16)/255, blue: parseInt(c.slice(4,6),16)/255 };
}

// Column index helper (0-indexed)
function ci(letter) { return letter.charCodeAt(0) - 65; }

(async () => {
  const fmt = [];

  // ────────────────────────────────────────────────────────────────────────────
  // Chart 1: Task Status Donut — embedded in Dashboard rows 5-11 (right side)
  // Source: inline series from KPI values (Complete, In Progress, Scheduled, Not Started, Overdue)
  // ────────────────────────────────────────────────────────────────────────────
  // We'll use the Dashboard's KPI formula cells as chart source
  // KPI row 7 (0-indexed row 6): cols A,C,E,G,I,K hold the values
  // Labels row: we'll write a helper label range off to the right (col N onwards)
  // Actually, use the Task Log directly — count by status for the donut

  // Chart 1: Status donut — pie chart on Dashboard sheet
  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: 'Task Status Breakdown',
          titleTextFormat: { bold: true, fontSize: 10, foregroundColor: hx(C.text), fontFamily: 'Arial' },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.45,
            domain: {
              sourceRange: {
                sources: [{
                  sheetId: DASH_SID,
                  startRowIndex: 43, endRowIndex: 56,
                  startColumnIndex: 0, endColumnIndex: 1,
                }],
              },
            },
            series: {
              sourceRange: {
                sources: [{
                  sheetId: DASH_SID,
                  startRowIndex: 43, endRowIndex: 56,
                  startColumnIndex: 1, endColumnIndex: 2,
                }],
              },
            },
          },
          backgroundColor: hx(C.panel),
          fontName: 'Arial',
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DASH_SID, rowIndex: 4, columnIndex: 7 },
            offsetXPixels: 5,
            offsetYPixels: 5,
            widthPixels: 340,
            heightPixels: 180,
          },
        },
      },
    },
  });

  // ────────────────────────────────────────────────────────────────────────────
  // Chart 2: Annual cost trend bar chart — embedded in Dashboard row 31-40
  // Source: annual cost trend table (rows 31-35, cols A-D = Year, Total, DIY, Pro)
  // ────────────────────────────────────────────────────────────────────────────
  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: 'Annual Maintenance Cost',
          titleTextFormat: { bold: true, fontSize: 10, foregroundColor: hx(C.text), fontFamily: 'Arial' },
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'RIGHT_LEGEND',
            axis: [
              { position: 'LEFT_AXIS', title: 'Year' },
              { position: 'BOTTOM_AXIS', title: 'Cost ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{
                    sheetId: DASH_SID,
                    startRowIndex: 31, endRowIndex: 36,
                    startColumnIndex: 0, endColumnIndex: 1,
                  }],
                },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: {
                    sources: [{
                      sheetId: DASH_SID,
                      startRowIndex: 31, endRowIndex: 36,
                      startColumnIndex: 2, endColumnIndex: 3,
                    }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hx(C.success),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{
                      sheetId: DASH_SID,
                      startRowIndex: 31, endRowIndex: 36,
                      startColumnIndex: 3, endColumnIndex: 4,
                    }],
                  },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: hx(C.secondary),
              },
            ],
            headerCount: 0,
            stackedType: 'STACKED',
          },
          backgroundColor: hx(C.panel),
          fontName: 'Arial',
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DASH_SID, rowIndex: 29, columnIndex: 0 },
            offsetXPixels: 5,
            offsetYPixels: 5,
            widthPixels: 480,
            heightPixels: 200,
          },
        },
      },
    },
  });

  // ────────────────────────────────────────────────────────────────────────────
  // Chart 3: Monthly task completion COMBO — embedded in Dashboard row 42+
  // Source: Monthly table rows 44-55, cols A (month), B (tasks done), C (cost)
  // Tasks done = COLUMN series, Actual Cost = LINE series
  // ────────────────────────────────────────────────────────────────────────────
  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: 'Monthly Tasks Completed & Cost',
          titleTextFormat: { bold: true, fontSize: 10, foregroundColor: hx(C.text), fontFamily: 'Arial' },
          basicChart: {
            chartType: 'COMBO',
            legendPosition: 'RIGHT_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Tasks Done' },
              { position: 'RIGHT_AXIS', title: 'Cost ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{
                    sheetId: DASH_SID,
                    startRowIndex: 44, endRowIndex: 56,
                    startColumnIndex: 0, endColumnIndex: 1,
                  }],
                },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: {
                    sources: [{
                      sheetId: DASH_SID,
                      startRowIndex: 44, endRowIndex: 56,
                      startColumnIndex: 1, endColumnIndex: 2,
                    }],
                  },
                },
                targetAxis: 'LEFT_AXIS',
                type: 'COLUMN',
                color: hx(C.primary),
              },
              {
                series: {
                  sourceRange: {
                    sources: [{
                      sheetId: DASH_SID,
                      startRowIndex: 44, endRowIndex: 56,
                      startColumnIndex: 2, endColumnIndex: 3,
                    }],
                  },
                },
                targetAxis: 'RIGHT_AXIS',
                type: 'LINE',
                color: hx(C.secondary),
              },
            ],
            headerCount: 0,
          },
          backgroundColor: hx(C.panel),
          fontName: 'Arial',
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DASH_SID, rowIndex: 42, columnIndex: 0 },
            offsetXPixels: 5,
            offsetYPixels: 5,
            widthPixels: 500,
            heightPixels: 200,
          },
        },
      },
    },
  });

  // ────────────────────────────────────────────────────────────────────────────
  // Chart 4: Appliance/System condition summary pie — embedded in Appliances tab
  // We embed directly in the Appliances & Systems sheet
  // Source: count by condition using a helper table
  // Use Dashboard cost-by-category section as a simpler alternative:
  // Actually embed directly as a donut of condition counts
  // ────────────────────────────────────────────────────────────────────────────
  const AST_SID = sheetMap['Appliances & Systems'];

  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: 'Asset Conditions',
          titleTextFormat: { bold: true, fontSize: 10, foregroundColor: hx(C.text), fontFamily: 'Arial' },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.4,
            domain: {
              sourceRange: {
                sources: [{
                  sheetId: AST_SID,
                  startRowIndex: 7, endRowIndex: 42,
                  startColumnIndex: 15, endColumnIndex: 16,
                }],
              },
            },
            series: {
              sourceRange: {
                sources: [{
                  sheetId: AST_SID,
                  startRowIndex: 7, endRowIndex: 42,
                  startColumnIndex: 15, endColumnIndex: 16,
                }],
              },
            },
          },
          backgroundColor: hx(C.panel),
          fontName: 'Arial',
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: AST_SID, rowIndex: 2, columnIndex: 12 },
            offsetXPixels: 5,
            offsetYPixels: 5,
            widthPixels: 320,
            heightPixels: 200,
          },
        },
      },
    },
  });

  await batchUpdate(id, fmt, 'charts');

  console.log('✓ Charts complete — 4 charts added');
})().catch(e => { console.error(e.message || e); process.exit(1); });
