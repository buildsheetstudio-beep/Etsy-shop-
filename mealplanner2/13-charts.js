'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_MCT = sheetMap['Meal Cost Tracker'];
const SID_GL  = sheetMap['Automated Grocery List'];

// Helpers for chart source range
const src = (sid, r1, r2, c1, c2) => ({
  sheetId: sid, startRowIndex: r1, endRowIndex: r2,
  startColumnIndex: c1, endColumnIndex: c2,
});

// overlayPosition anchor
const anchor = (sid, row, col, w, h) => ({
  overlayPosition: {
    anchorCell: { sheetId: sid, rowIndex: row, columnIndex: col },
    offsetXPixels: 0, offsetYPixels: 0,
    widthPixels: w, heightPixels: h,
  },
});

(async () => {
  const fmtReqs = [];   // batchUpdate requests (addChart)
  const valData = [];   // valuesBatchUpdate for helper tables

  // ── MEAL COST TRACKER ──────────────────────────────────────────────────────
  //
  // Weekly cost table (0-indexed):
  //   Row 11: header  (Day | Breakfast Cost | Lunch Cost | Dinner Cost | Snacks/Dessert | Total | Notes)
  //   Rows 12-18: Mon-Sun
  //   Row 19: Weekly Total (skipped in chart range)
  //
  // Recipe cost table:
  //   Row 21: separator
  //   Row 22: header  (Recipe ID | Recipe Name | Base Servings | Base Cost | Cost/Serving | Times | Servings | Weekly Cost | Avg | %)
  //   Rows 23-40: 18 planned recipes

  // Chart 1: Stacked COLUMN — daily meal cost by type (Breakfast / Lunch / Dinner / Snacks)
  fmtReqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Daily Meal Cost Breakdown',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            stackedType: 'STACKED',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Day of Week' },
              { position: 'LEFT_AXIS',   title: 'Estimated Cost ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [src(SID_MCT, 11, 19, 0, 1)] },
              },
            }],
            series: [
              { series: { sourceRange: { sources: [src(SID_MCT, 11, 19, 1, 2)] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN' },
              { series: { sourceRange: { sources: [src(SID_MCT, 11, 19, 2, 3)] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN' },
              { series: { sourceRange: { sources: [src(SID_MCT, 11, 19, 3, 4)] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN' },
              { series: { sourceRange: { sources: [src(SID_MCT, 11, 19, 4, 5)] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN' },
            ],
            headerCount: 1,
          },
        },
        position: anchor(SID_MCT, 11, 9, 480, 290),
      },
    },
  });

  // Chart 2: COLUMN — total daily planned cost
  fmtReqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Total Planned Meal Cost per Day',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Day of Week' },
              { position: 'LEFT_AXIS',   title: 'Total Cost ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [src(SID_MCT, 11, 19, 0, 1)] },
              },
            }],
            series: [{
              series: { sourceRange: { sources: [src(SID_MCT, 11, 19, 5, 6)] } },
              targetAxis: 'LEFT_AXIS',
              type: 'COLUMN',
            }],
            headerCount: 1,
          },
        },
        position: anchor(SID_MCT, 21, 9, 480, 270),
      },
    },
  });

  // Chart 3: BAR — recipe weekly planned cost (recipe name vs cost)
  // Recipe table: header at 0-indexed row 22, data at rows 23-40 (18 recipes)
  // Col B (index 1) = Recipe Name, Col H (index 7) = Weekly Planned Cost
  fmtReqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Weekly Planned Cost by Recipe',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'LEFT_AXIS',   title: 'Recipe' },
              { position: 'BOTTOM_AXIS', title: 'Weekly Cost ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [src(SID_MCT, 22, 41, 1, 2)] },
              },
            }],
            series: [{
              series: { sourceRange: { sources: [src(SID_MCT, 22, 41, 7, 8)] } },
              targetAxis: 'BOTTOM_AXIS',
              type: 'BAR',
            }],
            headerCount: 1,
          },
        },
        position: anchor(SID_MCT, 40, 0, 550, 420),
      },
    },
  });

  // ── AUTOMATED GROCERY LIST ─────────────────────────────────────────────────
  //
  // Grocery data: rows 12-70 (1-indexed), col L (index 11) = Status, col J (index 9) = Est. Grocery Cost
  // We add a small summary table at col R (index 17) rows 2-7 (0-indexed 1-6) for charting.

  // Helper table: status summary at col R-S, rows 2-6 (0-indexed 1-5)
  // R2 (0-idx 1,17): "Status", S2: "Count", T2: "Est. Cost"
  // Labels and COUNTIF/SUMIF formulas referencing grocery data col L and col J
  valData.push({
    range: "'Automated Grocery List'!R2:T6",
    values: [
      ['Status',       'Count',                                           'Est. Cost'],
      ['Needed',       '=COUNTIF($L$12:$L$70,"Needed")',                 '=SUMIF($L$12:$L$70,"Needed",$J$12:$J$70)'],
      ['Already Have', '=COUNTIF($L$12:$L$70,"Already Have")',           '=SUMIF($L$12:$L$70,"Already Have",$J$12:$J$70)'],
      ['Optional',     '=COUNTIF($L$12:$L$70,"Optional")',               '=SUMIF($L$12:$L$70,"Optional",$J$12:$J$70)'],
      ['Review Units', '=COUNTIF($L$12:$L$70,"Review Units")',           '=SUMIF($L$12:$L$70,"Review Units",$J$12:$J$70)'],
    ],
  });

  // Light header style for the helper table (very small, unobtrusive)
  fmtReqs.push({
    repeatCell: {
      range: gridRange(SID_GL, 1, 2, 17, 20),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.secText) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  fmtReqs.push({
    repeatCell: {
      range: gridRange(SID_GL, 2, 6, 17, 20),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 8, fontFamily: 'Arial' } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Chart 4: PIE — grocery status distribution (count)
  // summary table: 0-indexed rows 1-5, cols 17-18 (R-S), header at row 1
  fmtReqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Grocery Item Status Distribution',
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            threeDimensional: false,
            domain: {
              sourceRange: { sources: [src(SID_GL, 1, 6, 17, 18)] },
            },
            series: {
              sourceRange: { sources: [src(SID_GL, 1, 6, 18, 19)] },
            },
          },
        },
        position: anchor(SID_GL, 1, 20, 380, 280),
      },
    },
  });

  // Chart 5: COLUMN — grocery estimated cost by status
  fmtReqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Estimated Grocery Cost by Status',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Status' },
              { position: 'LEFT_AXIS',   title: 'Estimated Cost ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [src(SID_GL, 1, 6, 17, 18)] },
              },
            }],
            series: [{
              series: { sourceRange: { sources: [src(SID_GL, 1, 6, 19, 20)] } },
              targetAxis: 'LEFT_AXIS',
              type: 'COLUMN',
            }],
            headerCount: 1,
          },
        },
        position: anchor(SID_GL, 11, 20, 380, 280),
      },
    },
  });

  // Write helper table values first, then push chart requests
  await valuesBatchUpdate(id, valData, 'chart-data');
  await batchUpdate(id, fmtReqs, 'charts');
  console.log('✓ Charts complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
