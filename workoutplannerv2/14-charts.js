'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];
const STK = sheetMap['🔥 Streak & Consistency Tracker'];

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
    // Chart 1: Donut — Monthly Progress % (Completed vs Not Yet Done)
    // Data: DSH rows 22-23 (idx 22,23) cols A-B → labels + counts
    {
      addChart: {
        chart: {
          spec: {
            title: '📈 Monthly Progress',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.navyDark) },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND', pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(DSH, 22, 25, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(DSH, 22, 25, 1, 2)] } },
            },
            backgroundColor: hex(C.offWhite),
          },
          position: overlayPos(DSH, 14, 0, 400, 300),
        },
      },
    },
    // Chart 2: Donut — Muscle Group Split
    // Data: DSH rows 26-35 (idx 26,35) cols A-B
    {
      addChart: {
        chart: {
          spec: {
            title: '💪 Muscle Group Split',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.navyDark) },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND', pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(DSH, 26, 36, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(DSH, 26, 36, 1, 2)] } },
            },
            backgroundColor: hex(C.offWhite),
          },
          position: overlayPos(DSH, 14, 4, 450, 300),
        },
      },
    },
    // Chart 3: Column — Daily Completion (Day 1-30)
    // Data: DSH rows 37-67 (idx 37,67) cols A-B
    {
      addChart: {
        chart: {
          spec: {
            title: '📅 Daily Progress — Selected Month',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.navyDark) },
            basicChart: {
              chartType: 'COLUMN',
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: 'Day' },
                { position: 'LEFT_AXIS', title: 'Completed' },
              ],
              domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 37, 68, 0, 1)] } } }],
              series: [{
                series: { sourceRange: { sources: [srcRange(DSH, 37, 68, 1, 2)] } },
                targetAxis: 'LEFT_AXIS',
                color: hex(C.limeGreen),
              }],
              stackedType: 'NOT_STACKED',
              headerCount: 1,
            },
            backgroundColor: hex(C.offWhite),
          },
          position: overlayPos(DSH, 22, 0, 500, 290),
        },
      },
    },
    // Chart 4: Line — Running Streak over time
    // Data: STK rows 1-30 (idx 1,31) cols A (date) + C (streak)
    {
      addChart: {
        chart: {
          spec: {
            title: '🔥 Running Streak over Time',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.navyDark) },
            basicChart: {
              chartType: 'LINE',
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: 'Date' },
                { position: 'LEFT_AXIS', title: 'Streak (days)' },
              ],
              domains: [{ domain: { sourceRange: { sources: [srcRange(STK, 1, 31, 0, 1)] } } }],
              series: [{
                series: { sourceRange: { sources: [srcRange(STK, 1, 31, 2, 3)] } },
                targetAxis: 'LEFT_AXIS',
                color: hex('#E65100'),
              }],
              headerCount: 1,
            },
            backgroundColor: hex(C.offWhite),
          },
          position: overlayPos(DSH, 22, 5, 450, 290),
        },
      },
    },
  ];

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
