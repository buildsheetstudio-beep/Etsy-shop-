'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_WEIGHT  = sheetMap['Weight Progress'];
const SID_WORKOUT = sheetMap['Workout Tracker'];
const SID_TRENDS  = sheetMap['Progress & Trends'];
const SID_DASH    = sheetMap['Weight Dashboard'];

function srcRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const req = [];

  // ── Chart 1: Weight Trend Line (on Weight Progress tab)
  // B = date (col 1), C = weight lbs (col 2), E = 7-day avg (col 4)
  req.push({ addChart: { chart: {
    spec: {
      title: 'Weight Trend',
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Date' },
          { position: 'LEFT_AXIS',   title: 'Weight (lbs)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(SID_WEIGHT, 7, 107, 1, 2)] } } }],
        series: [
          {
            series: { sourceRange: { sources: [srcRange(SID_WEIGHT, 7, 107, 2, 3)] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.secondary),
          },
          {
            series: { sourceRange: { sources: [srcRange(SID_WEIGHT, 7, 107, 4, 5)] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.primary),
          },
        ],
        headerCount: 0,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: SID_WEIGHT, rowIndex: 7, columnIndex: 9 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 560, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 2: Workout Type Breakdown (bar chart on Workout Tracker tab)
  // Uses Progress & Trends section 2: rows 22-29, col A = type, col B = count
  req.push({ addChart: { chart: {
    spec: {
      title: 'Workouts by Type',
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Sessions' },
          { position: 'LEFT_AXIS',   title: 'Type' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(SID_TRENDS, 21, 29, 0, 1)] } } }],
        series: [
          {
            series: { sourceRange: { sources: [srcRange(SID_TRENDS, 21, 29, 1, 2)] } },
            targetAxis: 'BOTTOM_AXIS',
            color: hex(C.secondary),
          },
        ],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: SID_WORKOUT, rowIndex: 7, columnIndex: 9 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 400, heightPixels: 280,
      },
    },
  }}});

  // ── Chart 3: Monthly Weight Average (column chart on Progress & Trends tab)
  // Section 1 rows 7-18 (ri 6-18): col A = month name, col B = avg weight
  req.push({ addChart: { chart: {
    spec: {
      title: 'Monthly Average Weight',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Month' },
          { position: 'LEFT_AXIS',   title: 'Avg Weight (lbs)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(SID_TRENDS, 6, 18, 0, 1)] } } }],
        series: [
          {
            series: { sourceRange: { sources: [srcRange(SID_TRENDS, 6, 18, 1, 2)] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.primary),
          },
        ],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: SID_TRENDS, rowIndex: 6, columnIndex: 4 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 500, heightPixels: 280,
      },
    },
  }}});

  await batchUpdate(id, req, '18-charts');
  console.log('✅  Charts done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
