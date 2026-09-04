'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];
const BMT = sheetMap['📏 Body Measurements'];

// Chart data table positions on Dashboard (0-based row indices):
//   Muscle Group: rows 30-40 (label col 0, count col 1)
//   Cardio Distance: rows 42-48 (label col 0, distance col 1)
//   Sleep Quality: rows 50-54 (label col 0, count col 1)
// Body Weight chart sources directly from Body Measurements tab

function srcRange(sheetId, startRow, endRow, startCol, endCol) {
  return { sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol };
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

  // ── Chart 1: BAR — Workout Count by Muscle Group ─────────────────────────
  // Data: DSH rows 30-40 (0-based), label col 0, count col 1
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '💪 Workouts by Muscle Group',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Workouts' },
              { position: 'LEFT_AXIS', title: 'Muscle Group' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 30, 40, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(DSH, 30, 40, 1, 2)] } },
              targetAxis: 'BOTTOM_AXIS',
              color: hex(C.cobaltBlue),
            }],
            stackedType: 'NOT_STACKED',
            headerCount: 0,
          },
          backgroundColor: hex(C.offWhite),
        },
        position: overlayPos(17, 0, 460, 290),
      },
    },
  });

  // ── Chart 2: DONUT — Sleep Quality Distribution ───────────────────────────
  // Data: DSH rows 50-54 (0-based), label col 0, count col 1
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '😴 Sleep Quality Distribution',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.45,
            domain: { sourceRange: { sources: [srcRange(DSH, 50, 54, 0, 1)] } },
            series: { sourceRange: { sources: [srcRange(DSH, 50, 54, 1, 2)] } },
          },
          backgroundColor: hex(C.offWhite),
        },
        position: overlayPos(17, 4, 450, 290),
      },
    },
  });

  // ── Chart 3: COLUMN — Cardio Distance by Type ────────────────────────────
  // Data: DSH rows 42-48 (0-based), label col 0, distance col 1
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '🏃 Cardio Distance by Type (mi)',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Type' },
              { position: 'LEFT_AXIS', title: 'Distance (mi)' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 42, 48, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(DSH, 42, 48, 1, 2)] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.teal),
            }],
            stackedType: 'NOT_STACKED',
            headerCount: 0,
          },
          backgroundColor: hex(C.offWhite),
        },
        position: overlayPos(36, 0, 460, 290),
      },
    },
  });

  // ── Chart 4: LINE — Body Weight Over Time ────────────────────────────────
  // Data: BMT rows 2-10 (0-based), date col 0, weight col 1
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '📏 Body Weight Over Time (lbs)',
          titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Date' },
              { position: 'LEFT_AXIS', title: 'Weight (lbs)' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(BMT, 2, 10, 0, 1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(BMT, 2, 10, 1, 2)] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(C.coral),
            }],
            headerCount: 0,
          },
          backgroundColor: hex(C.offWhite),
        },
        position: overlayPos(36, 4, 450, 290),
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
