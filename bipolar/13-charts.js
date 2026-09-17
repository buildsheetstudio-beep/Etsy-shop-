'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard — Annual Overview'];

function sourceRange(r1, r2, c1, c2) {
  return { sources: [{ sheetId: DSH, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }] };
}

function colorObj(hexStr) { return hex(hexStr); }

(async () => {
  const reqs = [];

  // Chart 1: Overall Severity Trend — LINE — monthly severity by episode type
  // Data rows: header row 17 (idx 16), data rows 18-29 (idx 17-29)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Monthly Severity Trend',
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Avg Severity' },
            ],
            domains: [{ domain: { sourceRange: sourceRange(16, 29, 0, 1) }, reversed: false }],
            series: [
              {
                series: { sourceRange: sourceRange(16, 29, 1, 2) },
                targetAxis: 'LEFT_AXIS',
                color: colorObj(C.periwinkle),
                lineStyle: { width: 2 },
              },
              {
                series: { sourceRange: sourceRange(16, 29, 2, 3) },
                targetAxis: 'LEFT_AXIS',
                color: colorObj(C.taupe),
                lineStyle: { width: 2 },
              },
              {
                series: { sourceRange: sourceRange(16, 29, 3, 4) },
                targetAxis: 'LEFT_AXIS',
                color: colorObj(C.heatMid),
                lineStyle: { width: 2 },
              },
            ],
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DSH, rowIndex: 47, columnIndex: 0 },
            offsetXPixels: 5,
            offsetYPixels: 5,
            widthPixels: 420,
            heightPixels: 250,
          },
        },
      },
    },
  });

  // Chart 2: Sleep vs Severity — LINE, dual axis
  // Severity: DSH B17:B29 (left axis), Sleep Hours: DSH E17:E29 (right axis)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Sleep Hours vs. Severity',
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Avg Severity' },
              { position: 'RIGHT_AXIS', title: 'Avg Sleep Hrs' },
            ],
            domains: [{ domain: { sourceRange: sourceRange(16, 29, 0, 1) } }],
            series: [
              {
                series: { sourceRange: sourceRange(16, 29, 1, 2) },
                targetAxis: 'LEFT_AXIS',
                color: colorObj(C.periwinkle),
                lineStyle: { width: 2 },
              },
              {
                series: { sourceRange: sourceRange(16, 29, 4, 5) },
                targetAxis: 'RIGHT_AXIS',
                color: colorObj(C.sage),
                lineStyle: { width: 2 },
              },
            ],
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DSH, rowIndex: 47, columnIndex: 4 },
            offsetXPixels: 5,
            offsetYPixels: 5,
            widthPixels: 420,
            heightPixels: 250,
          },
        },
      },
    },
  });

  // Chart 3: Medication Adherence — COLUMN
  // Data rows: header row 33 (idx 32), data rows 34-45 (idx 33-45)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Monthly Medication Adherence',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Adherence %' },
            ],
            domains: [{ domain: { sourceRange: sourceRange(32, 45, 0, 1) } }],
            series: [
              {
                series: { sourceRange: sourceRange(32, 45, 1, 2) },
                targetAxis: 'LEFT_AXIS',
                color: colorObj(C.sage),
              },
            ],
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DSH, rowIndex: 66, columnIndex: 0 },
            offsetXPixels: 5,
            offsetYPixels: 5,
            widthPixels: 420,
            heightPixels: 250,
          },
        },
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
