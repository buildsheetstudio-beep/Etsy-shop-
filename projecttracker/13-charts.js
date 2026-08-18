'use strict';
const { sheets, hex, batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB_SID  = sheetMap['Project Dashboard'];
const MTL_SID = sheetMap['Master Task Log'];

// Helper to build an embedded chart request
const addChart = (sheetId, anchorRow, anchorCol, spec) => ({
  addChart: {
    chart: {
      spec,
      position: {
        overlayPosition: {
          anchorCell: { sheetId, rowIndex: anchorRow, columnIndex: anchorCol },
          offsetXPixels: 0, offsetYPixels: 0,
          widthPixels: 380, heightPixels: 260,
        },
      },
    },
  },
});

(async () => {
  const requests = [];

  // Chart 1: Task Status Donut — data in Dashboard rows 32-38 cols A-B (0-indexed 31-37)
  // Status (col A=0) and Count (col B=1)
  requests.push(addChart(DB_SID, 3, 0, {
    title: 'Task Status Distribution',
    pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.4,
      domain: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 31, endRowIndex: 38, startColumnIndex: 0, endColumnIndex: 1 }] } },
      series: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 31, endRowIndex: 38, startColumnIndex: 1, endColumnIndex: 2 }] } },
    },
  }));

  // Chart 2: Workload by Assignee — data in Dashboard rows 46-52 cols A-D (0-indexed 45-52)
  requests.push(addChart(DB_SID, 43, 0, {
    title: 'Workload by Assignee',
    basicChart: {
      chartType: 'BAR',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Task Count' },
        { position: 'LEFT_AXIS',   title: 'Assignee' },
      ],
      domains: [{
        domain: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 45, endRowIndex: 52, startColumnIndex: 0, endColumnIndex: 1 }] } },
      }],
      series: [
        {
          series: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 45, endRowIndex: 52, startColumnIndex: 1, endColumnIndex: 2 }] } },
          targetAxis: 'BOTTOM_AXIS',
        },
        {
          series: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 45, endRowIndex: 52, startColumnIndex: 2, endColumnIndex: 3 }] } },
          targetAxis: 'BOTTOM_AXIS',
        },
      ],
      headerCount: 0,
    },
  }));

  // Chart 3: Monthly Completion (Due vs Completed) — data rows 46-57 cols E-G (0-indexed 45-56, cols 4-6)
  requests.push(addChart(DB_SID, 43, 4, {
    title: 'Monthly Tasks: Due vs Completed (YTD)',
    basicChart: {
      chartType: 'COMBO',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Month' },
        { position: 'LEFT_AXIS',   title: 'Task Count' },
      ],
      domains: [{
        domain: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 45, endRowIndex: 57, startColumnIndex: 4, endColumnIndex: 5 }] } },
      }],
      series: [
        {
          series: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 45, endRowIndex: 57, startColumnIndex: 5, endColumnIndex: 6 }] } },
          targetAxis: 'LEFT_AXIS',
          type: 'COLUMN',
        },
        {
          series: { sourceRange: { sources: [{ sheetId: DB_SID, startRowIndex: 45, endRowIndex: 57, startColumnIndex: 6, endColumnIndex: 7 }] } },
          targetAxis: 'LEFT_AXIS',
          type: 'COLUMN',
        },
      ],
      headerCount: 0,
    },
  }));

  await batchUpdate(id, requests, 'charts');
  console.log('✓ Charts complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
