'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🌸 Dashboard'];
const BB   = sheetMap['💰 Baby Budget'];
const ST   = sheetMap['🤒 Symptom & Wellbeing Tracker'];

(async () => {
  const reqs = [];

  // ── Chart 1: Mood Trend by Week (Line chart) — on Dashboard ───────────────
  // Data: A29:B67 (Week, Avg Mood)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Average Mood Score by Week',
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Pregnancy Week' },
          { position: 'LEFT_AXIS', title: 'Mood Score (1-10)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: DASH, startRowIndex: 28, endRowIndex: 68, startColumnIndex: 0, endColumnIndex: 1 }] } } }],
        series: [{
          series: { sourceRange: { sources: [{ sheetId: DASH, startRowIndex: 28, endRowIndex: 68, startColumnIndex: 1, endColumnIndex: 2 }] } },
          targetAxis: 'LEFT_AXIS',
          color: hex(C.deepLavender),
        }],
        headerCount: 0,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DASH, rowIndex: 27, columnIndex: 7 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 560, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 2: Budget vs Actual — Column chart on Baby Budget ───────────────
  // Data: B6:B17 (categories), D6:D17 (allocated), E6:E17 (actual)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Baby Budget: Allocated vs Actual Spend',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Category' },
          { position: 'LEFT_AXIS', title: 'Amount ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: BB, startRowIndex: 5, endRowIndex: 17, startColumnIndex: 1, endColumnIndex: 2 }] } } }],
        series: [
          {
            series: { sourceRange: { sources: [{ sheetId: BB, startRowIndex: 5, endRowIndex: 17, startColumnIndex: 3, endColumnIndex: 4 }] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.taupeGold),
          },
          {
            series: { sourceRange: { sources: [{ sheetId: BB, startRowIndex: 5, endRowIndex: 17, startColumnIndex: 4, endColumnIndex: 5 }] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.deepLavender),
          },
        ],
        headerCount: 0,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: BB, rowIndex: 38, columnIndex: 0 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 600, heightPixels: 320,
      },
    },
  }}});

  // ── Chart 3: Symptom Frequency — Column chart on Symptom Tracker ──────────
  // Use COUNTIF on the symptom column for top symptoms
  // We'll add a helper mini-table in ST rows 3-4 cols 9-11 for top symptom counts
  // Actually let's anchor to the summary row which already has the data
  // Simpler: use a stacked bar of severity in the symptom tracker
  // Chart: count by symptom type — use COUNTIF values in reference data area
  // We'll pull from symptom col C and count against reference list
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Symptom Severity Distribution',
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Count' },
          { position: 'LEFT_AXIS', title: 'Severity Level' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: ST, startRowIndex: 1, endRowIndex: 6, startColumnIndex: 4, endColumnIndex: 5 }] } } }],
        series: [{
          series: { sourceRange: { sources: [{ sheetId: ST, startRowIndex: 1, endRowIndex: 6, startColumnIndex: 11, endColumnIndex: 12 }] } },
          targetAxis: 'BOTTOM_AXIS',
          color: hex(C.softLavender),
        }],
        headerCount: 0,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: ST, rowIndex: 3, columnIndex: 9 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 480, heightPixels: 280,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
