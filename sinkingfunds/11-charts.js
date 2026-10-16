'use strict';
const { sheets, hex, batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const reqs = [];

  const SETUP_SID = sheetMap['Fund Setup & Goals'];
  const LOG_SID   = sheetMap['Contribution Log'];
  const FORE_SID  = sheetMap['Savings Forecast'];
  const ANN_SID   = sheetMap['Annual Summary'];
  const DASH_SID  = sheetMap['Sinking Funds Dashboard'];

  // ── Chart 1: Fund Setup — Goal vs Balance bar chart ──────────────────────────
  // Anchor at row 30, col 0 (below fund data)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Current Balance vs Goal — All Funds',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Amount ($)' },
              { position: 'LEFT_AXIS', title: 'Fund' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [{ sheetId: SETUP_SID, startRowIndex: 7, endRowIndex: 33, startColumnIndex: 1, endColumnIndex: 2 }] },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: { sources: [{ sheetId: SETUP_SID, startRowIndex: 7, endRowIndex: 33, startColumnIndex: 8, endColumnIndex: 9 }] },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: { red: 0.353, green: 0.533, blue: 0.502 },
              },
              {
                series: {
                  sourceRange: { sources: [{ sheetId: SETUP_SID, startRowIndex: 7, endRowIndex: 33, startColumnIndex: 6, endColumnIndex: 7 }] },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: { red: 0.29, green: 0.431, blue: 0.541 },
              },
            ],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: SETUP_SID, rowIndex: 29, columnIndex: 0 },
            widthPixels: 700,
            heightPixels: 400,
          },
        },
      },
    },
  });

  // ── Chart 2: Annual Summary — Monthly Contributions bar chart ─────────────────
  // Month names are in col A, rows 14-25 (0-indexed 13-24); Contributions in col B
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Monthly Contributions — Selected Year',
          basicChart: {
            chartType: 'COMBO',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Amount ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [{ sheetId: ANN_SID, startRowIndex: 13, endRowIndex: 26, startColumnIndex: 0, endColumnIndex: 1 }] },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: { sources: [{ sheetId: ANN_SID, startRowIndex: 13, endRowIndex: 26, startColumnIndex: 1, endColumnIndex: 2 }] },
                },
                targetAxis: 'LEFT_AXIS',
                type: 'COLUMN',
                color: { red: 0.353, green: 0.533, blue: 0.502 },
              },
              {
                series: {
                  sourceRange: { sources: [{ sheetId: ANN_SID, startRowIndex: 13, endRowIndex: 26, startColumnIndex: 3, endColumnIndex: 4 }] },
                },
                targetAxis: 'LEFT_AXIS',
                type: 'COLUMN',
                color: { red: 0.29, green: 0.431, blue: 0.541 },
              },
              {
                series: {
                  sourceRange: { sources: [{ sheetId: ANN_SID, startRowIndex: 13, endRowIndex: 26, startColumnIndex: 7, endColumnIndex: 8 }] },
                },
                targetAxis: 'LEFT_AXIS',
                type: 'LINE',
                color: { red: 0.651, green: 0.365, blue: 0.431 },
              },
            ],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: ANN_SID, rowIndex: 4, columnIndex: 5 },
            widthPixels: 600,
            heightPixels: 300,
          },
        },
      },
    },
  });

  // ── Chart 3: Savings Forecast — % Funded bar chart ────────────────────────────
  // Fund names in col B (rows 10-35, 0-indexed 9-34), % Funded in col P (index 15)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '% Funded — All Sinking Funds',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: '% Funded' },
              { position: 'LEFT_AXIS', title: 'Fund' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [{ sheetId: FORE_SID, startRowIndex: 9, endRowIndex: 35, startColumnIndex: 1, endColumnIndex: 2 }] },
              },
            }],
            series: [
              {
                series: {
                  sourceRange: { sources: [{ sheetId: FORE_SID, startRowIndex: 9, endRowIndex: 35, startColumnIndex: 15, endColumnIndex: 16 }] },
                },
                targetAxis: 'BOTTOM_AXIS',
                color: { red: 0.353, green: 0.533, blue: 0.502 },
              },
            ],
            headerCount: 0,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: FORE_SID, rowIndex: 36, columnIndex: 0 },
            widthPixels: 700,
            heightPixels: 450,
          },
        },
      },
    },
  });

  // ── Chart 4: Dashboard — Fund Status Distribution (pie chart) ─────────────────
  // We'll use the status counts from Dashboard KPI area
  // Status values hardcoded via a helper table in rows 42-50 of dashboard
  // Simpler: use a named range approach — build mini table then chart it
  // Use rows 42-48 of Dashboard as status tally (col A=label, col B=count)
  const STATUS_LABELS = ['Active','Ahead of Plan','On Track','Behind Plan','Goal Reached','Paused','Not Started'];
  const SETUP_V_COL = 21; // col V = index 21
  // Write a hidden tally table at row 42 of dashboard
  const tallyVals = [];
  const tallyFmt  = [];
  STATUS_LABELS.forEach((label, i) => {
    const r = 42 + i;
    tallyVals.push({ range: `'Sinking Funds Dashboard'!T${r}`, values: [[label]] });
    tallyVals.push({ range: `'Sinking Funds Dashboard'!U${r}`, values: [[`=IFERROR(COUNTIF('Fund Setup & Goals'!$V$8:$V$33,"${label}"),0)`]] });
  });
  // We'll push these via the main batchUpdate below — need a values call
  // Actually, let me add a separate chart that uses inline data from Fund Setup directly

  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Fund Status Distribution',
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            threeDimensional: false,
            domain: {
              sourceRange: { sources: [{ sheetId: DASH_SID, startRowIndex: 41, endRowIndex: 48, startColumnIndex: 19, endColumnIndex: 20 }] },
            },
            series: {
              sourceRange: { sources: [{ sheetId: DASH_SID, startRowIndex: 41, endRowIndex: 48, startColumnIndex: 20, endColumnIndex: 21 }] },
            },
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DASH_SID, rowIndex: 41, columnIndex: 0 },
            widthPixels: 450,
            heightPixels: 300,
          },
        },
      },
    },
  });

  // Store the tally values so Chart 4 has data to reference
  const { valuesBatchUpdate } = require('./lib');
  await valuesBatchUpdate(id, [
    ...STATUS_LABELS.map((label, i) => ({ range: `'Sinking Funds Dashboard'!T${42+i}`, values: [[label]] })),
    ...STATUS_LABELS.map((label, i) => ({ range: `'Sinking Funds Dashboard'!U${42+i}`, values: [[`=IFERROR(COUNTIF('Fund Setup & Goals'!$V$8:$V$33,"${label}"),0)`]] })),
  ], 'chart-tally');

  await batchUpdate(id, reqs, 'charts');

  console.log('✓ Charts complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
