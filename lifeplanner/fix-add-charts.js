'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['✨ Dashboard'];
const HAB = "'🔁 Habit Tracker'";

// ── Color helpers ─────────────────────────────────────────────────────────────
// Raspberry #BE3A6A, Mint #4BAF9C, Amber #C97D2B, Teal #3A7D9C, Sage #5E8C5A
const rasperryRgb = { red: 0.745, green: 0.227, blue: 0.416, alpha: 1 };
const mintRgb     = { red: 0.294, green: 0.686, blue: 0.612, alpha: 1 };
const amberRgb    = { red: 0.788, green: 0.490, blue: 0.169, alpha: 1 };
const tealRgb     = { red: 0.227, green: 0.490, blue: 0.612, alpha: 1 };
const sageRgb     = { red: 0.369, green: 0.549, blue: 0.353, alpha: 1 };

const src = (r1, r2, c1, c2) => ({
  sheetId: SID,
  startRowIndex: r1, endRowIndex: r2,
  startColumnIndex: c1, endColumnIndex: c2,
});

const overlayChart = (anchorRow, anchorCol, w, h, spec) => ({
  addChart: {
    chart: {
      spec,
      position: {
        overlayPosition: {
          anchorCell: { sheetId: SID, rowIndex: anchorRow, columnIndex: anchorCol },
          offsetXPixels: 0,
          offsetYPixels: 0,
          widthPixels: w,
          heightPixels: h,
        },
      },
    },
  },
});

(async () => {
  const vals = [];
  const reqs = [];

  // ── Wellness numeric helper cells (col E=4, F=5, rows 35-36, 0-indexed) ────
  // These feed the wellness chart with actual numbers (not TEXT-formatted strings)
  vals.push({
    range: "'✨ Dashboard'!E36:F36",
    values: [['Completion %',
      `=IFERROR(COUNTIFS(${HAB}!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1,${HAB}!A:A,"<="&TODAY(),${HAB}!C:C,TRUE)/MAX(1,COUNTIFS(${HAB}!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1,${HAB}!A:A,"<="&TODAY())),0)`
    ]],
  });
  vals.push({
    range: "'✨ Dashboard'!E37:F37",
    values: [['Habits Logged',
      `=IFERROR(COUNTIFS(${HAB}!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1,${HAB}!A:A,"<="&TODAY(),${HAB}!C:C,TRUE),0)`
    ]],
  });

  // Format F36 as percent
  reqs.push({ repeatCell: {
    range: { sheetId: SID, startRowIndex: 35, endRowIndex: 36, startColumnIndex: 5, endColumnIndex: 6 },
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' }, textFormat: { fontSize: 9 } } },
    fields: 'userEnteredFormat(numberFormat,textFormat)',
  }});

  await valuesBatchUpdate(id, vals, 'wellness-helpers');

  // ── CHART 1: Finance — Column chart (Income / Spent / Left to Budget) ───────
  // Data: A17:B19 = labels col A (idx 0), values col B (idx 1), rows 16-18 (0-idx)
  reqs.push(overlayChart(15, 4, 360, 220, {
    title: '💰 Finances This Month',
    titleTextFormat: { bold: true, fontSize: 10, foregroundColor: rasperryRgb },
    basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'NO_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [src(16, 19, 0, 1)] } } }],
      series: [{
        series: { sourceRange: { sources: [src(16, 19, 1, 2)] } },
        targetAxis: 'LEFT_AXIS',
        color: rasperryRgb,
      }],
      headerCount: 0,
    },
  }));

  // ── CHART 2: Life Balance — Horizontal bar (Current vs Goal per category) ───
  // Data: A25:C30 — labels col A (idx 0), Current col B (idx 1), Goal col C (idx 2)
  // Rows 24-29 (0-indexed)
  reqs.push(overlayChart(23, 4, 360, 280, {
    title: '⚖️ Life Balance: Current vs Goal',
    titleTextFormat: { bold: true, fontSize: 10, foregroundColor: rasperryRgb },
    basicChart: {
      chartType: 'BAR',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [src(24, 30, 0, 1)] } } }],
      series: [
        {
          series: { sourceRange: { sources: [src(24, 30, 1, 2)] } },
          targetAxis: 'BOTTOM_AXIS',
          color: mintRgb,
          dataLabel: { type: 'DATA' },
        },
        {
          series: { sourceRange: { sources: [src(24, 30, 2, 3)] } },
          targetAxis: 'BOTTOM_AXIS',
          color: rasperryRgb,
        },
      ],
      headerCount: 0,
    },
  }));

  // ── CHART 3: Task Priority Counts — Column chart ─────────────────────────────
  // Data: A32:D32 = labels (row 31, 0-idx), A33:D33 = counts (row 32, 0-idx)
  // Because this is ROW-oriented (labels across cols), we treat row 31 as domain
  // and row 32 as the series. headerCount=0 since we specify domain separately.
  reqs.push(overlayChart(30, 4, 360, 200, {
    title: '✅ Task Priority Counts',
    titleTextFormat: { bold: true, fontSize: 10, foregroundColor: rasperryRgb },
    basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'NO_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: 'Count' },
      ],
      domains: [{ domain: { sourceRange: { sources: [src(31, 32, 0, 4)] } } }],
      series: [{
        series: { sourceRange: { sources: [src(32, 33, 0, 4)] } },
        targetAxis: 'LEFT_AXIS',
        color: tealRgb,
      }],
      headerCount: 1,
    },
  }));

  // ── CHART 4: Wellness — Bar chart (Completion % + Habits Logged) ─────────────
  // Helper data at E36:F37 (rows 35-36, cols 4-5, 0-indexed)
  // Labels in col E (idx 4), values in col F (idx 5)
  reqs.push(overlayChart(34, 4, 360, 200, {
    title: '💚 Wellness This Week',
    titleTextFormat: { bold: true, fontSize: 10, foregroundColor: rasperryRgb },
    basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'NO_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [src(35, 37, 4, 5)] } } }],
      series: [{
        series: { sourceRange: { sources: [src(35, 37, 5, 6)] } },
        targetAxis: 'LEFT_AXIS',
        color: sageRgb,
      }],
      headerCount: 0,
    },
  }));

  await batchUpdate(id, reqs, 'fix-add-charts');
  console.log('✅  4 dashboard charts added.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
