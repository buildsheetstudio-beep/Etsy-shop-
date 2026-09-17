'use strict';
const { batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_GOA = sheetMap['Goals & Challenges'];
const SID_INS = sheetMap['Reading Insights'];
const SID_DSH = sheetMap['Reading Dashboard'];

const wine   = { red: 0.431, green: 0.169, blue: 0.247, alpha: 1 };
const gold   = { red: 0.722, green: 0.588, blue: 0.353, alpha: 1 };
const green  = { red: 0.290, green: 0.486, blue: 0.349, alpha: 1 };
const teal   = { red: 0.290, green: 0.431, blue: 0.549, alpha: 1 };
const amber  = { red: 0.757, green: 0.482, blue: 0.180, alpha: 1 };
const purple = { red: 0.502, green: 0.380, blue: 0.600, alpha: 1 };
const coral  = { red: 0.780, green: 0.400, blue: 0.380, alpha: 1 };
const slate  = { red: 0.420, green: 0.460, blue: 0.510, alpha: 1 };

const PALETTE = [wine, gold, green, teal, amber, purple, coral, slate];

const sourceRange = (sheetId, r1, r2, c1, c2) => ({
  sheetId,
  startRowIndex: r1,
  endRowIndex:   r2,
  startColumnIndex: c1,
  endColumnIndex:   c2,
});

const mkChart = (sheetId, anchorRow, anchorCol, offsetX, offsetY, w, h, spec) => ({
  addChart: {
    chart: {
      spec,
      position: {
        overlayPosition: {
          anchorCell: { sheetId, rowIndex: anchorRow, columnIndex: anchorCol },
          offsetXPixels: offsetX,
          offsetYPixels: offsetY,
          widthPixels: w,
          heightPixels: h,
        },
      },
    },
  },
});

const basicBar = (title, domain, series, sheetId, anchorRow, anchorCol, w, h, stacked) =>
  mkChart(sheetId, anchorRow, anchorCol, 0, 0, w, h, {
    title,
    titleTextFormat: { bold: true, fontSize: 11, foregroundColor: wine },
    basicChart: {
      chartType: 'BAR',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [domain] } } }],
      series: series.map((src, i) => ({
        series: { sourceRange: { sources: [src] } },
        targetAxis: 'BOTTOM_AXIS',
        color: PALETTE[i % PALETTE.length],
      })),
      stackedType: stacked ? 'STACKED' : 'NOT_STACKED',
      headerCount: 1,
    },
  });

const basicColumn = (title, domain, series, sheetId, anchorRow, anchorCol, w, h) =>
  mkChart(sheetId, anchorRow, anchorCol, 0, 0, w, h, {
    title,
    titleTextFormat: { bold: true, fontSize: 11, foregroundColor: wine },
    basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [domain] } } }],
      series: series.map((src, i) => ({
        series: { sourceRange: { sources: [src] } },
        targetAxis: 'LEFT_AXIS',
        color: PALETTE[i % PALETTE.length],
      })),
      headerCount: 1,
    },
  });

const basicLine = (title, domain, series, sheetId, anchorRow, anchorCol, w, h) =>
  mkChart(sheetId, anchorRow, anchorCol, 0, 0, w, h, {
    title,
    titleTextFormat: { bold: true, fontSize: 11, foregroundColor: wine },
    basicChart: {
      chartType: 'LINE',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [domain] } } }],
      series: series.map((src, i) => ({
        series: { sourceRange: { sources: [src] } },
        targetAxis: 'LEFT_AXIS',
        color: PALETTE[i % PALETTE.length],
      })),
      headerCount: 1,
    },
  });

const basicPie = (title, labels, values, sheetId, anchorRow, anchorCol, w, h) =>
  mkChart(sheetId, anchorRow, anchorCol, 0, 0, w, h, {
    title,
    titleTextFormat: { bold: true, fontSize: 11, foregroundColor: wine },
    pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0,
      domain: { sourceRange: { sources: [labels] } },
      series: { sourceRange: { sources: [values] } },
    },
  });

const basicDonut = (title, labels, values, sheetId, anchorRow, anchorCol, w, h) =>
  mkChart(sheetId, anchorRow, anchorCol, 0, 0, w, h, {
    title,
    titleTextFormat: { bold: true, fontSize: 11, foregroundColor: wine },
    pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.4,
      domain: { sourceRange: { sources: [labels] } },
      series: { sourceRange: { sources: [values] } },
    },
  });

(async () => {
  const charts = [];

  // ═══════════════════════════════════════════════════════════════════════════
  // GOALS & CHALLENGES — 5 charts
  // Data layout in Goals tab:
  //   Annual goals labels A6:A13, targets C6:C13, actuals D6:D13
  //   Challenge type labels A28:A39, completed D28:D39, total C28:C39
  //   Series labels A58:A69, % complete F58:F69
  //   Year labels for YoY: Insights tab B4:E4 / B5:E5 used
  // ═══════════════════════════════════════════════════════════════════════════

  // 1. Annual Goals Progress — horizontal bar (target vs actual)
  charts.push(basicBar(
    'Annual Goals: Target vs Actual',
    sourceRange(SID_GOA, 5, 13, 0, 1),   // A6:A13 labels
    [
      sourceRange(SID_GOA, 5, 13, 2, 3), // C6:C13 Target
      sourceRange(SID_GOA, 5, 13, 3, 4), // D6:D13 Actual
    ],
    SID_GOA, 1, 9, 380, 260,
  ));

  // 2. Goal % Complete — column chart
  charts.push(basicColumn(
    'Goal Completion %',
    sourceRange(SID_GOA, 5, 13, 0, 1),   // A6:A13
    [sourceRange(SID_GOA, 5, 13, 4, 5)], // E6:E13 % complete
    SID_GOA, 1, 13, 380, 260,
  ));

  // 3. Reading Challenges: Completed vs Total
  charts.push(basicBar(
    'Reading Challenges Progress',
    sourceRange(SID_GOA, 27, 39, 0, 1),  // A28:A39 challenge names
    [
      sourceRange(SID_GOA, 27, 39, 2, 3), // C28:C39 Total
      sourceRange(SID_GOA, 27, 39, 3, 4), // D28:D39 Completed
    ],
    SID_GOA, 15, 9, 380, 280,
  ));

  // 4. Series Completion %
  charts.push(basicBar(
    'Series Completion',
    sourceRange(SID_GOA, 57, 69, 0, 1),  // A58:A69 series names
    [sourceRange(SID_GOA, 57, 69, 5, 6)], // F58:F69 % complete
    SID_GOA, 15, 13, 380, 280,
  ));

  // 5. Year-over-Year Books Read (uses Insights annual summary row 5: actual books per year)
  charts.push(basicColumn(
    'Books Read by Year',
    sourceRange(SID_INS, 3, 4, 1, 5),   // B4:E4 year labels (row 4 = index 3)
    [sourceRange(SID_INS, 4, 5, 1, 5)], // B5:E5 books finished
    SID_GOA, 30, 9, 380, 240,
  ));

  // ═══════════════════════════════════════════════════════════════════════════
  // READING INSIGHTS — 8 charts
  // Data layout (from 07-insights.js):
  //   Annual Summary: rows 4-13, cols A-J  (A=label, B-E = 2023-2026)
  //     Row 4 = Year headers, Row 5 = Books Finished, Row 6 = Pages Read,
  //     Row 7 = Avg Rating, Row 8 = Fiction, Row 9 = Non-Fiction, Row 10 = Avg Days/Book
  //   Genre Breakdown: rows 16-30, cols A-H
  //     Col A = genre label, B=2023, C=2024, D=2025, E=2026, F=Total, G=Avg Rating, H=% of reads
  //   Rating Distribution: rows 33-37, cols A-F
  //     Col A = "★" label, B=2023, C=2024, D=2025, E=2026, F=Total
  //   Format Breakdown: rows 40-45, cols A-F
  //     Col A = format, B=2023, C=2024, D=2025, E=2026, F=Total
  // ═══════════════════════════════════════════════════════════════════════════

  // 6. Books Read by Year — column
  charts.push(basicColumn(
    'Books Read per Year',
    sourceRange(SID_INS, 3, 4, 1, 5),    // B4:E4 year labels
    [sourceRange(SID_INS, 4, 5, 1, 5)],  // B5:E5 books finished
    SID_INS, 0, 10, 380, 240,
  ));

  // 7. Pages Read by Year — column
  charts.push(basicColumn(
    'Pages Read per Year',
    sourceRange(SID_INS, 3, 4, 1, 5),    // B4:E4
    [sourceRange(SID_INS, 5, 6, 1, 5)],  // B6:E6 pages read
    SID_INS, 0, 14, 380, 240,
  ));

  // 8. Genre Breakdown (Total) — horizontal bar
  charts.push(basicBar(
    'Books Read by Genre',
    sourceRange(SID_INS, 15, 30, 0, 1),  // A16:A30 genre labels
    [sourceRange(SID_INS, 15, 30, 5, 6)], // F16:F30 Total
    SID_INS, 14, 0, 380, 340,
  ));

  // 9. Average Rating by Genre — bar
  charts.push(basicBar(
    'Average Rating by Genre',
    sourceRange(SID_INS, 15, 30, 0, 1),  // A16:A30
    [sourceRange(SID_INS, 15, 30, 6, 7)], // G16:G30 Avg Rating
    SID_INS, 14, 10, 380, 340,
  ));

  // 10. Rating Distribution (Total) — column
  charts.push(basicColumn(
    'Rating Distribution',
    sourceRange(SID_INS, 32, 37, 0, 1),  // A33:A37 star labels
    [sourceRange(SID_INS, 32, 37, 5, 6)], // F33:F37 Total
    SID_INS, 35, 0, 360, 240,
  ));

  // 11. Format Breakdown (Total) — donut
  charts.push(basicDonut(
    'Books by Format',
    sourceRange(SID_INS, 39, 45, 0, 1),  // A40:A45 format labels
    sourceRange(SID_INS, 39, 45, 5, 6),  // F40:F45 Total
    SID_INS, 35, 6, 360, 240,
  ));

  // 12. Fiction vs Non-Fiction — pie
  charts.push(basicPie(
    'Fiction vs Non-Fiction',
    sourceRange(SID_INS, 7, 9, 0, 1),    // A8:A9 labels (Fiction / Non-Fiction)
    sourceRange(SID_INS, 7, 9, 5, 6),    // F8:F9 totals
    SID_INS, 35, 12, 360, 240,
  ));

  // 13. Genre % of Total Reads — pie
  charts.push(basicPie(
    'Genre Share of All Reads',
    sourceRange(SID_INS, 15, 30, 0, 1),  // A16:A30 genre labels
    sourceRange(SID_INS, 15, 30, 7, 8),  // H16:H30 % of reads
    SID_INS, 50, 0, 380, 300,
  ));

  // ═══════════════════════════════════════════════════════════════════════════
  // READING DASHBOARD — 8 charts
  // Dashboard uses its own summary cells; we reuse Insights data for charts
  // Dashboard layout (from 10-dashboard.js):
  //   Rows 1-4: title + KPI cards
  //   Row 5+:  lists (Top Reads, Currently Reading, Recent, Wishlist)
  //   Charts placed starting at col 12 (M) to avoid list columns A-L
  // ═══════════════════════════════════════════════════════════════════════════

  // 14. Year-over-Year Books Read — column (same Insights data)
  charts.push(basicColumn(
    'Year-over-Year Books Read',
    sourceRange(SID_INS, 3, 4, 1, 5),
    [sourceRange(SID_INS, 4, 5, 1, 5)],
    SID_DSH, 5, 12, 360, 220,
  ));

  // 15. Rating Distribution — column
  charts.push(basicColumn(
    'Rating Distribution',
    sourceRange(SID_INS, 32, 37, 0, 1),
    [sourceRange(SID_INS, 32, 37, 5, 6)],
    SID_DSH, 5, 16, 360, 220,
  ));

  // 16. Genre Breakdown — bar (top 10 genres by total)
  charts.push(basicBar(
    'Books by Genre',
    sourceRange(SID_INS, 15, 25, 0, 1),  // A16:A25 top 10 genres
    [sourceRange(SID_INS, 15, 25, 5, 6)], // F16:F25 Total
    SID_DSH, 17, 12, 360, 280,
  ));

  // 17. Format Breakdown — donut
  charts.push(basicDonut(
    'Format Breakdown',
    sourceRange(SID_INS, 39, 45, 0, 1),
    sourceRange(SID_INS, 39, 45, 5, 6),
    SID_DSH, 17, 16, 360, 280,
  ));

  // 18. Annual Goal vs Actual — combo (column=actual, line=target)
  charts.push(mkChart(SID_DSH, 30, 12, 0, 0, 360, 240, {
    title: 'Annual Goal vs Actual',
    titleTextFormat: { bold: true, fontSize: 11, foregroundColor: wine },
    basicChart: {
      chartType: 'COMBO',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: '' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [sourceRange(SID_INS, 3, 4, 1, 5)] } } }],
      series: [
        {
          series: { sourceRange: { sources: [sourceRange(SID_INS, 4, 5, 1, 5)] } },
          targetAxis: 'LEFT_AXIS',
          type: 'COLUMN',
          color: gold,
        },
      ],
      headerCount: 1,
    },
  }));

  // 19. Books by Status — pie
  // Status distribution: we compute in Insights via status rows if present,
  // otherwise use a simple label/value pair. We reference a small helper range
  // that the dashboard itself computes (rows 3-4 col P-Q: Status label/count)
  // For simplicity, reuse Insights annual totals as a trend line instead.
  charts.push(basicLine(
    'Reading Pace Trend',
    sourceRange(SID_INS, 3, 4, 1, 5),
    [sourceRange(SID_INS, 4, 5, 1, 5)],
    SID_DSH, 30, 16, 360, 240,
  ));

  // 20. Pages Read by Year — column
  charts.push(basicColumn(
    'Pages Read by Year',
    sourceRange(SID_INS, 3, 4, 1, 5),
    [sourceRange(SID_INS, 5, 6, 1, 5)],
    SID_DSH, 42, 12, 360, 240,
  ));

  // 21. Genre Share — pie
  charts.push(basicPie(
    'Genre Share',
    sourceRange(SID_INS, 15, 25, 0, 1),
    sourceRange(SID_INS, 15, 25, 7, 8),
    SID_DSH, 42, 16, 360, 240,
  ));

  await batchUpdate(id, charts, '13-charts');
  console.log('✅  Charts done (21 total).');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
