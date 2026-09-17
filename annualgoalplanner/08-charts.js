'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DASH = sheetMap['🏆 Goal Dashboard'];
const WKL  = sheetMap['📅 Weekly Check-In & Reflection'];

function src(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Donut — Goals by Status ──
  // Data: Dashboard H20:I27 (0-indexed rows 19-27, cols 7-9)
  // Header row 19 (H20/I20), data rows 20-26 (H21:I27)
  // Anchor: col K (index 10), row 2 (0-indexed 1)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Goals by Status',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.04, green: 0.06, blue: 0.24 } },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.45,
        series: { sourceRange: { sources: [src(DASH, 20, 27, 8, 9)] } },
        domain: { sourceRange: { sources: [src(DASH, 20, 27, 7, 8)] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DASH, rowIndex: 1, columnIndex: 10 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 380, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 2: Column — Progress by Life Area ──
  // Data: Dashboard H29:I40 (0-indexed rows 28-40, cols 7-9)
  // Header row 28, data rows 29-39
  // Anchor: col K (index 10), row 15 (0-indexed 14)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Avg Progress by Life Area',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.04, green: 0.06, blue: 0.24 } },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Life Area' },
          { position: 'LEFT_AXIS',  title: '% Complete' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(DASH, 28, 40, 7, 8)] } } }],
        series: [{
          series: { sourceRange: { sources: [src(DASH, 28, 40, 8, 9)] } },
          targetAxis: 'LEFT_AXIS',
          color: { red: 0.79, green: 0.63, blue: 0 },
        }],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DASH, rowIndex: 14, columnIndex: 10 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 440, heightPixels: 280,
      },
    },
  }}});

  // ── Chart 3: Line — Weekly Energy & Motivation Trends ──
  // Domain: Weekly A3:A54 (0-indexed rows 2-54, col 0)
  // Series 1: Energy H3:H54 (col 7)
  // Series 2: Motivation I3:I54 (col 8)
  // Anchor: col M (index 12), row 5 (0-indexed 4)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Weekly Energy & Motivation',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.04, green: 0.06, blue: 0.24 } },
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Week' },
          { position: 'LEFT_AXIS',  title: 'Score (1–5)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(WKL, 2, 54, 0, 1)] } } }],
        series: [
          {
            series: { sourceRange: { sources: [src(WKL, 2, 54, 7, 8)] } },
            targetAxis: 'LEFT_AXIS',
            color: { red: 0.79, green: 0.63, blue: 0 },
            lineStyle: { width: 2 },
          },
          {
            series: { sourceRange: { sources: [src(WKL, 2, 54, 8, 9)] } },
            targetAxis: 'LEFT_AXIS',
            color: { red: 0.04, green: 0.06, blue: 0.24 },
            lineStyle: { width: 2 },
          },
        ],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: WKL, rowIndex: 4, columnIndex: 12 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 500, heightPixels: 280,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
