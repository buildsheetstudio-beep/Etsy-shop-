'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BDG = sheetMap['💰 Budget & Expenses'];
const GST = sheetMap['👨‍👩‍👧‍👦 Guest List & RSVPs'];

function src(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Budget vs Actual by Category — Grouped Column (Budget tab) ──
  // Budget table: categories col B (index 1) rows 3-14 (0-indexed), budgeted col C (index 2), actual col D (index 3)
  // GSheets rows 4-15 → 0-indexed rows 3-15
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Budget vs Actual by Category',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.23, green: 0.35, blue: 0.16 } },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'TOP_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Category' },
          { position: 'LEFT_AXIS', title: 'Amount ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(BDG, 3, 15, 1, 2)] } } }],
        series: [
          {
            series: { sourceRange: { sources: [src(BDG, 3, 15, 2, 3)] } },
            targetAxis: 'LEFT_AXIS',
            color: { red: 0.84, green: 0.79, blue: 0.66 },
          },
          {
            series: { sourceRange: { sources: [src(BDG, 3, 15, 3, 4)] } },
            targetAxis: 'LEFT_AXIS',
            color: { red: 0.23, green: 0.35, blue: 0.16 },
          },
        ],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: BDG, rowIndex: 2, columnIndex: 10 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 520, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 2: RSVP Summary — Donut (Guest List tab) ──
  // Data block: P24:Q29 (0-indexed rows 23-29, cols 15-17)
  // P24=header, P25:Q29=data (5 statuses)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'RSVP Summary',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.23, green: 0.35, blue: 0.16 } },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.5,
        series: { sourceRange: { sources: [src(GST, 24, 29, 16, 17)] } },
        domain: { sourceRange: { sources: [src(GST, 24, 29, 15, 16)] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: GST, rowIndex: 8, columnIndex: 15 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 340, heightPixels: 260,
      },
    },
  }}});

  // ── Chart 3: Guests by Family Branch — Column (Guest List tab) ──
  // Data: O10:P15 (0-indexed rows 9-15, cols 14-16)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Guests by Family Branch',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.23, green: 0.35, blue: 0.16 } },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Branch' },
          { position: 'LEFT_AXIS', title: 'Guests' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(GST, 9, 15, 14, 15)] } } }],
        series: [{
          series: { sourceRange: { sources: [src(GST, 9, 15, 15, 16)] } },
          targetAxis: 'LEFT_AXIS',
          color: { red: 0.77, green: 0.48, blue: 0.10 },
        }],
        headerCount: 1,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: GST, rowIndex: 19, columnIndex: 15 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 360, heightPixels: 240,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
