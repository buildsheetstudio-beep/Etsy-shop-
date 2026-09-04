'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GST = sheetMap['👰 Guest List & Seating'];
const BUD = sheetMap['💰 Budget & Expenses'];
const GFT = sheetMap['🎁 Gift Tracker & Wishlist'];

function src(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Budget Allocation Bar Chart (Budget tab) ──
  // Data: col A (category names, rows 3-14 = 0-indexed 3-15), col B (budgeted, same range)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Budget Allocation',
      titleTextFormat: { foregroundColor: hex(C.dustyRose), bold: true, fontSize: 11 },
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Amount ($)' },
          { position: 'LEFT_AXIS', title: 'Category' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(BUD, 3, 15, 0, 1)] } } }],
        series: [
          {
            series: { sourceRange: { sources: [src(BUD, 3, 15, 1, 2)] } },
            targetAxis: 'BOTTOM_AXIS',
            color: hex(C.dustyRose),
          },
          {
            series: { sourceRange: { sources: [src(BUD, 3, 15, 2, 3)] } },
            targetAxis: 'BOTTOM_AXIS',
            color: hex(C.gold),
          },
        ],
        headerCount: 0,
      },
      backgroundColor: hex(C.ivory),
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: BUD, rowIndex: 3, columnIndex: 9 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 480, heightPixels: 300,
      },
    },
  }}});

  // ── Chart 2: RSVP Donut Chart (Guest List tab) ──
  // Data at cols O-P (14-15), rows 1-6 (0-indexed 1-6): labels col O, counts col P
  // (set in 03-guestlist.js: O1=label, P1=count header, O2:O6=status, P2:P6=COUNTIF)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'RSVP Breakdown',
      titleTextFormat: { foregroundColor: hex(C.dustyRose), bold: true, fontSize: 11 },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.4,
        domain: { sourceRange: { sources: [src(GST, 1, 6, 14, 15)] } },
        series: { sourceRange: { sources: [src(GST, 1, 6, 15, 16)] } },
      },
      backgroundColor: hex(C.ivory),
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: GST, rowIndex: 1, columnIndex: 16 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 380, heightPixels: 280,
      },
    },
  }}});

  // ── Chart 3: Gift Category Column Chart (Gift Tracker tab) ──
  // Data at cols K-L (10-11), rows 0-9 (0-indexed), label K, count L
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Gift Categories',
      titleTextFormat: { foregroundColor: hex(C.dustyRose), bold: true, fontSize: 11 },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Category' },
          { position: 'LEFT_AXIS', title: 'Count' },
        ],
        domains: [{ domain: { sourceRange: { sources: [src(GFT, 1, 10, 10, 11)] } } }],
        series: [
          {
            series: { sourceRange: { sources: [src(GFT, 1, 10, 11, 12)] } },
            targetAxis: 'LEFT_AXIS',
            color: hex(C.gold),
          },
        ],
        headerCount: 1,
      },
      backgroundColor: hex(C.ivory),
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: GFT, rowIndex: 22, columnIndex: 10 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 400, heightPixels: 280,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
