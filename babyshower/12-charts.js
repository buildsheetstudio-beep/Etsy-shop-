'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GL = sheetMap['👥 Guest List & RSVPs'];
const BU = sheetMap['💰 Budget & Expenses'];
const GF = sheetMap['🎁 Gift Tracker'];

function rgb(h) {
  return { red: parseInt(h.slice(1,3),16)/255, green: parseInt(h.slice(3,5),16)/255, blue: parseInt(h.slice(5,7),16)/255 };
}

function src(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

(async () => {
  const reqs = [];

  // ── CHART 1: RSVP Pie Chart on Guest List tab ──
  // Data source: O16:P20 (0-indexed: rows 15-20, cols 14-16)
  reqs.push({ addChart: {
    chart: {
      spec: {
        title: 'RSVP Status Breakdown',
        titleTextFormat: { foregroundColor: rgb(C.text), bold: true, fontSize: 11 },
        pieChart: {
          legendPosition: 'RIGHT_LEGEND',
          series: {
            sourceRange: { sources: [src(GL, 15, 21, 15, 16)] },
          },
          domain: {
            sourceRange: { sources: [src(GL, 15, 21, 14, 15)] },
          },
        },
        backgroundColor: rgb(C.champagne),
        fontName: 'Lato',
      },
      position: {
        overlayPosition: {
          anchorCell: { sheetId: GL, rowIndex: 2, columnIndex: 17 },
          widthPixels: 380,
          heightPixels: 280,
        },
      },
    },
  }});

  // ── CHART 2: Budget Bar Chart on Budget tab ──
  // Categories B4:B15, Budgeted C4:C15, Actual D4:D15
  reqs.push({ addChart: {
    chart: {
      spec: {
        title: 'Budget vs Actual by Category',
        titleTextFormat: { foregroundColor: rgb(C.text), bold: true, fontSize: 11 },
        basicChart: {
          chartType: 'BAR',
          legendPosition: 'BOTTOM_LEGEND',
          axis: [
            { position: 'BOTTOM_AXIS', title: 'Amount ($)' },
            { position: 'LEFT_AXIS', title: 'Category' },
          ],
          domains: [{
            domain: {
              sourceRange: { sources: [src(BU, 3, 15, 1, 2)] },
            },
          }],
          series: [
            {
              series: { sourceRange: { sources: [src(BU, 3, 15, 2, 3)] } },
              targetAxis: 'BOTTOM_AXIS',
              color: rgb(C.lilac),
            },
            {
              series: { sourceRange: { sources: [src(BU, 3, 15, 3, 4)] } },
              targetAxis: 'BOTTOM_AXIS',
              color: rgb(C.blush),
            },
          ],
          headerCount: 0,
          stackedType: 'NOT_STACKED',
        },
        backgroundColor: rgb(C.champagne),
        fontName: 'Lato',
      },
      position: {
        overlayPosition: {
          anchorCell: { sheetId: BU, rowIndex: 2, columnIndex: 10 },
          widthPixels: 480,
          heightPixels: 340,
        },
      },
    },
  }});

  // ── CHART 3: Gift Category Pie on Gift Tracker tab ──
  // Data: K4:L12 (0-indexed: rows 3-12, cols 10-12)
  reqs.push({ addChart: {
    chart: {
      spec: {
        title: 'Gifts by Category',
        titleTextFormat: { foregroundColor: rgb(C.text), bold: true, fontSize: 11 },
        pieChart: {
          legendPosition: 'RIGHT_LEGEND',
          series: {
            sourceRange: { sources: [src(GF, 3, 12, 11, 12)] },
          },
          domain: {
            sourceRange: { sources: [src(GF, 3, 12, 10, 11)] },
          },
        },
        backgroundColor: rgb(C.champagne),
        fontName: 'Lato',
      },
      position: {
        overlayPosition: {
          anchorCell: { sheetId: GF, rowIndex: 2, columnIndex: 13 },
          widthPixels: 360,
          heightPixels: 280,
        },
      },
    },
  }});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
