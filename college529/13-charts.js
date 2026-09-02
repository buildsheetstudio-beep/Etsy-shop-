'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_GT = sheetMap['Growth Tracker'];
const SID_DB = sheetMap['College Savings Dashboard'];
const SID_BS = sheetMap['Beneficiary Setup'];

// RGB helper: converts hex string like '#A9BDD1' → { red, green, blue }
const rgb = (h) => ({
  red:   parseInt(h.slice(1,3),16)/255,
  green: parseInt(h.slice(3,5),16)/255,
  blue:  parseInt(h.slice(5,7),16)/255,
});

// Growth Tracker projection section row indices (0-indexed):
// N_HIST = 33 → HIST_DATA_START=16, HIST_END=49
// PROJ_HEADER=50, PROJ_HDR_ROW=51, PROJ_DATA_START=52
// 60 projection rows: 0-indexed 52–111
const PROJ_DATA_START = 52;
const PROJ_DATA_END   = 112; // exclusive

const sourceRange = (sheetId, r1, r2, c1, c2) => ({
  sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2
});

(async () => {
  const fmt = [];

  // ── Chart 1: Growth Tracker — 5-Year Projection Line Chart ───────────────
  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: '5-Year Portfolio Projection (Oct 2026 – Sep 2031)',
          titleTextFormat: { bold: true, fontSize: 11 },
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Month' },
              { position: 'LEFT_AXIS', title: 'Portfolio Value ($)', format: { pattern: '"$"#,##0' } }
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [sourceRange(SID_GT, PROJ_DATA_START, PROJ_DATA_END, 1, 2)]
                }
              }
            }],
            series: [
              {
                series: { sourceRange: { sources: [sourceRange(SID_GT, PROJ_DATA_START, PROJ_DATA_END, 3, 4)] } },
                targetAxis: 'LEFT_AXIS',
                color: rgb('#A9BDD1'),
                lineStyle: { type: 'SOLID', width: 2 }
              },
              {
                series: { sourceRange: { sources: [sourceRange(SID_GT, PROJ_DATA_START, PROJ_DATA_END, 4, 5)] } },
                targetAxis: 'LEFT_AXIS',
                color: rgb('#7FA69A'),
                lineStyle: { type: 'SOLID', width: 2 }
              },
              {
                series: { sourceRange: { sources: [sourceRange(SID_GT, PROJ_DATA_START, PROJ_DATA_END, 5, 6)] } },
                targetAxis: 'LEFT_AXIS',
                color: rgb('#C39A74'),
                lineStyle: { type: 'SOLID', width: 2 }
              },
            ],
            headerCount: 0
          },
          backgroundColor: rgb('#F7F4F0')
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: SID_GT, rowIndex: 6, columnIndex: 7 },
            widthPixels: 580,
            heightPixels: 340
          }
        }
      }
    }
  });

  // ── Chart 2: Dashboard — Portfolio vs Goal by Beneficiary (Bar Chart) ────
  // Source: Beneficiary Setup rows 8-12 (0-indexed 7-11)
  //   Domain: col B (index 1) — Beneficiary Name
  //   Series 1: col O (index 14) — Current Total Saved
  //   Series 2: col L (index 11) — College Cost Goal
  fmt.push({
    addChart: {
      chart: {
        spec: {
          title: '529 Savings Progress by Beneficiary',
          titleTextFormat: { bold: true, fontSize: 11 },
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Amount ($)', format: { pattern: '"$"#,##0' } },
              { position: 'LEFT_AXIS', title: 'Beneficiary' }
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [sourceRange(SID_BS, 7, 12, 1, 2)]
                }
              }
            }],
            series: [
              {
                series: { sourceRange: { sources: [sourceRange(SID_BS, 7, 12, 14, 15)] } },
                targetAxis: 'BOTTOM_AXIS',
                color: rgb('#7FA69A')
              },
              {
                series: { sourceRange: { sources: [sourceRange(SID_BS, 7, 12, 11, 12)] } },
                targetAxis: 'BOTTOM_AXIS',
                color: rgb('#E3DFE8')
              },
            ],
            headerCount: 0
          },
          backgroundColor: rgb('#F7F4F0')
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: SID_DB, rowIndex: 36, columnIndex: 0 },
            widthPixels: 560,
            heightPixels: 300
          }
        }
      }
    }
  });

  await batchUpdate(id, fmt, '13-charts');
  console.log('13-charts done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
