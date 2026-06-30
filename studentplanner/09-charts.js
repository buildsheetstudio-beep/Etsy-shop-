'use strict';
const { batchUpdate, hex, COLOR } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_DASH   = sheetMap['Dashboard'];
const SID_ASSIGN = sheetMap['Assignments'];
const SID_GRADES = sheetMap['Grades & GPA'];
const SID_BUDGET = sheetMap['Student Budget'];
const SID_EXAM   = sheetMap['Exam Prep'];

function anchor(sheetId, r, c) {
  return { sheetId, rowIndex: r, columnIndex: c };
}

(async () => {
  const reqs = [];

  // ── Chart 1: Assignment Status Donut (Dashboard, anchored row 23) ──
  // Data from Dashboard!N5:O9 (rows 4-8, cols 13-14)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Assignment Status',
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.45,
            domain: { sourceRange: { sources: [{ sheetId: SID_DASH, startRowIndex: 4, endRowIndex: 9, startColumnIndex: 13, endColumnIndex: 14 }] } },
            series: { sourceRange: { sources: [{ sheetId: SID_DASH, startRowIndex: 4, endRowIndex: 9, startColumnIndex: 14, endColumnIndex: 15 }] } },
          },
          backgroundColor: hex(COLOR.iceBlue),
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(COLOR.nearBlack) },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(SID_DASH, 23, 0),
            widthPixels: 420, heightPixels: 280,
          },
        },
      },
    },
  });

  // ── Chart 2: Grade Bar Chart by Course (Dashboard, anchored row 23 col 5) ──
  // Data: Grades!A7:A11 (course names) and Grades!O7:O11 (Current %)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Current Grade by Course',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Grade %', format: { fontSize: 9 } },
              { position: 'LEFT_AXIS',   title: 'Course',  format: { fontSize: 9 } },
            ],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: SID_GRADES, startRowIndex: 6, endRowIndex: 11, startColumnIndex: 0, endColumnIndex: 1 }] } },
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: SID_GRADES, startRowIndex: 6, endRowIndex: 11, startColumnIndex: 14, endColumnIndex: 15 }] } },
              targetAxis: 'BOTTOM_AXIS',
              color: hex(COLOR.lavender),
            }],
            headerCount: 0,
          },
          backgroundColor: hex(COLOR.iceBlue),
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(COLOR.nearBlack) },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(SID_DASH, 23, 5),
            widthPixels: 420, heightPixels: 280,
          },
        },
      },
    },
  });

  // ── Chart 3: Budget Column Chart — Budgeted vs Actual ──
  // Uses contiguous helper table at Student Budget H2:J5
  // (Written by 09b-chart-helpers.js or inline below)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Budget: Planned vs Actual',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Category' },
              { position: 'LEFT_AXIS',   title: 'Amount ($)' },
            ],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: SID_BUDGET, startRowIndex: 1, endRowIndex: 5, startColumnIndex: 7, endColumnIndex: 8 }] } },
            }],
            series: [
              {
                series: { sourceRange: { sources: [{ sheetId: SID_BUDGET, startRowIndex: 1, endRowIndex: 5, startColumnIndex: 8, endColumnIndex: 9 }] } },
                targetAxis: 'LEFT_AXIS',
                color: hex(COLOR.mint),
              },
              {
                series: { sourceRange: { sources: [{ sheetId: SID_BUDGET, startRowIndex: 1, endRowIndex: 5, startColumnIndex: 9, endColumnIndex: 10 }] } },
                targetAxis: 'LEFT_AXIS',
                color: hex(COLOR.peach),
              },
            ],
            headerCount: 1,
          },
          backgroundColor: hex(COLOR.iceBlue),
          titleTextFormat: { bold: true, fontSize: 12, foregroundColor: hex(COLOR.nearBlack) },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(SID_DASH, 23, 9),
            widthPixels: 420, heightPixels: 280,
          },
        },
      },
    },
  });

  // ── Chart 4: Exam Confidence Bar Chart (Exam Prep tab) ──
  // Data: Exam Prep A6:A10 (course names), E6:E10 (confidence 1-5)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Exam Confidence by Course',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Confidence (1-5)' },
              { position: 'LEFT_AXIS',   title: 'Course' },
            ],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: SID_EXAM, startRowIndex: 5, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 1 }] } },
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: SID_EXAM, startRowIndex: 5, endRowIndex: 10, startColumnIndex: 4, endColumnIndex: 5 }] } },
              targetAxis: 'BOTTOM_AXIS',
              color: hex(COLOR.peach),
            }],
            headerCount: 0,
          },
          backgroundColor: hex(COLOR.iceBlue),
          titleTextFormat: { bold: true, fontSize: 12 },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(SID_EXAM, 40, 0),
            widthPixels: 480, heightPixels: 280,
          },
        },
      },
    },
  });

  // ── Chart 5: Assignments by Course Column (Assignments tab) ──
  // Build helper data at Assignments col L:M
  // We'll use COUNTIF-based summary written at cols L-M, rows 3-8
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Assignments by Course',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Course' },
              { position: 'LEFT_AXIS',   title: 'Count' },
            ],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: SID_ASSIGN, startRowIndex: 3, endRowIndex: 8, startColumnIndex: 11, endColumnIndex: 12 }] } },
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: SID_ASSIGN, startRowIndex: 3, endRowIndex: 8, startColumnIndex: 12, endColumnIndex: 13 }] } },
              targetAxis: 'LEFT_AXIS',
              color: hex(COLOR.skyBlue),
            }],
            headerCount: 1,
          },
          backgroundColor: hex(COLOR.iceBlue),
          titleTextFormat: { bold: true, fontSize: 12 },
        },
        position: {
          overlayPosition: {
            anchorCell: anchor(SID_ASSIGN, 6, 11),
            widthPixels: 480, heightPixels: 280,
          },
        },
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete —', reqs.length, 'charts added');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
