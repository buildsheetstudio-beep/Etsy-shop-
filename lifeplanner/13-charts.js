'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WHL = sheetMap['⚖️ Life Balance Wheel'];
const HAB = sheetMap['🔁 Habit Tracker'];
const BUD = sheetMap['💰 Budget & Expense Tracker'];
const JRN = sheetMap['📓 Journal & Gratitude Log'];

function chartReq(spec, w, h, anchorSheetId, row, col) {
  return {
    addChart: {
      chart: {
        spec,
        position: {
          overlayPosition: {
            anchorCell: { sheetId: anchorSheetId, rowIndex: row, columnIndex: col },
            widthPixels: w,
            heightPixels: h
          }
        }
      }
    }
  };
}

(async () => {
  const reqs = [];

  // Chart 1: Life Balance — Current vs Goal (grouped column) on Life Balance Wheel sheet
  reqs.push(chartReq({
    title: 'Life Balance: Current vs Goal',
    basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Life Category' },
        { position: 'LEFT_AXIS', title: 'Satisfaction (1-10)' }
      ],
      domains: [{
        domain: { sourceRange: { sources: [{ sheetId: WHL, startRowIndex: 0, endRowIndex: 7, startColumnIndex: 0, endColumnIndex: 1 }] } }
      }],
      series: [
        {
          series: { sourceRange: { sources: [{ sheetId: WHL, startRowIndex: 0, endRowIndex: 7, startColumnIndex: 1, endColumnIndex: 2 }] } },
          targetAxis: 'LEFT_AXIS'
        },
        {
          series: { sourceRange: { sources: [{ sheetId: WHL, startRowIndex: 0, endRowIndex: 7, startColumnIndex: 2, endColumnIndex: 3 }] } },
          targetAxis: 'LEFT_AXIS'
        }
      ],
      headerCount: 1
    }
  }, 520, 300, WHL, 8, 0));

  // Chart 2: Habit Completion by Day (column) on Habit Tracker sheet
  reqs.push(chartReq({
    title: 'Habit Completion by Day',
    basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Date' },
        { position: 'LEFT_AXIS', title: 'Habits Completed' }
      ],
      domains: [{
        domain: { sourceRange: { sources: [{ sheetId: HAB, startRowIndex: 0, endRowIndex: 41, startColumnIndex: 1, endColumnIndex: 2 }] } }
      }],
      series: [
        {
          series: { sourceRange: { sources: [{ sheetId: HAB, startRowIndex: 0, endRowIndex: 41, startColumnIndex: 2, endColumnIndex: 3 }] } },
          targetAxis: 'LEFT_AXIS'
        }
      ],
      headerCount: 1
    }
  }, 480, 280, HAB, 42, 0));

  // Chart 3: Budget by Category (donut) on Budget sheet
  reqs.push(chartReq({
    title: 'Expenses by Category',
    pieChart: {
      legendPosition: 'LABELED_LEGEND',
      pieHole: 0.5,
      domain: { sourceRange: { sources: [{ sheetId: BUD, startRowIndex: 0, endRowIndex: 31, startColumnIndex: 1, endColumnIndex: 2 }] } },
      series: { sourceRange: { sources: [{ sheetId: BUD, startRowIndex: 0, endRowIndex: 31, startColumnIndex: 4, endColumnIndex: 5 }] } }
    }
  }, 460, 300, BUD, 32, 0));

  // Chart 4: Mood Log (line) on Journal sheet
  reqs.push(chartReq({
    title: 'Mood Log',
    basicChart: {
      chartType: 'LINE',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Date' },
        { position: 'LEFT_AXIS', title: 'Mood' }
      ],
      domains: [{
        domain: { sourceRange: { sources: [{ sheetId: JRN, startRowIndex: 0, endRowIndex: 15, startColumnIndex: 0, endColumnIndex: 1 }] } }
      }],
      series: [
        {
          series: { sourceRange: { sources: [{ sheetId: JRN, startRowIndex: 0, endRowIndex: 15, startColumnIndex: 5, endColumnIndex: 6 }] } },
          targetAxis: 'LEFT_AXIS'
        }
      ],
      headerCount: 1
    }
  }, 460, 280, JRN, 16, 0));

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
