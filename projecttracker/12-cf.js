'use strict';
const { sheets, hex, batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const MTL_SID  = sheetMap['Master Task Log'];
const MST_SID  = sheetMap['Milestones & Progress'];
const DB_SID   = sheetMap['Project Dashboard'];
const GANTT_SID= sheetMap['Gantt Chart'];

// Shared helper
const cfRule = (sheetId, r1, r2, c1, c2, condType, formula, bg) => ({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(sheetId, r1, r2, c1, c2)],
      booleanRule: {
        condition: { type: condType, values: formula ? [{ userEnteredValue: formula }] : [] },
        format: { backgroundColor: hex(bg) },
      },
    },
    index: 0,
  },
});

(async () => {
  const requests = [];

  // ── Master Task Log: status color on Status column (col N = 13) ───────────
  const MTL_ROWS = 3000; // rows 8-3007 = indices 7-3006
  const statusRules = [
    { status: 'Complete',    bg: C.success },
    { status: 'In Progress', bg: '#D8EBF0' },
    { status: 'Blocked',     bg: '#F5D8D5' },
    { status: 'Waiting',     bg: '#FFF3D0' },
    { status: 'Review',      bg: '#D8E8F0' },
    { status: 'Cancelled',   bg: C.altRow },
  ];

  // Color full task row based on status
  statusRules.forEach(({ status, bg }) => {
    requests.push(cfRule(MTL_SID, 7, 7+MTL_ROWS, 0, 25,
      'CUSTOM_FORMULA', `=$N8="${status}"`, bg));
  });

  // Overdue highlight (col V = 21): bright red row tint
  requests.push(cfRule(MTL_SID, 7, 7+MTL_ROWS, 0, 25,
    'CUSTOM_FORMULA', '=$V8=TRUE', '#FFE0DC'));

  // Priority colors on Priority col (G=6)
  const priorityCF = [
    { val: 'Critical', bg: '#F5C6C3' },
    { val: 'High',     bg: '#FAE0C8' },
    { val: 'Medium',   bg: '#FFF9D4' },
    { val: 'Low',      bg: '#E8F5E8' },
  ];
  priorityCF.forEach(({ val, bg }) => {
    requests.push(cfRule(MTL_SID, 7, 7+MTL_ROWS, 6, 7,
      'CUSTOM_FORMULA', `=$G8="${val}"`, bg));
  });

  // ── Milestones: status colors on full row ────────────────────────────────
  const MST_ROWS = 500;
  [
    { status: 'Complete',    bg: C.success },
    { status: 'At Risk',     bg: '#F5D8D5' },
    { status: 'Delayed',     bg: '#FFE0DC' },
    { status: 'In Progress', bg: '#D8EBF0' },
  ].forEach(({ status, bg }) => {
    requests.push(cfRule(MST_SID, 7, 7+MST_ROWS, 0, 12,
      'CUSTOM_FORMULA', `=$H8="${status}"`, bg));
  });

  // ── Project Dashboard: Health column (col H=7) color ────────────────────
  const DB_ROWS = 20;
  requests.push(cfRule(DB_SID, 8, 8+DB_ROWS, 7, 8,
    'CUSTOM_FORMULA', '=$H9="On Track"', C.success));
  requests.push(cfRule(DB_SID, 8, 8+DB_ROWS, 7, 8,
    'CUSTOM_FORMULA', '=$H9="At Risk"', C.attention));
  requests.push(cfRule(DB_SID, 8, 8+DB_ROWS, 7, 8,
    'CUSTOM_FORMULA', '=$H9="No Tasks"', C.gray));

  // ── Gantt Chart: task info rows — status and overdue highlighting ─────────
  // Col H (7) = Status in Gantt, col F(5)/G(6) = dates
  const GANTT_ROWS = 150;
  const INFO_COLS = 10;
  // Complete rows: green tint on info cols
  requests.push(cfRule(GANTT_SID, 5, 5+GANTT_ROWS, 0, INFO_COLS,
    'CUSTOM_FORMULA', '=$H6="Complete"', C.success));
  // Blocked: red tint
  requests.push(cfRule(GANTT_SID, 5, 5+GANTT_ROWS, 0, INFO_COLS,
    'CUSTOM_FORMULA', '=$H6="Blocked"', '#F5D8D5'));
  // Gantt bars: highlight week column if task spans it
  // week col K onwards (10), each week date is in row 5
  // Bar fill: if task startDt <= week_date AND week_date <= dueDt → color the cell
  // "On Track" (due >= today): primary teal
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(GANTT_SID, 5, 5+GANTT_ROWS, INFO_COLS, INFO_COLS+48)],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=AND($F6<>"",K$5>=$F6,K$5<=$G6,$H6<>"Complete",$G6>=TODAY())` }],
          },
          format: { backgroundColor: hex(C.primary) },
        },
      },
      index: 0,
    },
  });
  // "At Risk" (due < today+7, not complete): orange
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(GANTT_SID, 5, 5+GANTT_ROWS, INFO_COLS, INFO_COLS+48)],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=AND($F6<>"",K$5>=$F6,K$5<=$G6,$H6<>"Complete",$G6>=TODAY(),$G6<=TODAY()+14)` }],
          },
          format: { backgroundColor: hex(C.warning) },
        },
      },
      index: 0,
    },
  });
  // "Overdue" (due < today, not complete): red
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(GANTT_SID, 5, 5+GANTT_ROWS, INFO_COLS, INFO_COLS+48)],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=AND($F6<>"",K$5>=$F6,K$5<=$G6,$H6<>"Complete",$G6<TODAY())` }],
          },
          format: { backgroundColor: hex(C.attention) },
        },
      },
      index: 0,
    },
  });
  // "Complete" bars: muted green
  requests.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(GANTT_SID, 5, 5+GANTT_ROWS, INFO_COLS, INFO_COLS+48)],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=AND($F6<>"",K$5>=$F6,K$5<=$G6,$H6="Complete")` }],
          },
          format: { backgroundColor: hex(C.success) },
        },
      },
      index: 0,
    },
  });

  await batchUpdate(id, requests, 'cf');
  console.log('✓ Conditional Formatting complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
