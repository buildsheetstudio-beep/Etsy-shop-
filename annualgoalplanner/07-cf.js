'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL  = sheetMap['🎯 Goals & Milestones'];
const ACT = sheetMap['📋 Action Steps'];
const WKL = sheetMap['📅 Weekly Check-In & Reflection'];
const DASH = sheetMap['🏆 Goal Dashboard'];

function cfRule(ranges, rule, priority) {
  return { addConditionalFormatRule: { rule: { ranges, ...rule }, index: priority || 0 }};
}

function boolCondition(type, values = []) {
  return { condition: { type, values: values.map(v => ({ userEnteredValue: v })) }};
}

function textFmt(fgHex, bgHex) {
  return {
    textFormat: { foregroundColor: hex(fgHex) },
    backgroundColor: hex(bgHex),
  };
}

(async () => {
  const reqs = [];

  // ── GOALS & MILESTONES tab ──
  // Goal header rows (GSheets 3,9,15,...,57 = 0-indexed 2,8,...,56)
  // Full row range A3:K62 for status-based row coloring

  const goalFullRange = [{ sheetId: GL, startRowIndex: 2, endRowIndex: 62, startColumnIndex: 0, endColumnIndex: 11 }];

  // Achieved: green row
  reqs.push(cfRule(goalFullRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=AND($F3<>"",MOD(ROW()-2,6)=1,$F3="Achieved")']),
      format: textFmt('#1A6B3A', '#D4F5E4'),
    },
  }, 0));

  // At Risk: red row
  reqs.push(cfRule(goalFullRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=AND($F3<>"",MOD(ROW()-2,6)=1,$F3="At Risk")']),
      format: textFmt('#8B2020', '#FAE0DE'),
    },
  }, 1));

  // In Progress: warm gold row
  reqs.push(cfRule(goalFullRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=AND($F3<>"",MOD(ROW()-2,6)=1,$F3="In Progress")']),
      format: textFmt('#8B6A00', '#FBF7E4'),
    },
  }, 2));

  // On Track: pale blue row
  reqs.push(cfRule(goalFullRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=AND($F3<>"",MOD(ROW()-2,6)=1,$F3="On Track")']),
      format: textFmt('#1A3A6B', '#EEF0FF'),
    },
  }, 3));

  // Days Left col E: overdue (negative) → red
  reqs.push(cfRule(
    [{ sheetId: GL, startRowIndex: 2, endRowIndex: 62, startColumnIndex: 4, endColumnIndex: 5 }],
    {
      booleanRule: {
        ...boolCondition('CUSTOM_FORMULA', ['=AND(ISNUMBER($E3),$E3<0,MOD(ROW()-2,6)=1)']),
        format: textFmt('#8B2020', '#FAE0DE'),
      },
    }, 4
  ));

  // Days Left col E: warning ≤30 days → amber
  reqs.push(cfRule(
    [{ sheetId: GL, startRowIndex: 2, endRowIndex: 62, startColumnIndex: 4, endColumnIndex: 5 }],
    {
      booleanRule: {
        ...boolCondition('CUSTOM_FORMULA', ['=AND(ISNUMBER($E3),$E3>=0,$E3<=30,MOD(ROW()-2,6)=1)']),
        format: textFmt('#8B6A00', '#FFF3CD'),
      },
    }, 5
  ));

  // % Complete col H: 100% → green
  reqs.push(cfRule(
    [{ sheetId: GL, startRowIndex: 2, endRowIndex: 62, startColumnIndex: 7, endColumnIndex: 8 }],
    {
      booleanRule: {
        ...boolCondition('CUSTOM_FORMULA', ['=AND($H3=100,MOD(ROW()-2,6)=1)']),
        format: textFmt('#1A6B3A', '#D4F5E4'),
      },
    }, 6
  ));

  // % Complete col H: <25% → red
  reqs.push(cfRule(
    [{ sheetId: GL, startRowIndex: 2, endRowIndex: 62, startColumnIndex: 7, endColumnIndex: 8 }],
    {
      booleanRule: {
        ...boolCondition('CUSTOM_FORMULA', ['=AND(ISNUMBER($H3),$H3<25,MOD(ROW()-2,6)=1)']),
        format: textFmt('#8B2020', '#FAE0DE'),
      },
    }, 7
  ));

  // ── ACTION STEPS tab ──
  const actDataRange = [{ sheetId: ACT, startRowIndex: 4, endRowIndex: 4+30, startColumnIndex: 0, endColumnIndex: 8 }];

  // Done = TRUE: strike / green
  reqs.push(cfRule(actDataRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=$G5=TRUE']),
      format: textFmt('#1A6B3A', '#D4F5E4'),
    },
  }, 0));

  // Status = Blocked: red
  reqs.push(cfRule(actDataRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=$F5="Blocked"']),
      format: textFmt('#8B2020', '#FAE0DE'),
    },
  }, 1));

  // Overdue (due date past, not done): red
  reqs.push(cfRule(actDataRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=AND($G5<>TRUE,$D5<>"",$D5<TODAY())']),
      format: textFmt('#8B2020', '#FAE0DE'),
    },
  }, 2));

  // Priority = High (not done): amber
  reqs.push(cfRule(actDataRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=AND($G5<>TRUE,$E5="High")']),
      format: textFmt('#8B6A00', '#FFF3CD'),
    },
  }, 3));

  // Status = Done (belt-and-suspenders): green
  reqs.push(cfRule(actDataRange, {
    booleanRule: {
      ...boolCondition('CUSTOM_FORMULA', ['=$F5="Done"']),
      format: textFmt('#1A6B3A', '#D4F5E4'),
    },
  }, 4));

  // ── WEEKLY CHECK-IN tab ──
  const wklDataRange = [{ sheetId: WKL, startRowIndex: 2, endRowIndex: 54, startColumnIndex: 0, endColumnIndex: 11 }];

  // Energy (col H, index 7) ≤ 2: red
  reqs.push(cfRule(
    [{ sheetId: WKL, startRowIndex: 2, endRowIndex: 54, startColumnIndex: 7, endColumnIndex: 8 }],
    {
      booleanRule: {
        ...boolCondition('CUSTOM_FORMULA', ['=AND(ISNUMBER($H3),$H3<=2)']),
        format: textFmt('#8B2020', '#FAE0DE'),
      },
    }, 0
  ));

  // Motivation (col I, index 8) ≤ 2: red
  reqs.push(cfRule(
    [{ sheetId: WKL, startRowIndex: 2, endRowIndex: 54, startColumnIndex: 8, endColumnIndex: 9 }],
    {
      booleanRule: {
        ...boolCondition('CUSTOM_FORMULA', ['=AND(ISNUMBER($I3),$I3<=2)']),
        format: textFmt('#8B2020', '#FAE0DE'),
      },
    }, 1
  ));

  // Week Rating (col J, index 9) = 5: green
  reqs.push(cfRule(
    [{ sheetId: WKL, startRowIndex: 2, endRowIndex: 54, startColumnIndex: 9, endColumnIndex: 10 }],
    {
      booleanRule: {
        ...boolCondition('CUSTOM_FORMULA', ['=$J3=5']),
        format: textFmt('#1A6B3A', '#D4F5E4'),
      },
    }, 2
  ));

  // Quarter col K: Q1 blue, Q2 green, Q3 amber, Q4 purple
  const qColors = [
    { val: 'Q1', fg: '#0A0F3D', bg: '#EEF0FF' },
    { val: 'Q2', fg: '#1A6B3A', bg: '#D4F5E4' },
    { val: 'Q3', fg: '#8B6A00', bg: '#FFF3CD' },
    { val: 'Q4', fg: '#5A189A', bg: '#F3E6FF' },
  ];
  qColors.forEach(({ val, fg, bg }, i) => {
    reqs.push(cfRule(
      [{ sheetId: WKL, startRowIndex: 2, endRowIndex: 54, startColumnIndex: 10, endColumnIndex: 11 }],
      {
        booleanRule: {
          ...boolCondition('CUSTOM_FORMULA', [`=$K3="${val}"`]),
          format: textFmt(fg, bg),
        },
      }, 3 + i
    ));
  });

  // ── DASHBOARD: progress % col D (0-indexed 3) rows 9-18 — gradient color scale ──
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId: DASH, startRowIndex: 9, endRowIndex: 19, startColumnIndex: 3, endColumnIndex: 4 }],
      gradientRule: {
        minpoint: { color: hex('#FAE0DE'), type: 'MIN' },
        midpoint: { color: hex('#FBF7E4'), type: 'PERCENTILE', value: '50' },
        maxpoint: { color: hex('#D4F5E4'), type: 'MAX' },
      },
    },
    index: 0,
  }});

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional Formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
