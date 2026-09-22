'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const EL  = sheetMap['📋 Events Log'];
const TT  = sheetMap['✅ Task Tracker'];
const BUD = sheetMap['💰 Event Budget'];
const GL  = sheetMap['👥 Guest List & RSVP'];
const VT  = sheetMap['🏪 Vendor & Supplier Tracker'];
const ROS = sheetMap['⏱️ Run-of-Show'];
const AN  = sheetMap['📊 Event Analytics'];

(async () => {
  const reqs = [];

  // ── Events Log CF ─────────────────────────────────────────────────────────

  // Confirmed — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(EL, 2, 65, 0, 17)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L3="Confirmed"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});

  // Completed — deep green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(EL, 2, 65, 0, 17)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L3="Completed"' }] },
      format: { backgroundColor: hex(C.sageGreen), textFormat: { foregroundColor: hex(C.white) } },
    },
  }, index: 1 }});

  // Cancelled — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(EL, 2, 65, 0, 17)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L3="Cancelled"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), strikethrough: true } },
    },
  }, index: 2 }});

  // Postponed — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(EL, 2, 65, 0, 17)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L3="Postponed"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 3 }});

  // High priority — bold blush
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(EL, 2, 65, 14, 15)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
      format: { textFormat: { foregroundColor: hex(C.deepCrimson), bold: true } },
    },
  }, index: 4 }});

  // ── Task Tracker CF ───────────────────────────────────────────────────────

  // Done — grey strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TT, 4, 65, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F5="Done"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true } },
    },
  }, index: 0 }});

  // Blocked — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TT, 4, 65, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F5="Blocked"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg) } },
    },
  }, index: 1 }});

  // In Progress — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TT, 4, 65, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F5="In Progress"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 2 }});

  // Overdue — due date past, not done
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TT, 4, 65, 6, 7)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($G5<TODAY(),$F5<>"Done")' }] },
      format: { backgroundColor: hex(C.deepCrimson), textFormat: { foregroundColor: hex(C.white), bold: true } },
    },
  }, index: 3 }});

  // High priority — bold crimson
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TT, 4, 65, 4, 5)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
      format: { textFormat: { foregroundColor: hex(C.deepCrimson), bold: true } },
    },
  }, index: 4 }});

  // ── Budget CF ─────────────────────────────────────────────────────────────

  // Positive variance (under budget) — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BUD, 3, 17, 6, 7), gridRange(BUD, 19, 55, 6, 7)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});

  // Negative variance (over budget) — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BUD, 3, 17, 6, 7), gridRange(BUD, 19, 55, 6, 7)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), bold: true } },
    },
  }, index: 1 }});

  // Overdue payment — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BUD, 3, 17, 8, 9), gridRange(BUD, 19, 55, 8, 9)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Overdue' }] },
      format: { backgroundColor: hex(C.deepCrimson), textFormat: { foregroundColor: hex(C.white), bold: true } },
    },
  }, index: 2 }});

  // Paid in Full — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BUD, 3, 17, 8, 9), gridRange(BUD, 19, 55, 8, 9)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Paid in Full' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 3 }});

  // ── Guest List CF ─────────────────────────────────────────────────────────

  // Confirmed — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(GL, 4, 65, 0, 17)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F5="Confirmed"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});

  // Declined — red strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(GL, 4, 65, 0, 17)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F5="Declined"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), strikethrough: true } },
    },
  }, index: 1 }});

  // Waitlisted — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(GL, 4, 65, 0, 17)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F5="Waitlisted"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 2 }});

  // ── Vendor CF ─────────────────────────────────────────────────────────────

  // Paid in Full — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(VT, 4, 45, 0, 16)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L5="Paid in Full"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});

  // Overdue — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(VT, 4, 45, 0, 16)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L5="Overdue"' }] },
      format: { backgroundColor: hex(C.deepCrimson), textFormat: { foregroundColor: hex(C.white), bold: true } },
    },
  }, index: 1 }});

  // Deposit Paid — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(VT, 4, 45, 0, 16)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L5="Deposit Paid"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 2 }});

  // ── Run-of-Show CF ────────────────────────────────────────────────────────

  // Done — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ROS, 4, 45, 0, 12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$K5="Done"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});

  // In Progress — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ROS, 4, 45, 0, 12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$K5="In Progress"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 1 }});

  // Overlap alert — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ROS, 4, 45, 11, 12)],
    booleanRule: {
      condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'OVERLAP' }] },
      format: { backgroundColor: hex(C.deepCrimson), textFormat: { foregroundColor: hex(C.white), bold: true } },
    },
  }, index: 2 }});

  // Gap alert — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ROS, 4, 45, 11, 12)],
    booleanRule: {
      condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'GAP' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 3 }});

  // ── Analytics CF ──────────────────────────────────────────────────────────

  // High attendance (>90%) — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AN, 4, 35, 7, 8)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.9' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), bold: true } },
    },
  }, index: 0 }});

  // Low attendance (<60%) — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AN, 4, 35, 7, 8)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0.6' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg) } },
    },
  }, index: 1 }});

  // Positive net P/L — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AN, 4, 35, 10, 11)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 2 }});

  // Negative net P/L — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AN, 4, 35, 10, 11)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg) } },
    },
  }, index: 3 }});

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
