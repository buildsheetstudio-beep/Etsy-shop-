'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CK   = sheetMap['✅ Executor Checklist'];
const AD   = sheetMap['💼 Assets & Debts'];
const DIST = sheetMap['🏦 Estate Distribution'];
const DA   = sheetMap['💻 Digital Assets & Online Accounts'];
const CL   = sheetMap['📞 Beneficiary Communication Log'];
const PS   = sheetMap['🏛️ Professional Services & Fees'];
const TL   = sheetMap['📅 Estate Timeline & Probate Progress'];
const DR   = sheetMap['📁 Documents Register'];
const NC   = sheetMap['📝 Notes & Contacts'];

(async () => {
  const reqs = [];

  // ── Executor Checklist ────────────────────────────────────────────────────
  // Completed — green strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(CK, 3, 53, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E4="Completed"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true } },
    },
  }, index: 0 }});
  // In Progress — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(CK, 3, 53, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E4="In Progress"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 1 }});
  // High priority not started — light red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(CK, 3, 53, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($D4="High",$E4="Not Started")' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg) } },
    },
  }, index: 2 }});
  // Overdue (not done, due date < today)
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(CK, 3, 53, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($G4=FALSE,$F4<TODAY(),$F4<>"")' }] },
      format: { backgroundColor: hex('#9B2335'), textFormat: { foregroundColor: hex(C.white), bold: true } },
    },
  }, index: 3 }});

  // ── Assets & Debts ────────────────────────────────────────────────────────
  // Transferred — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AD, 4, 16, 0, 11)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G5="Transferred"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});
  // Disputed — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AD, 4, 16, 0, 11)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G5="Disputed"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), bold: true } },
    },
  }, index: 1 }});
  // Pending Valuation — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AD, 4, 16, 0, 11)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G5="Pending Valuation"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 2 }});
  // Paid in Full (debts) — green strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AD, 20, 25, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G21="Paid in Full"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true } },
    },
  }, index: 0 }});
  // In Dispute (debts) — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AD, 20, 25, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G21="In Dispute"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), bold: true } },
    },
  }, index: 1 }});

  // ── Distribution ──────────────────────────────────────────────────────────
  // Complete — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DIST, 3, 7, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I4="Complete"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});
  // On Hold — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DIST, 3, 7, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I4="On Hold"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 1 }});

  // ── Digital Assets ────────────────────────────────────────────────────────
  // Closed — green strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DA, 3, 18, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E4="Closed"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true } },
    },
  }, index: 0 }});
  // Active — amber (needs attention)
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DA, 3, 18, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E4="Active"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 1 }});

  // ── Communication Log ─────────────────────────────────────────────────────
  // Follow-up outstanding — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(CL, 3, 9, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($G4=TRUE,$I4=FALSE)' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 0 }});
  // All resolved — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(CL, 3, 9, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I4=TRUE' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 1 }});

  // ── Professional Services ─────────────────────────────────────────────────
  // Paid — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(PS, 3, 8, 0, 15)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$M4="Paid"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});
  // Disputed — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(PS, 3, 8, 0, 15)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$M4="Disputed"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), bold: true } },
    },
  }, index: 1 }});
  // Invoiced — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(PS, 3, 8, 0, 15)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$M4="Invoiced"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 2 }});

  // ── Timeline ──────────────────────────────────────────────────────────────
  // Completed — green strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TL, 3, 23, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F4="Completed"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true } },
    },
  }, index: 0 }});
  // In Progress — amber bold
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TL, 3, 23, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F4="In Progress"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 1 }});
  // Overdue not started — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TL, 3, 23, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F4="Not Started",$D4<TODAY(),$D4<>"—")' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), bold: true } },
    },
  }, index: 2 }});

  // ── Documents Register ────────────────────────────────────────────────────
  // Required — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DR, 3, 15, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E4="Required"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), bold: true } },
    },
  }, index: 0 }});
  // Submitted / Certified Copy — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DR, 3, 15, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($E4="Submitted",$E4="Certified Copy")' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 1 }});

  // ── Notes & Contacts: note status ─────────────────────────────────────────
  // Open — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(NC, 17, 27, 0, 6)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F18="Open"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 0 }});
  // Resolved — green strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(NC, 17, 27, 0, 6)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F18="Resolved"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true } },
    },
  }, index: 1 }});

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
