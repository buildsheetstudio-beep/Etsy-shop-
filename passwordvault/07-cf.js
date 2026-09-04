'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const AV  = sheetMap['🔐 Account Vault'];
const SUB = sheetMap['💳 Subscriptions & Billing'];
const DB  = sheetMap['🛡️ Security Dashboard'];

(async () => {
  const reqs = [];

  // ── Account Vault CF ──────────────────────────────────────────────────
  const DATA = gridRange(AV, 9, 109, 0, 18); // rows 10-110, cols A-R

  // 1. Critical password — bright red row
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L10="Critical"' }] },
      format: { backgroundColor: hex('#FFD0D0'), textFormat: { foregroundColor: hex(C.deepRed), bold: true } },
    },
  }, index: 0 }});

  // 2. Weak password — light red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L10="Weak"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg) } },
    },
  }, index: 1 }});

  // 3. Medium password — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L10="Medium"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 2 }});

  // 4. Strong password — pale green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$L10="Strong"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 3 }});

  // 5. No MFA — amber highlight on MFA col
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AV, 9, 109, 6, 7)], // col G
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G10="No"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 0 }});

  // 6. Inactive account — muted grey strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$M10="Inactive"' }] },
      format: { backgroundColor: hex('#F5F5F5'), textFormat: { foregroundColor: hex('#999999'), strikethrough: true } },
    },
  }, index: 4 }});

  // 7. Closed account — even more muted
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$M10="Closed"' }] },
      format: { backgroundColor: hex('#EEEEEE'), textFormat: { foregroundColor: hex('#AAAAAA'), strikethrough: true } },
    },
  }, index: 5 }});

  // 8. Password expires < 30 days — amber on expires col
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AV, 9, 109, 15, 16)], // col P
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($P10<>"",ISNUMBER(DATEVALUE($P10)),TODAY()>DATEVALUE($P10)-30)' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 0 }});

  // ── Subscriptions CF ──────────────────────────────────────────────────
  const SUB_DATA = gridRange(SUB, 9, 24, 0, 15);

  // 9. Trial ending — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [SUB_DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I10="Trial"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg), bold: true } },
    },
  }, index: 0 }});

  // 10. Cancelled — strikethrough muted
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [SUB_DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I10="Cancelled"' }] },
      format: { backgroundColor: hex('#F5F5F5'), textFormat: { foregroundColor: hex('#999999'), strikethrough: true } },
    },
  }, index: 1 }});

  // 11. Paused — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [SUB_DATA],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I10="Paused"' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 2 }});

  // 12. Active — pale green on status col only
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SUB, 9, 24, 8, 9)], // col I
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I10="Active"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), bold: true } },
    },
  }, index: 0 }});

  // ── Security Dashboard CF ─────────────────────────────────────────────
  // 13. Attention table: Weak — light red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DB, 13, 23, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D14="Weak"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex(C.errorFg) } },
    },
  }, index: 0 }});

  // 14. Attention table: Critical — bright red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DB, 13, 23, 0, 8)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D14="Critical"' }] },
      format: { backgroundColor: hex('#FFD0D0'), textFormat: { foregroundColor: hex(C.deepRed), bold: true } },
    },
  }, index: 1 }});

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
