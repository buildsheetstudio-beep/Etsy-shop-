'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const AV  = sheetMap['🔐 Account Vault'];
const SUB = sheetMap['💳 Subscriptions & Billing'];

function dr(range) {
  return { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: range }] };
}
function list(...vals) {
  return { type: 'ONE_OF_LIST', values: vals.map(v => ({ userEnteredValue: v })) };
}

(async () => {
  const reqs = [];

  // ── Account Vault ──────────────────────────────────────────────────────
  // E: Category (col 4)
  reqs.push({ setDataValidation: { range: gridRange(AV, 9, 109, 4, 5), rule: { condition: dr("='📋 Reference Data'!$A$2:$A$15"), showCustomUi: true, strict: false } } });
  // G: MFA Enabled (col 6)
  reqs.push({ setDataValidation: { range: gridRange(AV, 9, 109, 6, 7), rule: { condition: list('Yes','No'), showCustomUi: true, strict: true } } });
  // H: MFA Type (col 7)
  reqs.push({ setDataValidation: { range: gridRange(AV, 9, 109, 7, 8), rule: { condition: dr("='📋 Reference Data'!$D$2:$D$7"), showCustomUi: true, strict: false } } });
  // L: Password Strength (col 11)
  reqs.push({ setDataValidation: { range: gridRange(AV, 9, 109, 11, 12), rule: { condition: dr("='📋 Reference Data'!$B$2:$B$5"), showCustomUi: true, strict: false } } });
  // M: Account Status (col 12)
  reqs.push({ setDataValidation: { range: gridRange(AV, 9, 109, 12, 13), rule: { condition: dr("='📋 Reference Data'!$C$2:$C$6"), showCustomUi: true, strict: false } } });

  // ── Subscriptions ──────────────────────────────────────────────────────
  // C: Category (col 2)
  reqs.push({ setDataValidation: { range: gridRange(SUB, 9, 24, 2, 3), rule: { condition: dr("='📋 Reference Data'!$G$2:$G$13"), showCustomUi: true, strict: false } } });
  // D: Billing Cycle (col 3)
  reqs.push({ setDataValidation: { range: gridRange(SUB, 9, 24, 3, 4), rule: { condition: dr("='📋 Reference Data'!$E$2:$E$7"), showCustomUi: true, strict: false } } });
  // I: Status (col 8)
  reqs.push({ setDataValidation: { range: gridRange(SUB, 9, 24, 8, 9), rule: { condition: dr("='📋 Reference Data'!$F$2:$F$6"), showCustomUi: true, strict: false } } });
  // K: Auto-Renew (col 10)
  reqs.push({ setDataValidation: { range: gridRange(SUB, 9, 24, 10, 11), rule: { condition: list('Yes','No'), showCustomUi: true, strict: true } } });

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
