'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WJ = sheetMap['📅 Weekly Pregnancy Journal'];
const ST = sheetMap['🤒 Symptom & Wellbeing Tracker'];
const AP = sheetMap['🏥 Appointments Log'];
const BB = sheetMap['💰 Baby Budget'];
const NP = sheetMap['🛏️ Nursery Planner'];
const HB = sheetMap['🎒 Hospital Bag Checklist'];
const BN = sheetMap['👶 Baby Names'];
const FM = sheetMap['🍱 Freezer Meal Planner'];
const PP = sheetMap['🌙 Postpartum Recovery Tracker'];
const NC = sheetMap['📝 Notes & Contacts'];

function dr(range) {
  return { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: range }] };
}
function checkbox() {
  return { type: 'BOOLEAN' };
}
function numBetween(min, max) {
  return { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }] };
}

(async () => {
  const reqs = [];

  // ── Symptom Tracker ────────────────────────────────────────────────────────
  // C Symptom (col 2)
  reqs.push({ setDataValidation: { range: gridRange(ST, 5, 210, 2, 3), rule: { condition: dr("='📋 Reference Data'!$D$2:$D$21"), showCustomUi: true, strict: false } } });
  // D Time of Day (col 3) — free text, no validation
  // E Severity (col 4)
  reqs.push({ setDataValidation: { range: gridRange(ST, 5, 210, 4, 5), rule: { condition: dr("='📋 Reference Data'!$E$2:$E$6"), showCustomUi: true, strict: false } } });
  // I Mood 1-10 (col 8)
  reqs.push({ setDataValidation: { range: gridRange(ST, 5, 210, 8, 9), rule: { condition: numBetween(1, 10), showCustomUi: true, strict: false } } });

  // ── Appointments Log ────────────────────────────────────────────────────────
  // C Type (col 2)
  reqs.push({ setDataValidation: { range: gridRange(AP, 5, 210, 2, 3), rule: { condition: dr("='📋 Reference Data'!$G$2:$G$13"), showCustomUi: true, strict: false } } });
  // H Completed checkbox (col 7)
  reqs.push({ setDataValidation: { range: gridRange(AP, 5, 210, 7, 8), rule: { condition: checkbox(), showCustomUi: true } } });
  // I Follow-up needed checkbox (col 8)
  reqs.push({ setDataValidation: { range: gridRange(AP, 5, 210, 8, 9), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Baby Budget ────────────────────────────────────────────────────────────
  // C Priority allocation (col 2, rows 5-16)
  reqs.push({ setDataValidation: { range: gridRange(BB, 5, 17, 2, 3), rule: { condition: dr("='📋 Reference Data'!$J$2:$J$4"), showCustomUi: true, strict: false } } });
  // H Paid checkbox allocation (col 7)
  reqs.push({ setDataValidation: { range: gridRange(BB, 5, 17, 7, 8), rule: { condition: checkbox(), showCustomUi: true } } });
  // I Payment Status allocation (col 8)
  reqs.push({ setDataValidation: { range: gridRange(BB, 5, 17, 8, 9), rule: { condition: dr("='📋 Reference Data'!$H$2:$H$13"), showCustomUi: true, strict: false } } });
  // B Category expense (col 1)
  reqs.push({ setDataValidation: { range: gridRange(BB, 19, 65, 1, 2), rule: { condition: dr("='📋 Reference Data'!$H$2:$H$13"), showCustomUi: true, strict: false } } });
  // D Priority expense (col 3)
  reqs.push({ setDataValidation: { range: gridRange(BB, 19, 65, 3, 4), rule: { condition: dr("='📋 Reference Data'!$J$2:$J$4"), showCustomUi: true, strict: false } } });
  // I Gift/Registry checkbox (col 8)
  reqs.push({ setDataValidation: { range: gridRange(BB, 19, 65, 8, 9), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Nursery Planner ────────────────────────────────────────────────────────
  // C Category (col 2)
  reqs.push({ setDataValidation: { range: gridRange(NP, 5, 37, 2, 3), rule: { condition: dr("='📋 Reference Data'!$I$2:$I$11"), showCustomUi: true, strict: false } } });
  // E Priority (col 4)
  reqs.push({ setDataValidation: { range: gridRange(NP, 5, 37, 4, 5), rule: { condition: dr("='📋 Reference Data'!$J$2:$J$4"), showCustomUi: true, strict: false } } });
  // F Purchased checkbox (col 5) — wait, column 6 is G (index 6)
  reqs.push({ setDataValidation: { range: gridRange(NP, 5, 37, 6, 7), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Hospital Bag ────────────────────────────────────────────────────────────
  // B Section (col 1)
  reqs.push({ setDataValidation: { range: gridRange(HB, 4, 75, 1, 2), rule: { condition: dr("='📋 Reference Data'!$K$2:$K$7"), showCustomUi: true, strict: false } } });
  // C Packed checkbox (col 2)
  reqs.push({ setDataValidation: { range: gridRange(HB, 4, 75, 2, 3), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Baby Names ────────────────────────────────────────────────────────────
  // C Gender (col 2)
  reqs.push({ setDataValidation: { range: gridRange(BN, 4, 55, 2, 3), rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Girl' }, { userEnteredValue: 'Boy' }, { userEnteredValue: 'Neutral' }] }, showCustomUi: true, strict: false } } });
  // F Vibe (col 5)
  reqs.push({ setDataValidation: { range: gridRange(BN, 4, 55, 5, 6), rule: { condition: dr("='📋 Reference Data'!$L$2:$L$9"), showCustomUi: true, strict: false } } });
  // G-I Scores 1-10 (cols 6-8)
  [6,7,8].forEach(c => reqs.push({ setDataValidation: { range: gridRange(BN, 4, 55, c, c+1), rule: { condition: numBetween(1, 10), showCustomUi: true, strict: false } } }));

  // ── Freezer Meals ────────────────────────────────────────────────────────
  // C Category (col 2)
  reqs.push({ setDataValidation: { range: gridRange(FM, 4, 35, 2, 3), rule: { condition: dr("='📋 Reference Data'!$M$2:$M$8"), showCustomUi: true, strict: false } } });
  // G Status (col 6)
  reqs.push({ setDataValidation: { range: gridRange(FM, 4, 35, 6, 7), rule: { condition: dr("='📋 Reference Data'!$O$2:$O$5"), showCustomUi: true, strict: false } } });
  // H Dietary Tags (col 7)
  reqs.push({ setDataValidation: { range: gridRange(FM, 4, 35, 7, 8), rule: { condition: dr("='📋 Reference Data'!$N$2:$N$7"), showCustomUi: true, strict: false } } });

  // ── Postpartum Recovery ────────────────────────────────────────────────────
  // C Wellbeing (col 2)
  reqs.push({ setDataValidation: { range: gridRange(PP, 5, 17, 2, 3), rule: { condition: dr("='📋 Reference Data'!$P$2:$P$7"), showCustomUi: true, strict: false } } });
  // G Feeding method (col 6)
  reqs.push({ setDataValidation: { range: gridRange(PP, 5, 17, 6, 7), rule: { condition: dr("='📋 Reference Data'!$Q$2:$Q$5"), showCustomUi: true, strict: false } } });
  // E Energy 1-10 (col 4)
  reqs.push({ setDataValidation: { range: gridRange(PP, 5, 17, 4, 5), rule: { condition: numBetween(1, 10), showCustomUi: true, strict: false } } });
  // F Pain level 0-10 (col 5)
  reqs.push({ setDataValidation: { range: gridRange(PP, 5, 17, 5, 6), rule: { condition: numBetween(0, 10), showCustomUi: true, strict: false } } });

  // ── Notes & Contacts ────────────────────────────────────────────────────────
  // C Role (col 2, contacts rows 3-18)
  reqs.push({ setDataValidation: { range: gridRange(NC, 3, 19, 2, 3), rule: { condition: dr("='📋 Reference Data'!$R$2:$R$11"), showCustomUi: true, strict: false } } });
  // H Emergency Contact checkbox (col 7, contacts)
  reqs.push({ setDataValidation: { range: gridRange(NC, 3, 19, 7, 8), rule: { condition: checkbox(), showCustomUi: true } } });
  // C Topic (col 2, notes rows 38-69)
  reqs.push({ setDataValidation: { range: gridRange(NC, 37, 70, 2, 3), rule: { condition: dr("='📋 Reference Data'!$S$2:$S$9"), showCustomUi: true, strict: false } } });
  // D Status (col 3, notes)
  reqs.push({ setDataValidation: { range: gridRange(NC, 37, 70, 3, 4), rule: { condition: dr("='📋 Reference Data'!$T$2:$T$4"), showCustomUi: true, strict: false } } });

  // ── Weekly Journal ────────────────────────────────────────────────────────
  // No strict dropdowns needed — free text columns

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
