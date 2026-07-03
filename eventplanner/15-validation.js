'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const EL  = sheetMap['📋 Events Log'];
const TT  = sheetMap['✅ Task Tracker'];
const BUD = sheetMap['💰 Event Budget'];
const GL  = sheetMap['👥 Guest List & RSVP'];
const VT  = sheetMap['🏪 Vendor & Supplier Tracker'];
const ROS = sheetMap['⏱️ Run-of-Show'];
const AN  = sheetMap['📊 Event Analytics'];
const NC  = sheetMap['📝 Event Notes & Contacts'];

function dr(rangeStr) {
  return { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: rangeStr }] };
}
function checkbox() {
  return { type: 'BOOLEAN' };
}
function numBetween(min, max) {
  return { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }] };
}

(async () => {
  const reqs = [];

  // ── Events Log validations ─────────────────────────────────────────────────
  // C (Category, col 2)
  reqs.push({ setDataValidation: { range: gridRange(EL, 2, 65, 2, 3), rule: { condition: dr("='📋 Reference Data'!$A$2:$A$13"), showCustomUi: true, strict: false } } });
  // D (Type, col 3)
  reqs.push({ setDataValidation: { range: gridRange(EL, 2, 65, 3, 4), rule: { condition: dr("='📋 Reference Data'!$B$2:$B$4"), showCustomUi: true, strict: false } } });
  // L (Status, col 11)
  reqs.push({ setDataValidation: { range: gridRange(EL, 2, 65, 11, 12), rule: { condition: dr("='📋 Reference Data'!$C$2:$C$7"), showCustomUi: true, strict: false } } });
  // N (Priority, col 13 = index 13? No: A=0,B=1,...N=13)
  // Actually: A#, B Event, C Cat, D Type, E Start, F End, G Setup, H Strike, I Guests(formula), J Venue, K VenueCost, L Status, M ActualSpend(formula), N Revenue, O Priority, P Created, Q Notes
  // O = index 14
  reqs.push({ setDataValidation: { range: gridRange(EL, 2, 65, 14, 15), rule: { condition: dr("='📋 Reference Data'!$D$2:$D$4"), showCustomUi: true, strict: false } } });

  // ── Task Tracker validations ───────────────────────────────────────────────
  // C (Category, col 2)
  reqs.push({ setDataValidation: { range: gridRange(TT, 4, 65, 2, 3), rule: { condition: dr("='📋 Reference Data'!$E$2:$E$11"), showCustomUi: true, strict: false } } });
  // E (Priority, col 4)
  reqs.push({ setDataValidation: { range: gridRange(TT, 4, 65, 4, 5), rule: { condition: dr("='📋 Reference Data'!$G$2:$G$4"), showCustomUi: true, strict: false } } });
  // F (Status, col 5)
  reqs.push({ setDataValidation: { range: gridRange(TT, 4, 65, 5, 6), rule: { condition: dr("='📋 Reference Data'!$F$2:$F$6"), showCustomUi: true, strict: false } } });
  // J (Done checkbox, col 9)
  reqs.push({ setDataValidation: { range: gridRange(TT, 4, 65, 9, 10), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Budget validations ─────────────────────────────────────────────────────
  // C (Category in allocation, col 2)
  reqs.push({ setDataValidation: { range: gridRange(BUD, 3, 17, 2, 3), rule: { condition: dr("='📋 Reference Data'!$N$2:$N$13"), showCustomUi: true, strict: false } } });
  // H (Paid checkbox allocation, col 7)
  reqs.push({ setDataValidation: { range: gridRange(BUD, 3, 17, 7, 8), rule: { condition: checkbox(), showCustomUi: true } } });
  // I (Payment Status allocation, col 8)
  reqs.push({ setDataValidation: { range: gridRange(BUD, 3, 17, 8, 9), rule: { condition: dr("='📋 Reference Data'!$I$2:$I$7"), showCustomUi: true, strict: false } } });
  // C expense (Category, col 2)
  reqs.push({ setDataValidation: { range: gridRange(BUD, 19, 55, 2, 3), rule: { condition: dr("='📋 Reference Data'!$N$2:$N$13"), showCustomUi: true, strict: false } } });
  // H expense paid checkbox
  reqs.push({ setDataValidation: { range: gridRange(BUD, 19, 55, 7, 8), rule: { condition: checkbox(), showCustomUi: true } } });
  // I expense payment status
  reqs.push({ setDataValidation: { range: gridRange(BUD, 19, 55, 8, 9), rule: { condition: dr("='📋 Reference Data'!$I$2:$I$7"), showCustomUi: true, strict: false } } });

  // ── Guest List validations ─────────────────────────────────────────────────
  // F (RSVP Status, col 5)
  reqs.push({ setDataValidation: { range: gridRange(GL, 4, 65, 5, 6), rule: { condition: dr("='📋 Reference Data'!$J$2:$J$6"), showCustomUi: true, strict: false } } });
  // G (Ticket Type, col 6)
  reqs.push({ setDataValidation: { range: gridRange(GL, 4, 65, 6, 7), rule: { condition: dr("='📋 Reference Data'!$L$2:$L$8"), showCustomUi: true, strict: false } } });
  // K (Dietary Req, col 10)
  reqs.push({ setDataValidation: { range: gridRange(GL, 4, 65, 10, 11), rule: { condition: dr("='📋 Reference Data'!$K$2:$K$10"), showCustomUi: true, strict: false } } });
  // M (Seated checkbox, col 12)
  reqs.push({ setDataValidation: { range: gridRange(GL, 4, 65, 12, 13), rule: { condition: checkbox(), showCustomUi: true } } });
  // N (Badge Printed checkbox, col 13)
  reqs.push({ setDataValidation: { range: gridRange(GL, 4, 65, 13, 14), rule: { condition: checkbox(), showCustomUi: true } } });
  // O (Checked In checkbox, col 14)
  reqs.push({ setDataValidation: { range: gridRange(GL, 4, 65, 14, 15), rule: { condition: checkbox(), showCustomUi: true } } });
  // P (Plus One checkbox, col 15)
  reqs.push({ setDataValidation: { range: gridRange(GL, 4, 65, 15, 16), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Vendor validations ─────────────────────────────────────────────────────
  // C (Category, col 2)
  reqs.push({ setDataValidation: { range: gridRange(VT, 4, 45, 2, 3), rule: { condition: dr("='📋 Reference Data'!$H$2:$H$15"), showCustomUi: true, strict: false } } });
  // J (Deposit Paid checkbox, col 9)
  reqs.push({ setDataValidation: { range: gridRange(VT, 4, 45, 9, 10), rule: { condition: checkbox(), showCustomUi: true } } });
  // L (Payment Status, col 11)
  reqs.push({ setDataValidation: { range: gridRange(VT, 4, 45, 11, 12), rule: { condition: dr("='📋 Reference Data'!$I$2:$I$7"), showCustomUi: true, strict: false } } });
  // M (Contract Signed checkbox, col 12)
  reqs.push({ setDataValidation: { range: gridRange(VT, 4, 45, 12, 13), rule: { condition: checkbox(), showCustomUi: true } } });
  // O (Rating 1-5, col 14)
  reqs.push({ setDataValidation: { range: gridRange(VT, 4, 45, 14, 15), rule: { condition: numBetween(1, 5), showCustomUi: true, strict: false } } });

  // ── Run-of-Show validations ────────────────────────────────────────────────
  // C (Type, col 2)
  reqs.push({ setDataValidation: { range: gridRange(ROS, 4, 45, 2, 3), rule: { condition: dr("='📋 Reference Data'!$M$2:$M$15"), showCustomUi: true, strict: false } } });
  // K (Status, col 10)
  reqs.push({ setDataValidation: { range: gridRange(ROS, 4, 45, 10, 11), rule: { condition: dr("='📋 Reference Data'!$F$2:$F$6"), showCustomUi: true, strict: false } } });

  // ── Analytics validations ──────────────────────────────────────────────────
  // C (Category, col 2)
  reqs.push({ setDataValidation: { range: gridRange(AN, 4, 35, 2, 3), rule: { condition: dr("='📋 Reference Data'!$A$2:$A$13"), showCustomUi: true, strict: false } } });
  // D (Type, col 3)
  reqs.push({ setDataValidation: { range: gridRange(AN, 4, 35, 3, 4), rule: { condition: dr("='📋 Reference Data'!$B$2:$B$4"), showCustomUi: true, strict: false } } });
  // L (Rating 1-5, col 11)
  reqs.push({ setDataValidation: { range: gridRange(AN, 4, 35, 11, 12), rule: { condition: numBetween(1, 5), showCustomUi: true, strict: false } } });
  // M (Satisfaction 1-5, col 12)
  reqs.push({ setDataValidation: { range: gridRange(AN, 4, 35, 12, 13), rule: { condition: numBetween(1, 5), showCustomUi: true, strict: false } } });

  // ── Notes & Contacts validations ───────────────────────────────────────────
  // C (Role, col 2 in contacts section)
  reqs.push({ setDataValidation: { range: gridRange(NC, 3, 21, 2, 3), rule: { condition: dr("='📋 Reference Data'!$P$2:$P$11"), showCustomUi: true, strict: false } } });
  // C (Note Type, col 2 in notes section)
  reqs.push({ setDataValidation: { range: gridRange(NC, 23, 65, 2, 3), rule: { condition: dr("='📋 Reference Data'!$O$2:$O$8"), showCustomUi: true, strict: false } } });

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
