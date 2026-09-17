'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TT  = sheetMap['✅ Task Tracker'];
const PT  = sheetMap['📦 Packing Tracker'];
const BUD = sheetMap['💰 Moving Budget'];
const QT  = sheetMap['🚛 Moving Company Quotes'];
const COA = sheetMap['📋 Change of Address'];
const UT  = sheetMap['⚡ Utilities & Services'];
const SC  = sheetMap['🏫 School & Childcare'];
const TL  = sheetMap['🗓️ Moving Timeline'];
const INV = sheetMap['📊 Inventory List'];
const NB  = sheetMap['🏘️ Neighbourhood Comparison'];

function dropdownRange(rangeStr) {
  return { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: rangeStr }] };
}
function dropdown(items) {
  return { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) };
}
function checkbox() {
  return { type: 'BOOLEAN' };
}

(async () => {
  const reqs = [];

  // ── Task Tracker ──────────────────────────────────────────────────────────
  // B6:B45 — Category
  reqs.push({ setDataValidation: { range: gridRange(TT, 5, 45, 1, 2), rule: { condition: dropdownRange("='📋 Reference Data'!$A$2:$A$9"), showCustomUi: true, strict: false } } });
  // D6:D45 — Priority
  reqs.push({ setDataValidation: { range: gridRange(TT, 5, 45, 3, 4), rule: { condition: dropdownRange("='📋 Reference Data'!$B$2:$B$4"), showCustomUi: true, strict: false } } });
  // E6:E45 — Status
  reqs.push({ setDataValidation: { range: gridRange(TT, 5, 45, 4, 5), rule: { condition: dropdownRange("='📋 Reference Data'!$C$2:$C$5"), showCustomUi: true, strict: false } } });
  // I6:I45 — Done? checkbox
  reqs.push({ setDataValidation: { range: gridRange(TT, 5, 45, 8, 9), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Packing Tracker ───────────────────────────────────────────────────────
  // B3:B22 — Room
  reqs.push({ setDataValidation: { range: gridRange(PT, 2, 22, 1, 2), rule: { condition: dropdownRange("='📋 Reference Data'!$D$2:$D$13"), showCustomUi: true, strict: false } } });
  // D3:D22 — Pack Priority
  reqs.push({ setDataValidation: { range: gridRange(PT, 2, 22, 3, 4), rule: { condition: dropdownRange("='📋 Reference Data'!$E$2:$E$4"), showCustomUi: true, strict: false } } });
  // E3:E22 — Pack Status
  reqs.push({ setDataValidation: { range: gridRange(PT, 2, 22, 4, 5), rule: { condition: dropdownRange("='📋 Reference Data'!$F$2:$F$6"), showCustomUi: true, strict: false } } });
  // F3:F22 — Fragile? checkbox
  reqs.push({ setDataValidation: { range: gridRange(PT, 2, 22, 5, 6), rule: { condition: checkbox(), showCustomUi: true } } });
  // H3:H22 — Condition
  reqs.push({ setDataValidation: { range: gridRange(PT, 2, 22, 7, 8), rule: { condition: dropdownRange("='📋 Reference Data'!$S$2:$S$5"), showCustomUi: true, strict: false } } });

  // ── Moving Budget ─────────────────────────────────────────────────────────
  // B6:B15 — Budget Category
  reqs.push({ setDataValidation: { range: gridRange(BUD, 5, 15, 1, 2), rule: { condition: dropdownRange("='📋 Reference Data'!$H$2:$H$11"), showCustomUi: true, strict: false } } });

  // ── Moving Company Quotes ──────────────────────────────────────────────────
  // F6:F8 — Includes Insurance? checkbox
  reqs.push({ setDataValidation: { range: gridRange(QT, 5, 8, 5, 6), rule: { condition: checkbox(), showCustomUi: true } } });
  // H6:H8 — Rating 1–5
  reqs.push({ setDataValidation: { range: gridRange(QT, 5, 8, 7, 8), rule: { condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: '1' }, { userEnteredValue: '5' }] }, showCustomUi: true, strict: false } } });

  // ── Change of Address ──────────────────────────────────────────────────────
  // B6:B50 — Address Category
  reqs.push({ setDataValidation: { range: gridRange(COA, 5, 50, 1, 2), rule: { condition: dropdownRange("='📋 Reference Data'!$J$2:$J$8"), showCustomUi: true, strict: false } } });
  // E6:E50 — Done? checkbox
  reqs.push({ setDataValidation: { range: gridRange(COA, 5, 50, 4, 5), rule: { condition: checkbox(), showCustomUi: true } } });
  // F6:F50 — Notify Method
  reqs.push({ setDataValidation: { range: gridRange(COA, 5, 50, 5, 6), rule: { condition: dropdownRange("='📋 Reference Data'!$I$2:$I$6"), showCustomUi: true, strict: false } } });

  // ── Utilities & Services ───────────────────────────────────────────────────
  // B6:B25 — Utility Service
  reqs.push({ setDataValidation: { range: gridRange(UT, 5, 25, 1, 2), rule: { condition: dropdownRange("='📋 Reference Data'!$K$2:$K$14"), showCustomUi: true, strict: false } } });
  // G6:G25 — Old Utility Status
  reqs.push({ setDataValidation: { range: gridRange(UT, 5, 25, 6, 7), rule: { condition: dropdownRange("='📋 Reference Data'!$L$2:$L$4"), showCustomUi: true, strict: false } } });
  // K6:K25 — New Utility Status
  reqs.push({ setDataValidation: { range: gridRange(UT, 5, 25, 10, 11), rule: { condition: dropdownRange("='📋 Reference Data'!$M$2:$M$5"), showCustomUi: true, strict: false } } });

  // ── School & Childcare ────────────────────────────────────────────────────
  // C4:C11 — School Type
  reqs.push({ setDataValidation: { range: gridRange(SC, 3, 11, 2, 3), rule: { condition: dropdownRange("='📋 Reference Data'!$N$2:$N$7"), showCustomUi: true, strict: false } } });
  // G4:G11 — School Status
  reqs.push({ setDataValidation: { range: gridRange(SC, 3, 11, 6, 7), rule: { condition: dropdownRange("='📋 Reference Data'!$O$2:$O$6"), showCustomUi: true, strict: false } } });
  // C15:C20 — Childcare Type
  reqs.push({ setDataValidation: { range: gridRange(SC, 14, 20, 2, 3), rule: { condition: dropdownRange("='📋 Reference Data'!$P$2:$P$6"), showCustomUi: true, strict: false } } });
  // G15:G20 — Childcare Status
  reqs.push({ setDataValidation: { range: gridRange(SC, 14, 20, 6, 7), rule: { condition: dropdownRange("='📋 Reference Data'!$Q$2:$Q$7"), showCustomUi: true, strict: false } } });
  // H15:H20 — Spots Available? checkbox
  reqs.push({ setDataValidation: { range: gridRange(SC, 14, 20, 7, 8), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Moving Timeline ────────────────────────────────────────────────────────
  // E3:E10 — Timeline Status
  reqs.push({ setDataValidation: { range: gridRange(TL, 2, 10, 4, 5), rule: { condition: dropdownRange("='📋 Reference Data'!$R$2:$R$4"), showCustomUi: true, strict: false } } });

  // ── Inventory List ────────────────────────────────────────────────────────
  // B6:B35 — Room
  reqs.push({ setDataValidation: { range: gridRange(INV, 5, 35, 1, 2), rule: { condition: dropdownRange("='📋 Reference Data'!$D$2:$D$13"), showCustomUi: true, strict: false } } });
  // I6:I35 — Condition
  reqs.push({ setDataValidation: { range: gridRange(INV, 5, 35, 8, 9), rule: { condition: dropdownRange("='📋 Reference Data'!$S$2:$S$5"), showCustomUi: true, strict: false } } });

  // ── Neighbourhood Comparison ───────────────────────────────────────────────
  // B4:B15 — Weight 1–3
  reqs.push({ setDataValidation: { range: gridRange(NB, 3, 15, 1, 2), rule: { condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: '1' }, { userEnteredValue: '3' }] }, showCustomUi: true, strict: false } } });
  // C4:F15 — Scores 1–5
  reqs.push({ setDataValidation: { range: gridRange(NB, 3, 15, 2, 6), rule: { condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: '1' }, { userEnteredValue: '5' }] }, showCustomUi: true, strict: false } } });

  // ── Dashboard — Move Type dropdown ────────────────────────────────────────
  const DASH = sheetMap['🏠 Dashboard'];
  reqs.push({ setDataValidation: { range: gridRange(DASH, 6, 7, 1, 2), rule: { condition: dropdownRange("='📋 Reference Data'!$G$2:$G$4"), showCustomUi: true, strict: false } } });

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
