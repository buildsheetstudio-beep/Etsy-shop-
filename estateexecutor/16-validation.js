'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CK   = sheetMap['✅ Executor Checklist'];
const AD   = sheetMap['💼 Assets & Debts'];
const BEN  = sheetMap['👨‍👩‍👧 Beneficiaries'];
const DIST = sheetMap['🏦 Estate Distribution'];
const DA   = sheetMap['💻 Digital Assets & Online Accounts'];
const CL   = sheetMap['📞 Beneficiary Communication Log'];
const PS   = sheetMap['🏛️ Professional Services & Fees'];
const TL   = sheetMap['📅 Estate Timeline & Probate Progress'];
const DR   = sheetMap['📁 Documents Register'];
const NC   = sheetMap['📝 Notes & Contacts'];

function dr(range) {
  return { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: range }] };
}
function list(...vals) {
  return { type: 'ONE_OF_LIST', values: vals.map(v => ({ userEnteredValue: v })) };
}
function checkbox() {
  return { type: 'BOOLEAN' };
}

(async () => {
  const reqs = [];

  // ── Executor Checklist ────────────────────────────────────────────────────
  // B Phase (col 1)
  reqs.push({ setDataValidation: { range: gridRange(CK, 3, 53, 1, 2), rule: { condition: dr("='📋 Reference Data'!C2:C7"), showCustomUi: true, strict: false } } });
  // D Priority (col 3)
  reqs.push({ setDataValidation: { range: gridRange(CK, 3, 53, 3, 4), rule: { condition: dr("='📋 Reference Data'!T2:T4"), showCustomUi: true, strict: false } } });
  // E Status (col 4)
  reqs.push({ setDataValidation: { range: gridRange(CK, 3, 53, 4, 5), rule: { condition: dr("='📋 Reference Data'!D2:D7"), showCustomUi: true, strict: false } } });
  // G Done? checkbox (col 6)
  reqs.push({ setDataValidation: { range: gridRange(CK, 3, 53, 6, 7), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Assets & Debts ────────────────────────────────────────────────────────
  // C Asset Type (col 2, rows 4-15)
  reqs.push({ setDataValidation: { range: gridRange(AD, 4, 16, 2, 3), rule: { condition: dr("='📋 Reference Data'!A2:A17"), showCustomUi: true, strict: false } } });
  // G Asset Status (col 6)
  reqs.push({ setDataValidation: { range: gridRange(AD, 4, 16, 6, 7), rule: { condition: dr("='📋 Reference Data'!E2:E8"), showCustomUi: true, strict: false } } });
  // C Debt Type (col 2, rows 20-24)
  reqs.push({ setDataValidation: { range: gridRange(AD, 20, 25, 2, 3), rule: { condition: dr("='📋 Reference Data'!B2:B10"), showCustomUi: true, strict: false } } });
  // G Debt Status (col 6)
  reqs.push({ setDataValidation: { range: gridRange(AD, 20, 25, 6, 7), rule: { condition: dr("='📋 Reference Data'!F2:F7"), showCustomUi: true, strict: false } } });

  // ── Beneficiaries ─────────────────────────────────────────────────────────
  // D Type (col 3)
  reqs.push({ setDataValidation: { range: gridRange(BEN, 3, 8, 3, 4), rule: { condition: dr("='📋 Reference Data'!G2:G7"), showCustomUi: true, strict: false } } });

  // ── Estate Distribution ───────────────────────────────────────────────────
  // D Distribution Type (col 3)
  reqs.push({ setDataValidation: { range: gridRange(DIST, 3, 7, 3, 4), rule: { condition: dr("='📋 Reference Data'!H2:H8"), showCustomUi: true, strict: false } } });
  // I Status (col 8)
  reqs.push({ setDataValidation: { range: gridRange(DIST, 3, 7, 8, 9), rule: { condition: dr("='📋 Reference Data'!V2:V6"), showCustomUi: true, strict: false } } });

  // ── Digital Assets ────────────────────────────────────────────────────────
  // C Account Type (col 2)
  reqs.push({ setDataValidation: { range: gridRange(DA, 3, 18, 2, 3), rule: { condition: dr("='📋 Reference Data'!I2:I12"), showCustomUi: true, strict: false } } });
  // E Status (col 4)
  reqs.push({ setDataValidation: { range: gridRange(DA, 3, 18, 4, 5), rule: { condition: dr("='📋 Reference Data'!J2:J7"), showCustomUi: true, strict: false } } });

  // ── Communication Log ─────────────────────────────────────────────────────
  // D Comm Type (col 3)
  reqs.push({ setDataValidation: { range: gridRange(CL, 3, 9, 3, 4), rule: { condition: dr("='📋 Reference Data'!K2:K8"), showCustomUi: true, strict: false } } });
  // E Method (col 4)
  reqs.push({ setDataValidation: { range: gridRange(CL, 3, 9, 4, 5), rule: { condition: dr("='📋 Reference Data'!L2:L7"), showCustomUi: true, strict: false } } });
  // G Follow-up Required checkbox (col 6)
  reqs.push({ setDataValidation: { range: gridRange(CL, 3, 9, 6, 7), rule: { condition: checkbox(), showCustomUi: true } } });
  // I Follow-up Done checkbox (col 8)
  reqs.push({ setDataValidation: { range: gridRange(CL, 3, 9, 8, 9), rule: { condition: checkbox(), showCustomUi: true } } });

  // ── Professional Services ─────────────────────────────────────────────────
  // C Service Type (col 2)
  reqs.push({ setDataValidation: { range: gridRange(PS, 3, 8, 2, 3), rule: { condition: dr("='📋 Reference Data'!M2:M9"), showCustomUi: true, strict: false } } });
  // M Fee Status (col 12)
  reqs.push({ setDataValidation: { range: gridRange(PS, 3, 8, 12, 13), rule: { condition: dr("='📋 Reference Data'!N2:N6"), showCustomUi: true, strict: false } } });

  // ── Timeline ──────────────────────────────────────────────────────────────
  // C Phase (col 2)
  reqs.push({ setDataValidation: { range: gridRange(TL, 3, 23, 2, 3), rule: { condition: dr("='📋 Reference Data'!U2:U7"), showCustomUi: true, strict: false } } });
  // F Status (col 5)
  reqs.push({ setDataValidation: { range: gridRange(TL, 3, 23, 5, 6), rule: { condition: dr("='📋 Reference Data'!D2:D7"), showCustomUi: true, strict: false } } });

  // ── Documents Register ────────────────────────────────────────────────────
  // C Category (col 2)
  reqs.push({ setDataValidation: { range: gridRange(DR, 3, 15, 2, 3), rule: { condition: dr("='📋 Reference Data'!O2:O12"), showCustomUi: true, strict: false } } });
  // E Status (col 4)
  reqs.push({ setDataValidation: { range: gridRange(DR, 3, 15, 4, 5), rule: { condition: dr("='📋 Reference Data'!P2:P7"), showCustomUi: true, strict: false } } });

  // ── Notes & Contacts ──────────────────────────────────────────────────────
  // C Role (col 2, contacts rows 3-12)
  reqs.push({ setDataValidation: { range: gridRange(NC, 3, 13, 2, 3), rule: { condition: dr("='📋 Reference Data'!S2:S12"), showCustomUi: true, strict: false } } });
  // C Topic (col 2, notes rows 17-26)
  reqs.push({ setDataValidation: { range: gridRange(NC, 17, 27, 2, 3), rule: { condition: dr("='📋 Reference Data'!Q2:Q10"), showCustomUi: true, strict: false } } });
  // F Status (col 5, notes)
  reqs.push({ setDataValidation: { range: gridRange(NC, 17, 27, 5, 6), rule: { condition: dr("='📋 Reference Data'!R2:R5"), showCustomUi: true, strict: false } } });

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
