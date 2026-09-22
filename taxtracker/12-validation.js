'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const LOG = sheetMap['📝 Deduction Log'];
const MIL = sheetMap['🚗 Mileage Log'];
const QTR = sheetMap['📅 Quarterly Tax'];

const REF = "'📋 Reference Data'";

function dv(sheetId, r1, r2, c1, c2, refRange, strict) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!${refRange}` }] },
      showCustomUi: true,
      strict: !!strict,
    },
  }};
}

function dvList(sheetId, r1, r2, c1, c2, values, strict) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: {
      condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
      showCustomUi: true,
      strict: !!strict,
    },
  }};
}

function dvBool(sheetId, r1, r2, c1, c2) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: { condition: { type: 'BOOLEAN' }, strict: true },
  }};
}

(async () => {
  const reqs = [];

  // ── Deduction Log ──
  // Col B (index 1): Category dropdown from Reference A2:A14
  reqs.push(dv(LOG, 2, 200, 1, 2, '$A$2:$A$14', true));
  // Col F (index 5): Documentation checkbox
  reqs.push(dvBool(LOG, 2, 200, 5, 6));

  // ── Mileage Log ──
  // Col D (index 3): Tax Year from Reference C2:C4
  reqs.push(dv(MIL, 2, 200, 3, 4, '$C$2:$C$4', true));

  // ── Quarterly Tax ──
  // Col J (index 9): Paid checkbox
  reqs.push(dvBool(QTR, 3, 7, 9, 10));

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
