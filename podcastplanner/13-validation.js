'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const CAL = sheetMap['🎙 Content Calendar'];
const GST = sheetMap['👥 Guest Outreach'];
const SPO = sheetMap['💰 Sponsorship Tracker'];
const SOC = sheetMap['📱 Social Promotion'];
const EQP = sheetMap['🎧 Equipment Checklist'];

function oneOfRange(col, startRow, endRow) {
  const colLetter = String.fromCharCode(65 + col);
  return {
    condition: {
      type: 'ONE_OF_RANGE',
      values: [{ userEnteredValue: `='📋 Reference Data'!$${colLetter}$${startRow}:$${colLetter}$${endRow}` }],
    },
    strict: true,
    showCustomUi: true,
  };
}

function oneOfList(items) {
  return {
    condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) },
    strict: true,
    showCustomUi: true,
  };
}

function boolCheck() {
  return { condition: { type: 'BOOLEAN' }, strict: true };
}

function dv(rule, range) {
  return { setDataValidation: { range, rule } };
}

(async () => {
  const reqs = [];

  // ── Content Calendar ──────────────────────────────────────────────────────
  // Col B (1): Content Pillar — Ref A3:A12
  reqs.push(dv(oneOfRange(0, 3, 12), gridRange(CAL, 2, 500, 1, 2)));
  // Col C (2): Episode Type — Ref B3:B12
  reqs.push(dv(oneOfRange(1, 3, 12), gridRange(CAL, 2, 500, 2, 3)));
  // Col D (3): Target Audience — Ref C3:C10
  reqs.push(dv(oneOfRange(2, 3, 10), gridRange(CAL, 2, 500, 3, 4)));
  // Col G (6): Status — Ref D3:D8
  reqs.push(dv(oneOfRange(3, 3, 8), gridRange(CAL, 2, 500, 6, 7)));

  // ── Guest Outreach ────────────────────────────────────────────────────────
  // Col F (5): Status — Ref E3:E8
  reqs.push(dv(oneOfRange(4, 3, 8), gridRange(GST, 2, 500, 5, 6)));

  // ── Sponsorship Tracker ───────────────────────────────────────────────────
  // Col C (2): Deal Type — ONE_OF_LIST
  reqs.push(dv(
    oneOfList(['Pre-roll','Mid-roll','Post-roll','Full Episode','Affiliate']),
    gridRange(SPO, 2, 500, 2, 3),
  ));
  // Col E (4): Payment Status — Ref F3:F6
  reqs.push(dv(oneOfRange(5, 3, 6), gridRange(SPO, 2, 500, 4, 5)));

  // ── Social Promotion ──────────────────────────────────────────────────────
  // Col B (1): Platform — Ref G3:G8
  reqs.push(dv(oneOfRange(6, 3, 8), gridRange(SOC, 2, 500, 1, 2)));
  // Col C (2): Content Type — ONE_OF_LIST
  reqs.push(dv(
    oneOfList(['Clip','Quote Graphic','Story','Full Post']),
    gridRange(SOC, 2, 500, 2, 3),
  ));
  // Col E (4): Posted — checkbox
  reqs.push(dv(boolCheck(), gridRange(SOC, 2, 500, 4, 5)));

  // ── Equipment Checklist ───────────────────────────────────────────────────
  // Col B (1): Category — Ref H3:H10
  reqs.push(dv(oneOfRange(7, 3, 10), gridRange(EQP, 2, 200, 1, 2)));
  // Col C (2): Status — ONE_OF_LIST
  reqs.push(dv(
    oneOfList(['Owned','Need to Buy','Rented']),
    gridRange(EQP, 2, 200, 2, 3),
  ));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
