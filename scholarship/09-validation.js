'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TRK = sheetMap['📋 Application Tracker'];

const REF = "'📋 Reference Data'";

// ONE_OF_RANGE dropdown helper
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

(async () => {
  const reqs = [];

  // Application Tracker data rows: 0-indexed 3-27 (25 scholarships + buffer to 200)
  // Col C (index 2): Type — RefA (Scholarship Type) A2:A11 (10 types, header row 1)
  reqs.push(dv(TRK, 3, 200, 2, 3, '$A$3:$A$12'));   // Type (10 items)
  // Col E (index 4): Frequency — RefB B3:B8 (6 items)
  reqs.push(dv(TRK, 3, 200, 4, 5, '$B$3:$B$8'));    // Frequency (6 items)
  // Col G (index 6): Status — RefC C3:C10 (8 items) — strict
  reqs.push(dv(TRK, 3, 200, 6, 7, '$C$3:$C$10', true)); // Status (strict)
  // Col H (index 7): Priority — RefD D3:D5 (3 items) — strict
  reqs.push(dv(TRK, 3, 200, 7, 8, '$D$3:$D$5', true));  // Priority (strict)

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
