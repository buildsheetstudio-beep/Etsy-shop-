'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PH  = sheetMap['📸 Photos & Notes'];
const SC  = sheetMap['📜 Source & Citation Tracker'];
const DNA = sheetMap['🧬 DNA & Ethnicity Estimate'];
const MIG = sheetMap['🚢 Immigration & Migration Tracker'];
const LR  = sheetMap['📞 Living Relatives Directory'];

function dropdown(range, refRange, strict = true) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: refRange }] },
        showCustomUi: true,
        strict,
      },
    },
  };
}

(async () => {
  const reqs = [];

  // Photos & Notes — col A (idx 0): ID from Family Members!A
  reqs.push(dropdown(gridRange(PH,1,17,0,1), "='👨‍👩‍👧‍👦 Family Members'!$A$2:$A$100"));

  // Source & Citation — col B (idx 1): Person ID
  reqs.push(dropdown(gridRange(SC,1,26,1,2), "='👨‍👩‍👧‍👦 Family Members'!$A$2:$A$100"));
  // col C (idx 2): Source Type
  reqs.push(dropdown(gridRange(SC,1,26,2,3), "='📋 Reference Data'!$A$2:$A$11"));
  // col E (idx 4): Confidence Level
  reqs.push(dropdown(gridRange(SC,1,26,4,5), "='📋 Reference Data'!$B$2:$B$4"));

  // DNA & Ethnicity — col B (idx 1): Region/Ethnicity
  reqs.push(dropdown(gridRange(DNA,1,11,1,2), "='📋 Reference Data'!$C$2:$C$14"));

  // Immigration & Migration — col A (idx 0): Person ID
  reqs.push(dropdown(gridRange(MIG,1,13,0,1), "='👨‍👩‍👧‍👦 Family Members'!$A$2:$A$100"));
  // col B (idx 1): Migration Type
  reqs.push(dropdown(gridRange(MIG,1,13,1,2), "='📋 Reference Data'!$D$2:$D$4"));

  // Living Relatives — col B (idx 1): Relationship to User
  reqs.push(dropdown(gridRange(LR,1,16,1,2), "='📋 Reference Data'!$E$2:$E$8"));
  // col E (idx 4): Linked Family Member ID — not strict (allow blank)
  reqs.push(dropdown(gridRange(LR,1,16,4,5), "='👨‍👩‍👧‍👦 Family Members'!$A$2:$A$100", false));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
