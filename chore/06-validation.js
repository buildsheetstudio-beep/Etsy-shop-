'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];
const RWD = sheetMap['Reward & Prize Redemption'];

function dropdown(sheetId, r1, r2, c1, c2, formula) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: formula }] },
        strict: true,
        showCustomUi: true
      }
    }
  };
}

function inlineDropdown(sheetId, r1, r2, c1, c2, values) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
        strict: true,
        showCustomUi: true
      }
    }
  };
}

function checkbox(sheetId, r1, r2, c1, c2) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true }
    }
  };
}

(async () => {
  const reqs = [];

  // Block A: Category (col B=1), rows 2-15 (0-indexed 1-14, allows expansion)
  reqs.push(dropdown(TAB, 1, 15, 1, 2, "='Reference Data'!$B$2:$B$7"));

  // Block A: Type (col C=2), rows 2-15
  reqs.push(dropdown(TAB, 1, 15, 2, 3, "='Reference Data'!$A$2:$A$3"));

  // Block C: Family Member (col B=1), rows 25-200 (0-indexed 24-199)
  // Source: Block B family member names in A12:A21 (0-indexed A11:A20)
  reqs.push(dropdown(TAB, 24, 200, 1, 2, "='Chore & Allowance Tracker'!$A$12:$A$21"));

  // Block C: Chore (col C=2), rows 25-200 — source: Block A chore names A2:A9
  reqs.push(dropdown(TAB, 24, 200, 2, 3, "='Chore & Allowance Tracker'!$A$2:$A$9"));

  // Block C: Paid Out (col F=5): checkbox
  reqs.push(checkbox(TAB, 24, 200, 5, 6));

  // Reward tab: Family Member (col A=0), rows 2-50 (0-indexed 1-49)
  reqs.push(dropdown(RWD, 1, 50, 0, 1, "='Chore & Allowance Tracker'!$A$12:$A$21"));

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
