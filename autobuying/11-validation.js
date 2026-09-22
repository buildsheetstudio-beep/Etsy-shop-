'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const VQ = sheetMap['Vehicle Details & Quotes'];
const BS = sheetMap['Buyer Setup'];
const DB = sheetMap['Dashboard'];
const MB = sheetMap['Monthly Budget Impact'];

function listValidation(values, sheetId, r1, r2, c1, c2) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: String(v) })) },
        strict: true, showCustomUi: true
      }
    }
  };
}

function rangeValidation(sheetId, r1, r2, c1, c2, formula1, formula2) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: formula1 }] },
        strict: true, showCustomUi: true
      }
    }
  };
}

(async () => {
  const reqs = [];

  // Vehicle Details & Quotes validations
  // Col G (6) = Condition (rows 6-105, index 5-104)
  reqs.push(listValidation(['New','Certified Pre-Owned','Used','Private Party'], VQ, 5, 105, 6, 7));

  // Col H (7) = Vehicle Type
  reqs.push(listValidation(['Sedan','Hatchback','SUV','Crossover','Pickup Truck','Minivan','Coupe','Wagon','Electric Vehicle','Plug-In Hybrid','Hybrid','Luxury','Sports Car','Other'], VQ, 5, 105, 7, 8));

  // Col I (8) = Fuel Type
  reqs.push(listValidation(['Gasoline','Diesel','Hybrid','Plug-In Hybrid','Electric','Other'], VQ, 5, 105, 8, 9));

  // Col Y (24) = Loan Term
  reqs.push(listValidation(['24','36','48','60','72','84'], VQ, 5, 105, 24, 25));

  // Col AD (29) = Quote Status
  reqs.push(listValidation(['Researching','Quote Requested','Quote Received','Negotiating','Final Offer','Rejected','Selected'], VQ, 5, 105, 29, 30));

  // Buyer Setup validations
  // Preferred Purchase Method (row 23, index 22, col B = 1)
  reqs.push(listValidation(['Cash','Finance','Lease'], BS, 22, 23, 1, 2));

  // Ownership Period (row 22, index 21, col B = 1)
  reqs.push(listValidation(['1','3','5','7'], BS, 21, 22, 1, 2));

  // Dashboard controls
  // Purchase Method (row 5, index 4, col D = 3)
  reqs.push(listValidation(['Cash','Finance','Lease'], DB, 4, 5, 3, 4));

  // Ownership Horizon (row 5, index 4, col F = 5)
  reqs.push(listValidation(['1','3','5','7'], DB, 4, 5, 5, 6));

  await batchUpdate(id, reqs, 'validations');
  console.log('✓ Data Validation');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
