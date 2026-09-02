'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Total Funding Gap = max(0, Total Savings Goal − Total Portfolio Value)
// A4 = Total Portfolio Value (merged A4:F4)
// G4 = Total Savings Goal    (merged G4:L4)
// M4 = Total Funding Gap     (merged M4:R4)
// The original formula used B4−A4 (both in the same merged cell → always 0).
// Fix: G4−A4.

(async () => {
  await valuesBatchUpdate(id, [{
    range: "'College Savings Dashboard'!M4",
    values: [["=IFERROR(MAX(0,G4-A4),\"\")"]]
  }], 'fix-dashboard-m4');
  console.log('fix-dashboard-m4 done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
