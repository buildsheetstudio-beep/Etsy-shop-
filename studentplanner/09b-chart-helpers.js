'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Add chart helper data
(async () => {
  await valuesBatchUpdate(id, [
    // Assignments-by-course helper: L4:M9 (header + 5 courses)
    {
      range: 'Assignments!L4:M9',
      values: [
        ['Course', 'Count'],
        ['English 101',   `=COUNTIF(C5:C204,"English 101")`],
        ['Calculus II',   `=COUNTIF(C5:C204,"Calculus II")`],
        ['World History', `=COUNTIF(C5:C204,"World History")`],
        ['Biology 201',   `=COUNTIF(C5:C204,"Biology 201")`],
        ['Economics 101', `=COUNTIF(C5:C204,"Economics 101")`],
      ],
    },
    // Budget chart helper: H2:J5 (header + 3 categories)
    {
      range: 'Student Budget!H2:J5',
      values: [
        ['Category', 'Budgeted', 'Actual'],
        ['Income',   '=C11',     '=D11'],
        ['Expenses', '=C27',     '=D27'],
        ['Savings',  '=C32',     '=D32'],
      ],
    },
  ], 'chart-helpers');
  console.log('Chart helper data written');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
