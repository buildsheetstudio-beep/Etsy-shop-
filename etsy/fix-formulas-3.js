'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const yr = 2026;

(async () => {
  // Dashboard H25:H36 — Monthly Orders column
  // Original: =COUNTIFS(MONTH(range),mo,YEAR(range),yr,...) → #VALUE
  // MONTH()/YEAR() applied to a range cannot be used as COUNTIFS criteria ranges
  // Fix: SUMPRODUCT boolean multiplication
  const data = [];

  for (let mi = 0; mi < 12; mi++) {
    const mo = mi + 1;
    const row = 25 + mi;
    data.push({ range: `'Dashboard'!H${row}`, values: [[
      `=SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=${mo})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$P$6:$P$500="Received"))`,
    ]] });
  }

  await valuesBatchUpdate(id, data, 'fix-dashboard-monthly-orders');

  console.log('Dashboard monthly orders column fixed');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
