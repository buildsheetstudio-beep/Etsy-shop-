'use strict';
const { valuesBatchUpdate, batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB = sheetMap['🛡️ Security Dashboard'];
const AV = "'🔐 Account Vault'";
const DS = "'🛡️ Security Dashboard'";

// The filter condition (same for all 8 columns)
const cond = `(${AV}!L10:L109="Weak")+(${AV}!L10:L109="Critical")+(${AV}!G10:G109="No")`;

// Constrained FILTER helper — limits spill to 10 rows so it never reaches
// the merged row-27 header that was blocking all FILTER formulas.
const cf = (src, fallback = '""') =>
  `=IFERROR(ARRAY_CONSTRAIN(FILTER(${AV}!${src},${cond}),10,1),${fallback})`;

(async () => {
  // 1. Clear C17 — stray duplicate FILTER formula that immediately blocked C16's spill
  await batchUpdate(id, [{
    repeatCell: {
      range: gridRange(DB, 16, 17, 2, 3),   // C17 (0-indexed row 16, col 2)
      cell: { userEnteredValue: {} },
      fields: 'userEnteredValue',
    }
  }], 'clear-C17');

  // 2. Rewrite all 8 filter columns with corrected + constrained formulas
  //    A16 additionally fixes the wrong condition range (L10:L49 → L10:L109)
  const vals = [
    { range: `${DS}!A16`, values: [[cf('B10:B109', '"✅ All accounts in good standing"')]] },
    { range: `${DS}!B16`, values: [[cf('E10:E109')]] },
    { range: `${DS}!C16`, values: [[cf('D10:D109')]] },
    { range: `${DS}!D16`, values: [[cf('L10:L109')]] },
    { range: `${DS}!E16`, values: [[cf('M10:M109')]] },
    { range: `${DS}!F16`, values: [[cf('G10:G109')]] },
    { range: `${DS}!G16`, values: [[cf('H10:H109')]] },
    // H: action-needed label — wraps a duplicate-FILTER IF with ARRAY_CONSTRAIN
    { range: `${DS}!H16`, values: [[
      `=IFERROR(ARRAY_CONSTRAIN(IF(FILTER(${AV}!L10:L109,${cond})="Weak","🔑 Update password",IF(FILTER(${AV}!L10:L109,${cond})="Critical","🚨 Update password immediately","📱 Enable MFA")),10,1),"—")`
    ]] },
  ];

  await valuesBatchUpdate(id, vals, 'fix-dashboard-filters');
  console.log('fix-dashboard-filters done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
