'use strict';
const { sheets } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));

// Cells containing any of these strings are formula errors
const ERROR_RE = /^#(VALUE!|REF!|ERROR!|N\/A|NAME\?|DIV\/0!|NUM!|NULL!)/;

function colLetter(i) {
  if (i < 26) return String.fromCharCode(65 + i);
  return String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26));
}

// Tabs to scan (skip hidden Reference Data)
const CONTENT_TABS = [
  'Master People',
  'Relationships',
  'Vital Records',
  'Locations & Migration',
  'Sources & Citations',
  'Research Log',
  'Person Profile',
  'Family Tree',
  'Search & Filter',
  'Genealogy Dashboard',
];

async function scanTab(tabName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: id,
    range: `'${tabName}'`,
    valueRenderOption: 'FORMATTED_VALUE',
  });

  const rows = res.data.values || [];
  const errors = [];

  rows.forEach((row, rIdx) => {
    row.forEach((cell, cIdx) => {
      if (typeof cell === 'string' && ERROR_RE.test(cell.trim())) {
        errors.push(`    ${colLetter(cIdx)}${rIdx + 1}: ${cell}`);
      }
    });
  });

  return { tabName, errors, rowCount: rows.length };
}

// Spot-check: data rows on these tabs must have a non-empty first data column (col B, 0-indexed 1)
// after the header block (which ends at row 7 = index 7, so data starts at index 7)
const DATA_TABS_CHECK = {
  'Master People':         { firstDataCol: 0, dataStart: 7, label: 'Person ID (col A)' },
  'Relationships':         { firstDataCol: 4, dataStart: 7, label: 'Person 2 ID (col E)' },
  'Vital Records':         { firstDataCol: 1, dataStart: 7, label: 'Person ID (col B)' },
  'Locations & Migration': { firstDataCol: 1, dataStart: 7, label: 'Person ID (col B)' },
  'Sources & Citations':   { firstDataCol: 0, dataStart: 7, label: 'Source ID (col A)' },
  'Research Log':          { firstDataCol: 1, dataStart: 7, label: 'Date (col B)' },
};

async function checkDataIntegrity(tabName, { firstDataCol, dataStart, label }) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: id,
    range: `'${tabName}'`,
    valueRenderOption: 'FORMATTED_VALUE',
  });

  const rows = (res.data.values || []).slice(dataStart);
  let orphanCount = 0;

  rows.forEach((row, rIdx) => {
    // Row has some content but the keyed column is empty
    if (row.length > 0 && row.some(c => c !== '') && (!row[firstDataCol] || row[firstDataCol] === '')) {
      orphanCount++;
    }
  });

  return orphanCount;
}

async function main() {
  console.log('═══════════════════════════════════════════════════');
  console.log('   GENEALOGY TRACKER — QC SCAN');
  console.log('═══════════════════════════════════════════════════\n');

  // ── Pass 1: Formula errors ────────────────────────────────────────────────
  console.log('Pass 1: Scanning for formula errors (#VALUE, #REF, #ERROR, etc.)\n');
  let totalErrors = 0;
  const results = [];

  for (const tab of CONTENT_TABS) {
    try {
      const { errors, rowCount } = await scanTab(tab);
      results.push({ tab, errors, rowCount });
      if (errors.length > 0) {
        console.log(`  ❌  ${tab}  →  ${errors.length} error(s)`);
        errors.forEach(e => console.log(e));
        totalErrors += errors.length;
      } else {
        console.log(`  ✓   ${tab}  →  clean (${rowCount} rows)`);
      }
    } catch (e) {
      console.log(`  ⚠️   ${tab}  →  could not read (${e.message})`);
    }
  }

  console.log();
  if (totalErrors === 0) {
    console.log('Pass 1 result: ✅  No formula errors found.\n');
  } else {
    console.log(`Pass 1 result: ❌  ${totalErrors} formula error(s) found — fix before delivery.\n`);
  }

  // ── Pass 2: Data integrity spot-checks ───────────────────────────────────
  console.log('Pass 2: Checking data integrity (orphaned rows)\n');
  let totalOrphans = 0;

  for (const [tab, spec] of Object.entries(DATA_TABS_CHECK)) {
    try {
      const orphans = await checkDataIntegrity(tab, spec);
      if (orphans > 0) {
        console.log(`  ⚠️   ${tab}  →  ${orphans} row(s) with content but empty ${spec.label}`);
        totalOrphans += orphans;
      } else {
        console.log(`  ✓   ${tab}  →  no orphaned rows`);
      }
    } catch (e) {
      console.log(`  ⚠️   ${tab}  →  could not check (${e.message})`);
    }
  }

  console.log();
  if (totalOrphans === 0) {
    console.log('Pass 2 result: ✅  No data integrity issues found.\n');
  } else {
    console.log(`Pass 2 result: ⚠️   ${totalOrphans} orphaned row(s) found — review before delivery.\n`);
  }

  // ── Summary ───────────────────────────────────────────────────────────────
  console.log('═══════════════════════════════════════════════════');
  if (totalErrors === 0 && totalOrphans === 0) {
    console.log('QC RESULT:  ✅  PASSED — spreadsheet is ready for delivery');
  } else {
    const issues = [];
    if (totalErrors > 0) issues.push(`${totalErrors} formula error(s)`);
    if (totalOrphans > 0) issues.push(`${totalOrphans} orphaned row(s)`);
    console.log(`QC RESULT:  ❌  ISSUES FOUND — ${issues.join(', ')}`);
    console.log('\nReview issues above and re-run after fixing.');
    process.exit(1);
  }
  console.log('═══════════════════════════════════════════════════');
}

main().catch(e => { console.error(e); process.exit(1); });
