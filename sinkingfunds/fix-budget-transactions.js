'use strict';
// Fix Couples & Family Budget Planner — Transactions tab rows 2-4
// Problem: row 2 has column headers, row 3 has KPI stats — completely backwards.
// Correct layout (frozen rows 1-4):
//   Row 1: Title (already correct)
//   Row 2: KPI labels      (#2d4270 navy, white bold)
//   Row 3: KPI values      (#f9f0d8 cream, #1b2b4b bold, large currency)
//   Row 4: Column headers  (#e0d8c8 tan, #1b2b4b bold)
//   Row 5+: Data (no changes)
const { sheets, batchUpdate, valuesBatchUpdate } = require('./lib');

const SSID = '1hhikX3AOgJQoHKoMwAakcteJ9oTv4vVPGpQf8NRjdfY';
const SID  = 3; // Transactions sheetId

// Color helpers
function rgb(hex) {
  const c = hex.replace('#', '');
  return { red: parseInt(c.slice(0,2),16)/255, green: parseInt(c.slice(2,4),16)/255, blue: parseInt(c.slice(4,6),16)/255 };
}
function gRange(r1, r2, c1, c2) {
  return { sheetId: SID, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

const PRIMARY   = '#1b2b4b';
const SECONDARY = '#2d4270';
const CREAM     = '#f9f0d8';
const TAN       = '#e0d8c8';
const WHITE     = '#ffffff';

(async () => {
  const reqs = [];
  const vals = [];

  // ── Step 1: Clear any existing merges in rows 2-4 ────────────────────────────
  // (the sheet had no merges, but unmerge is safe even if none exist)
  // We'll just overwrite — no unmerge needed since there were none.

  // ── Step 2: Merge KPI groups in rows 2-3 (0-indexed rows 1-2) ────────────────
  // 4 groups × 3 cols each = 12 cols total
  const KPI_GROUPS = [
    { col: 0, label: 'TOTAL INCOME (THIS MONTH)',    formula: `=IFERROR(SUMIFS($F$5:$F$504,$E$5:$E$504,"Income",$K$5:$K$504,'Settings & Reference'!$B$11),0)` },
    { col: 3, label: 'TOTAL EXPENSES (THIS MONTH)',  formula: `=IFERROR(SUMIFS($F$5:$F$504,$E$5:$E$504,"Expense",$K$5:$K$504,'Settings & Reference'!$B$11),0)` },
    { col: 6, label: 'PARTNER A TOTAL (ALL TIME)',   formula: `=IFERROR(SUM($I$5:$I$504),0)` },
    { col: 9, label: 'PARTNER B TOTAL (ALL TIME)',   formula: `=IFERROR(SUM($J$5:$J$504),0)` },
  ];

  KPI_GROUPS.forEach(({ col, label, formula }) => {
    const endCol = col + 3;

    // Merge label cell (row 2, 0-indexed row 1)
    reqs.push({ mergeCells: { range: gRange(1, 2, col, endCol), mergeType: 'MERGE_ALL' } });
    // Merge value cell (row 3, 0-indexed row 2)
    reqs.push({ mergeCells: { range: gRange(2, 3, col, endCol), mergeType: 'MERGE_ALL' } });

    // Style label cell
    reqs.push({ repeatCell: {
      range: gRange(1, 2, col, endCol),
      cell: { userEnteredFormat: {
        backgroundColor: rgb(SECONDARY),
        textFormat: { bold: true, fontSize: 8, foregroundColor: rgb(WHITE) },
        horizontalAlignment: 'CENTER',
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });

    // Style value cell
    reqs.push({ repeatCell: {
      range: gRange(2, 3, col, endCol),
      cell: { userEnteredFormat: {
        backgroundColor: rgb(CREAM),
        textFormat: { bold: true, fontSize: 16, foregroundColor: rgb(PRIMARY) },
        horizontalAlignment: 'CENTER',
        verticalAlignment: 'MIDDLE',
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)',
    } });

    // Values
    vals.push({ range: `'Transactions'!${colLetter(col)}2`, values: [[label]] });
    vals.push({ range: `'Transactions'!${colLetter(col)}3`, values: [[formula]] });
  });

  // ── Step 3: Column headers in row 4 (0-indexed row 3) ────────────────────────
  const HEADERS = ['#','Date','Description','Category','Type','Amount','Paid By','Account','Partner A Share','Partner B Share','Month Key','Notes'];

  reqs.push({ repeatCell: {
    range: gRange(3, 4, 0, 12),
    cell: { userEnteredFormat: {
      backgroundColor: rgb(TAN),
      textFormat: { bold: true, fontSize: 9, foregroundColor: rgb(PRIMARY) },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  } });

  vals.push({ range: `'Transactions'!A4`, values: [HEADERS] });

  // ── Step 4: Row height — label row thin, value row tall ──────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 20 },
    fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 42 },
    fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 32 },
    fields: 'pixelSize',
  } });

  function colLetter(i) { return String.fromCharCode(65 + i); }

  await batchUpdate(SSID, reqs, 'tx-fix-fmt');
  await valuesBatchUpdate(SSID, vals, 'tx-fix-vals');

  console.log('✓ Transactions rows 2-4 fixed: KPI labels (row 2) → KPI values (row 3) → Column headers (row 4)');
})().catch(e => { console.error(e.message || e); process.exit(1); });
