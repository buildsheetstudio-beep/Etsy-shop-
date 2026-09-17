'use strict';
// Patch: fix Contribution Log row 7 (0-indexed 6) header labels for cols I-P.
// Root cause: KPI value merges spanned gridRange(SID, 5, 7, ...) which bled into
// the header row (row 7, 0-indexed 6). Fix by unmerging those ranges and
// re-merging only through row 6 (1-indexed) = 0-indexed row 5, then rewriting headers.
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Contribution Log'];
const S = "'Contribution Log'";

(async () => {
  // Three KPI groups at cols 8, 11, 14 — each was wrongly merged rows 5-6 (0-indexed)
  // which equals rows 6-7 (1-indexed), bleeding into the header row.
  const KPI_COLS = [8, 11, 14];

  const reqs = [];

  // Step 1: Unmerge the offending ranges (rows 5-6 in 0-indexed, i.e., the old wrong merge)
  KPI_COLS.forEach(col => {
    reqs.push({ unmergeCells: { range: gridRange(SID, 5, 7, col, col+3) } });
  });

  // Step 2: Re-merge correctly — only rows 5-5 (0-indexed = row 6, 1-indexed)
  KPI_COLS.forEach(col => {
    reqs.push({ mergeCells: { range: gridRange(SID, 5, 6, col, col+3), mergeType: 'MERGE_ALL' } });
  });

  // Step 3: Reapply KPI value cell formatting for the corrected range
  KPI_COLS.forEach(col => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 5, 6, col, col+3),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary),
        textFormat: { bold: true, fontSize: 18, foregroundColor: hex(C.white) },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)',
    } });
  });

  // Step 4: Re-apply full-width header row formatting (row 7, 0-indexed 6) for ALL 16 cols
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) },
      horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy,verticalAlignment)',
  } });

  await batchUpdate(id, reqs, 'fix-cl-headers-fmt');

  // Step 5: Write header values for row 7 (1-indexed)
  const HDRS = [
    'Transaction ID','Date','Year','Month','Fund ID','Fund Name','Owner','Transaction Type',
    'Contribution Source','Amount','From Fund ID','To Fund ID','Transfer Pair ID',
    'Balance Effect','Running Fund Balance','Notes',
  ];
  await valuesBatchUpdate(id, [{ range: `${S}!A7`, values: [HDRS] }], 'fix-cl-headers-vals');

  console.log('✓ Contribution Log row 7 headers fixed (cols A–P)');
})().catch(e => { console.error(e.message || e); process.exit(1); });
