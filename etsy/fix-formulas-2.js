'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['Dashboard'];
const BS   = sheetMap['Business Setup'];

const yr = 2026;

(async () => {
  // ── 1. Dashboard A6 — Total Orders YTD (COUNTIFS(YEAR()) → SUMPRODUCT) ──
  // ── 2. Dashboard D6 — Avg Order Value (embedded COUNTIFS(YEAR()) → SUMPRODUCT) ──
  const ordersExpr = `SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$P$6:$P$500="Received"))`;

  const dashData = [];
  dashData.push({ range: `'Dashboard'!A6`, values: [[
    `=${ordersExpr}`,
  ]] });
  dashData.push({ range: `'Dashboard'!D6`, values: [[
    `=IFERROR(SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$P$6:$P$500="Received")*'Income & Orders'!$H$6:$H$500)/${ordersExpr},0)`,
  ]] });

  await valuesBatchUpdate(id, dashData, 'fix-dashboard-orders');

  // ── 3. Business Setup B11 — Remove invalid months dropdown validation ────
  // Row index 10 (0-indexed) = row 11 (1-indexed) = "Primary Niche"
  // A setDataValidation request with no `rule` field clears the validation.
  const bsReqs = [];
  bsReqs.push({ setDataValidation: { range: gridRange(BS, 10, 11, 1, 2) } });
  await batchUpdate(id, bsReqs, 'fix-bs-validation');

  console.log('Fixes applied successfully');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
