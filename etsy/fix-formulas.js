'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Reusable expression fragments (no leading =)
const grossRevExpr = `SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=2026)*'Income & Orders'!$H$6:$H$500)`;
const totalDeductibleExpr = `SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=2026)*('Expense Log'!$I$6:$I$500="Yes")*'Expense Log'!$H$6:$H$500)`;
const totalExpensesExpr   = `SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=2026)*'Expense Log'!$H$6:$H$500)`;
const totalFeeExpr = `SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=2026)*(ISNUMBER(MATCH('Expense Log'!$E$6:$E$500,{"Etsy Transaction Fees","Etsy Listing Fees","Etsy Payment Processing Fees","Etsy Ads","Offsite Ads"},0)))*'Expense Log'!$H$6:$H$500)`;

(async () => {
  const data = [];

  // ── 1. Monthly Dashboard E4 — Orders (COUNTIFS with YEAR() → SUMPRODUCT) ──
  data.push({ range: `'Monthly Dashboard'!E4`, values: [[
    `=SUMPRODUCT((TEXT('Income & Orders'!$B$6:$B$500,"mmmm")='Monthly Dashboard'!C$2)*(YEAR('Income & Orders'!$B$6:$B$500)=2026)*('Income & Orders'!$P$6:$P$500="Received"))`,
  ]] });

  // ── 2. Quarterly Summary C8:F8 — Order Count per quarter ─────────────────
  // (COUNTIFS with ISNUMBER(MATCH()) as criteria → SUMPRODUCT)
  data.push({ range: `'Quarterly Summary'!C8`, values: [[
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{1,2,3},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=2026)*('Income & Orders'!$P$6:$P$500="Received"))`,
  ]] });
  data.push({ range: `'Quarterly Summary'!D8`, values: [[
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{4,5,6},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=2026)*('Income & Orders'!$P$6:$P$500="Received"))`,
  ]] });
  data.push({ range: `'Quarterly Summary'!E8`, values: [[
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{7,8,9},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=2026)*('Income & Orders'!$P$6:$P$500="Received"))`,
  ]] });
  data.push({ range: `'Quarterly Summary'!F8`, values: [[
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{10,11,12},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=2026)*('Income & Orders'!$P$6:$P$500="Received"))`,
  ]] });

  // ── 3. Multi-Year Comparison E7 — 2026 Order Count ───────────────────────
  // (COUNTIFS with YEAR() as criteria range → SUMPRODUCT)
  data.push({ range: `'Multi-Year Comparison'!E7`, values: [[
    `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=2026)*('Income & Orders'!$P$6:$P$500="Received"))`,
  ]] });

  // ── 4. Etsy Fee Analysis C3 — Fees as % of Revenue KPI ──────────────────
  // (=IFERROR(=SUMPRODUCT.../=SUMPRODUCT...) → strip leading = from sub-exprs)
  data.push({ range: `'Etsy Fee Analysis'!C3`, values: [[
    `=IFERROR(${totalFeeExpr}/${grossRevExpr},0)`,
  ]] });

  // ── 5. Etsy Fee Analysis G8:G12 — % of Revenue for each fee type ─────────
  for (let r = 8; r <= 12; r++) {
    data.push({ range: `'Etsy Fee Analysis'!G${r}`, values: [[
      `=IFERROR(F${r}/${grossRevExpr},0)`,
    ]] });
  }

  // ── 6. Etsy Fee Analysis G18 — Total Fees % of Revenue ───────────────────
  data.push({ range: `'Etsy Fee Analysis'!G18`, values: [[
    `=IFERROR(F18/${grossRevExpr},0)`,
  ]] });

  // ── 7. Etsy Fee Analysis H34 — Monthly Total Fee Rate ────────────────────
  data.push({ range: `'Etsy Fee Analysis'!H34`, values: [[
    `=IFERROR(G34/${grossRevExpr},0)`,
  ]] });

  // ── 8. Tax Deduction Summary C3 — Net Business Income ────────────────────
  // (==SUMPRODUCT...-=SUMPRODUCT... → correct subtraction)
  data.push({ range: `'Tax Deduction Summary'!C3`, values: [[
    `=${grossRevExpr}-${totalDeductibleExpr}`,
  ]] });

  // ── 9. Tax Deduction Summary E3 — Deductible % of Total Expenses ─────────
  // (=IFERROR(=SUMPRODUCT.../=SUMPRODUCT...) → strip leading = from sub-exprs)
  data.push({ range: `'Tax Deduction Summary'!E3`, values: [[
    `=IFERROR(${totalDeductibleExpr}/${totalExpensesExpr},0)`,
  ]] });

  await valuesBatchUpdate(id, data, 'fix-formulas');
  console.log('Formula fixes applied successfully');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
