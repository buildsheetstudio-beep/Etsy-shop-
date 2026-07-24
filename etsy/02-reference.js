'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C,
        SHOP_TYPES, INCOME_SOURCES, EXPENSE_CATS, FEE_TYPES, PRODUCT_TYPES,
        TAX_GROUP_MAP, TAX_GROUPS, MONTHS, INCOME_STATUSES, EXPENSE_STATUSES, IMPORT_STATUSES } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['Reference Data'];
const R   = "'Reference Data'";

(async () => {
  const data = [];

  // A: Shop Types (A1 header, A2:A8)
  data.push({ range: `${R}!A1`, values: [['Shop Type']] });
  data.push({ range: `${R}!A2:A${1+SHOP_TYPES.length}`, values: SHOP_TYPES.map(v => [v]) });

  // B: Income Sources (B1, B2:B10)
  data.push({ range: `${R}!B1`, values: [['Income Source']] });
  data.push({ range: `${R}!B2:B${1+INCOME_SOURCES.length}`, values: INCOME_SOURCES.map(v => [v]) });

  // C: Expense Categories (C1, C2:C29)
  data.push({ range: `${R}!C1`, values: [['Expense Category']] });
  data.push({ range: `${R}!C2:C${1+EXPENSE_CATS.length}`, values: EXPENSE_CATS.map(v => [v]) });

  // D: Fee Types (D1, D2:D12)
  data.push({ range: `${R}!D1`, values: [['Fee Type']] });
  data.push({ range: `${R}!D2:D${1+FEE_TYPES.length}`, values: FEE_TYPES.map(v => [v]) });

  // E: Product Types (E1, E2:E9)
  data.push({ range: `${R}!E1`, values: [['Product Type']] });
  data.push({ range: `${R}!E2:E${1+PRODUCT_TYPES.length}`, values: PRODUCT_TYPES.map(v => [v]) });

  // F: Income Statuses (F1, F2:F6)
  data.push({ range: `${R}!F1`, values: [['Income Status']] });
  data.push({ range: `${R}!F2:F${1+INCOME_STATUSES.length}`, values: INCOME_STATUSES.map(v => [v]) });

  // G: Expense Statuses (G1, G2:G6)
  data.push({ range: `${R}!G1`, values: [['Expense Status']] });
  data.push({ range: `${R}!G2:G${1+EXPENSE_STATUSES.length}`, values: EXPENSE_STATUSES.map(v => [v]) });

  // H: Import Statuses (H1, H2:H7)
  data.push({ range: `${R}!H1`, values: [['Import Status']] });
  data.push({ range: `${R}!H2:H${1+IMPORT_STATUSES.length}`, values: IMPORT_STATUSES.map(v => [v]) });

  // I: Yes/No (I1, I2:I3)
  data.push({ range: `${R}!I1`, values: [['Yes / No']] });
  data.push({ range: `${R}!I2:I3`, values: [['Yes'],['No']] });

  // J: Months (J1, J2:J13)
  data.push({ range: `${R}!J1`, values: [['Month']] });
  data.push({ range: `${R}!J2:J13`, values: MONTHS.map(v => [v]) });

  // K: Quarters (K1, K2:K5)
  data.push({ range: `${R}!K1`, values: [['Quarter']] });
  data.push({ range: `${R}!K2:K5`, values: [['Q1'],['Q2'],['Q3'],['Q4']] });

  // L: Tax Groups (L1, L2:L14)
  data.push({ range: `${R}!L1`, values: [['Tax Group']] });
  data.push({ range: `${R}!L2:L${1+TAX_GROUPS.length}`, values: TAX_GROUPS.map(v => [v]) });

  // N:O — Expense Category → Tax Group VLOOKUP table
  data.push({ range: `${R}!N1:O1`, values: [['Expense Category','Tax Group']] });
  data.push({ range: `${R}!N2:O${1+EXPENSE_CATS.length}`, values: EXPENSE_CATS.map(c => [c, TAX_GROUP_MAP[c] || 'Other / Review Required']) });

  // P: Month filter (P1, P2:P14 = "All Months" + 12 months)
  data.push({ range: `${R}!P1`, values: [['Month Filter']] });
  data.push({ range: `${R}!P2:P14`, values: [['All Months'], ...MONTHS.map(m => [m])] });

  // Q: Year filter (Q1, Q2:Q6)
  data.push({ range: `${R}!Q1`, values: [['Year Filter']] });
  data.push({ range: `${R}!Q2:Q6`, values: [['All Years'],['2023'],['2024'],['2025'],['2026']] });

  // R: Quarter filter (R1, R2:R6)
  data.push({ range: `${R}!R1`, values: [['Quarter Filter']] });
  data.push({ range: `${R}!R2:R6`, values: [['All Quarters'],['Q1'],['Q2'],['Q3'],['Q4']] });

  await valuesBatchUpdate(id, data, 'ref-values');

  // Formatting
  const reqs = [];

  // Hide the Reference Data sheet
  reqs.push({ updateSheetProperties: { properties: { sheetId: REF, hidden: true }, fields: 'hidden' } });

  // Header row style
  const hdr = { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
  }};
  reqs.push({ repeatCell: { range: gridRange(REF, 0, 1, 0, 20), cell: hdr, fields: 'userEnteredFormat' } });

  // N:O header distinct color
  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 13, 15),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 } } },
    fields: 'userEnteredFormat',
  }});

  // Column widths
  [[0,110],[1,130],[2,160],[3,140],[4,130],[5,110],[6,110],[7,110],[8,60],[9,100],[10,50],[11,170],[13,160],[14,180],[15,110],[16,80],[17,100]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Freeze row 1
  reqs.push({ updateSheetProperties: { properties: { sheetId: REF, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
