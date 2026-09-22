'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS, ENTERPRISES, INCOME_CATS, EXPENSE_CATS, TAX_GROUPS, PAYMENT_METHODS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['Reference Data'];
const R = "'Reference Data'";

(async () => {
  const data = [];

  // A: Enterprises (A1 header, A2:A16 = 15 enterprises)
  data.push({ range: `${R}!A1`, values: [['Enterprise']] });
  data.push({ range: `${R}!A2:A16`, values: ENTERPRISES.map(e => [e]) });

  // C: Income Categories (C1 header, C2:C18)
  data.push({ range: `${R}!C1`, values: [['Income Category']] });
  data.push({ range: `${R}!C2:C18`, values: INCOME_CATS.map(c => [c]) });

  // E: Expense Categories (E1 header, E2:E31)
  data.push({ range: `${R}!E1`, values: [['Expense Category']] });
  data.push({ range: `${R}!E2:E31`, values: EXPENSE_CATS.map(c => [c]) });

  // G: Payment Methods (G1 header, G2:G9)
  data.push({ range: `${R}!G1`, values: [['Payment Method']] });
  data.push({ range: `${R}!G2:G9`, values: PAYMENT_METHODS.map(m => [m]) });

  // H: Income Status (H1, H2:H6)
  data.push({ range: `${R}!H1`, values: [['Income Status']] });
  data.push({ range: `${R}!H2:H6`, values: [['Received'],['Pending'],['Partially Received'],['Refunded'],['Cancelled']] });

  // I: Expense Status (I1, I2:I6)
  data.push({ range: `${R}!I1`, values: [['Expense Status']] });
  data.push({ range: `${R}!I2:I6`, values: [['Paid'],['Pending'],['Partially Paid'],['Refunded'],['Cancelled']] });

  // J: Yes/No (J1, J2:J3)
  data.push({ range: `${R}!J1`, values: [['Yes / No']] });
  data.push({ range: `${R}!J2:J3`, values: [['Yes'],['No']] });

  // L: Month Names (L1, L2:L13)
  data.push({ range: `${R}!L1`, values: [['Month']] });
  data.push({ range: `${R}!L2:L13`, values: MONTHS.map(m => [m]) });

  // M: Month Numbers (M1, M2:M13)
  data.push({ range: `${R}!M1`, values: [['Month #']] });
  data.push({ range: `${R}!M2:M13`, values: MONTHS.map((_, i) => [i+1]) });

  // N: Quarter labels (N1, N2:N5)
  data.push({ range: `${R}!N1`, values: [['Quarter']] });
  data.push({ range: `${R}!N2:N5`, values: [['Q1'],['Q2'],['Q3'],['Q4']] });

  // P:Q — Tax Group VLOOKUP table (P1:Q1 headers, P2:Q31 data)
  data.push({ range: `${R}!P1:Q1`, values: [['Expense Category','Tax Group']] });
  data.push({ range: `${R}!P2:Q31`, values: EXPENSE_CATS.map(c => [c, TAX_GROUPS[c] || 'Other / Review Required']) });

  // R: Enterprise filter options (R1, R2:R17 = "All Enterprises" + 15 enterprises)
  data.push({ range: `${R}!R1`, values: [['Enterprise Filter']] });
  data.push({ range: `${R}!R2:R17`, values: [['All Enterprises'], ...ENTERPRISES.map(e => [e])] });

  // S: Period filter (S1, S2:S14 = "Full Year" + 12 months)
  data.push({ range: `${R}!S1`, values: [['Period']] });
  data.push({ range: `${R}!S2:S14`, values: [['Full Year'], ...MONTHS.map(m => [m])] });

  // T: Quarter filter (T1, T2:T6 = "Full Year" + Q1-Q4)
  data.push({ range: `${R}!T1`, values: [['Quarter Filter']] });
  data.push({ range: `${R}!T2:T6`, values: [['Full Year'],['Q1'],['Q2'],['Q3'],['Q4']] });

  await valuesBatchUpdate(id, data, 'ref-values');

  // Format reference data
  const reqs = [];
  const hdr = { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9, fontFamily: 'Arial' } } };
  // Format header row
  reqs.push({ repeatCell: { range: gridRange(REF,0,1,0,20), cell: hdr, fields: 'userEnteredFormat' } });
  // Col widths
  [[0,110],[2,150],[4,160],[6,130],[7,110],[8,110],[9,60],[11,100],[12,60],[13,50],[15,160],[16,180],[17,120],[18,100],[19,100]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  // Freeze row 1
  reqs.push({ updateSheetProperties: { properties: { sheetId: REF, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
