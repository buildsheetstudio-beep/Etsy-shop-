'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

// Row positions exported as constants (0-indexed internally, 1-indexed in formula strings)
const R = {
  householdType: { start: 4, end: 6 },     // A4:A5
  accountType:   { start: 8, end: 15 },    // A8:A14
  incomeType:    { start: 17, end: 25 },   // A17:A24
  expenseCat:    { start: 27, end: 45 },   // A27:A44
  scenario:      { start: 47, end: 53 },   // A47:C52
  milestoneCat:  { start: 55, end: 62 },   // A55:A61
  milestoneStatus:{ start: 64, end: 69 },  // A64:A68
  accountStatus: { start: 71, end: 75 },   // A71:A74
  incomeStatus:  { start: 77, end: 80 },   // A77:A79
  expenseStatus: { start: 82, end: 85 },   // A82:A84
  withdrawalMethod:{ start: 87, end: 93 }, // A87:A92
  riskProfile:   { start: 95, end: 99 },   // A95:A98
  priority:      { start: 101, end: 105 }, // A101:A104
  assumptions:   { start: 107, end: 115 }, // B107:B114
};

const SECTIONS = [
  {
    label: 'Household Type', row: 2, /* header at row 3 (0-idx=2), data at 4-5 */
    data: [['Individual'],['Couple']],
  },
  {
    label: 'Account Types', row: 6,
    data: [
      ['Pre-Tax 401(k)'],['Pre-Tax 403(b)'],['Pre-Tax IRA'],
      ['Roth 401(k)'],['Roth IRA'],['Taxable Brokerage'],['Tax-Deferred Annuity'],
    ],
  },
  {
    label: 'Income Types', row: 15,
    data: [
      ['Social Security'],['Pension'],['Part-time Work'],
      ['Rental Income'],['Dividends / Interest'],['Annuity'],
      ['Inheritance / Gift'],['Other'],
    ],
  },
  {
    label: 'Expense Categories', row: 25,
    data: [
      ['Essential Living'],['Housing / Rent / Mortgage'],['Healthcare / Medical'],
      ['Transportation'],['Food / Groceries'],['Utilities'],['Insurance'],
      ['Entertainment'],['Travel & Leisure'],['Dining Out'],['Hobbies'],
      ['Gifts & Charitable'],['Personal Care'],['Home Repair / Maintenance'],
      ['Taxes'],['Subscriptions'],['One-Time / Large Purchase'],['Other'],
    ],
  },
  {
    label: 'Scenarios', row: 45,
    data: [
      ['SCN-01','Baseline','#7B9FC2','No changes; current plan'],
      ['SCN-02','Earlier Retirement','#8FB6A0','Retire 2 years sooner'],
      ['SCN-03','Later Retirement','#A89BC2','Retire 2 years later'],
      ['SCN-04','Higher Inflation','#D4A84B','Inflation at 5% instead of 3%'],
      ['SCN-05','Lower Returns','#B35C54','Portfolio returns 2% lower'],
      ['SCN-06','Reduced Spending','#7EB8B5','Expenses cut by 15% in retirement'],
    ],
  },
  {
    label: 'Milestone Categories', row: 53,
    data: [
      ['Financial'],['Health & Wellness'],['Housing'],
      ['Social & Family'],['Travel & Leisure'],['Legal & Estate'],['Personal Growth'],
    ],
  },
  {
    label: 'Milestone Statuses', row: 62,
    data: [['Not Started'],['In Progress'],['Completed'],['On Hold'],['Skipped']],
  },
  {
    label: 'Account Statuses', row: 69,
    data: [['Active'],['Paused'],['Closed'],['Rolled Over']],
  },
  {
    label: 'Income Statuses', row: 75,
    data: [['Active'],['Projected'],['Ended']],
  },
  {
    label: 'Expense Statuses', row: 81,
    data: [['Active'],['Projected'],['Eliminated']],
  },
  {
    label: 'Withdrawal Methods', row: 85,
    data: [
      ['4% Rule (Safe Withdrawal)'],
      ['Required Minimum Distribution (RMD)'],
      ['Bucket Strategy'],
      ['Dynamic Withdrawal'],
      ['Fixed Dollar Amount'],
      ['Fixed Percentage'],
    ],
  },
  {
    label: 'Risk Profiles', row: 93,
    data: [['Conservative'],['Moderate-Conservative'],['Moderate'],['Aggressive']],
  },
  {
    label: 'Priority Levels', row: 99,
    data: [['High'],['Medium'],['Low'],['None']],
  },
];

// Global Assumptions table (row 105 = header, 106+ = data)
const ASSUMPTIONS = [
  ['Inflation Rate (%)', 3.0],
  ['Healthcare Inflation (%)', 5.5],
  ['Expected Return — Conservative (%)', 5.0],
  ['Expected Return — Moderate (%)', 7.0],
  ['Expected Return — Aggressive (%)', 9.0],
  ['Life Expectancy (Years)', 90],
  ['Social Security COLA (%)', 2.5],
  ['Safe Withdrawal Rate (%)', 4.0],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // Title
  vals.push({ range: `${S}!A1`, values: [['ULTIMATE RETIREMENT PLANNER — REFERENCE DATA (DO NOT EDIT)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,5), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Column widths
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 240 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 200 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 100 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 200 }, fields: 'pixelSize' }});

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,900,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Sections
  for (const sec of SECTIONS) {
    const hRow = sec.row; // 0-indexed header row
    const dStart = hRow + 1;

    // Section header
    vals.push({ range: `${S}!A${hRow+1}`, values: [[sec.label]] });
    fmt.push({ mergeCells: { range: gridRange(SID, hRow, hRow+1, 0, 5), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID, hRow, hRow+1, 0, 5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: hRow, endIndex: hRow+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});

    // Data rows
    vals.push({ range: `${S}!A${dStart+1}`, values: sec.data });
    fmt.push({ repeatCell: { range: gridRange(SID, dStart, dStart + sec.data.length, 0, 5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
  }

  // Global Assumptions section
  const assHdr = 104; // 0-indexed
  vals.push({ range: `${S}!A${assHdr+1}`, values: [['Global Assumptions']] });
  fmt.push({ mergeCells: { range: gridRange(SID, assHdr, assHdr+1, 0, 5), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID, assHdr, assHdr+1, 0, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: assHdr, endIndex: assHdr+1 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const assStart = assHdr + 1;
  vals.push({ range: `${S}!A${assStart+1}`, values: ASSUMPTIONS });
  fmt.push({ repeatCell: { range: gridRange(SID, assStart, assStart + ASSUMPTIONS.length, 0, 2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
    textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
    borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
  }}, fields: 'userEnteredFormat' }});
  // Value column bold
  fmt.push({ repeatCell: { range: gridRange(SID, assStart, assStart + ASSUMPTIONS.length, 1, 2), cell: { userEnteredFormat: {
    textFormat: { bold: true, fontSize: 9 },
    horizontalAlignment: 'RIGHT',
  }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment' }});

  await valuesBatchUpdate(id, vals, '02-reference values');
  await batchUpdate(id, fmt, '02-reference format');

  console.log('✅ Reference Data done.');
  console.log('  Key ranges for validation:');
  console.log('  Household Type:     A4:A5');
  console.log('  Account Types:      A8:A14');
  console.log('  Income Types:       A17:A24');
  console.log('  Expense Categories: A27:A44');
  console.log('  Scenarios:          A47:D52');
  console.log('  Milestone Categories:A55:A61');
  console.log('  Milestone Statuses: A64:A68');
  console.log('  Account Statuses:   A71:A74');
  console.log('  Income Statuses:    A77:A79');
  console.log('  Expense Statuses:   A82:A84');
  console.log('  Withdrawal Methods: A87:A92');
  console.log('  Risk Profiles:      A95:A98');
  console.log('  Priority Levels:    A101:A104');
  console.log('  Global Assumptions: B106:B113');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
