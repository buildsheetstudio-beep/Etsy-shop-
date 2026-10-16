'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// Col A: Participants (A2:A5)
// Col B: Expense Category (B2:B12)
// Col C: Payment Method (C2:C8)
// Col D: Claim/Reimbursement Status (D2:D7)
// Col E: Tax Year (E2:E4)
// Col F: HSA/FSA Lookup Key (F2:F10)  ← "2024 HSA Self-Only" format
// Col G: Annual Contribution Limit (G2:G10)
// Col H: (label: Account Type for readability)
// Col I: Insurance Claim Status (I2:I6)

const participants = ['Self', 'Spouse', 'Child 1', 'Child 2'];
const categories = [
  'Doctor Visits', 'Prescription Medications', 'Dental Care', 'Vision Care',
  'Emergency Services', 'Hospital Services', 'Therapy/Mental Health',
  'Medical Equipment', 'Lab/Diagnostic Tests', 'Preventive Care', 'Other',
];
const paymentMethods = ['HSA', 'FSA', 'Credit Card', 'Debit Card', 'Cash', 'Check', 'Insurance Direct Pay'];
const claimStatuses = ['Not Submitted', 'Submitted', 'Processing', 'Approved', 'Denied', 'Paid'];
const taxYears = [2024, 2025, 2026];

// HSA/FSA contribution limits (PLACEHOLDER — verify with IRS Publication 969 annually)
const hsaFsaLimits = [
  ['2024 HSA Self-Only', 4150],
  ['2024 HSA Family',   8300],
  ['2024 FSA',          3200],
  ['2025 HSA Self-Only', 4300],
  ['2025 HSA Family',   8550],
  ['2025 FSA',          3300],
  ['2026 HSA Self-Only', 4400],
  ['2026 HSA Family',   8750],
  ['2026 FSA',          3400],
];

// Insurance claim status (no "Not Submitted")
const insClaimStatuses = ['Submitted', 'Processing', 'Approved', 'Denied', 'Paid'];

(async () => {
  const reqs = [];

  // Header row styling
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Data rows background
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 20, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.softIvory) } },
      fields: 'userEnteredFormat',
    },
  });

  // HSA/FSA limit section header style
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 2, 5, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Warning note styling (row 11 in G)
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 11, 12, 5, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightAmber),
          textFormat: { foregroundColor: hex(C.darkText), italic: true, fontSize: 8 },
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  const widths = [100, 170, 140, 130, 80, 170, 120, 120, 120];
  for (const [i, px] of widths.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Currency format for G col
  reqs.push({ repeatCell: {
    range: gridRange(REF, 2, 11, 6, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  await batchUpdate(id, reqs, 'reference-format');

  // Values
  const data = [];

  // Headers row 1
  data.push({
    range: `'📋 Reference Data'!A1:I1`,
    values: [['Participant','Expense Category','Payment Method','Claim/Reimb. Status','Tax Year','HSA/FSA Lookup Key','Annual Limit','Account Type','Insurance Claim Status']],
  });

  // A2:A5 — Participants
  data.push({ range: `'📋 Reference Data'!A2:A5`, values: participants.map(p => [p]) });

  // B2:B12 — Categories
  data.push({ range: `'📋 Reference Data'!B2:B12`, values: categories.map(c => [c]) });

  // C2:C8 — Payment Methods
  data.push({ range: `'📋 Reference Data'!C2:C8`, values: paymentMethods.map(m => [m]) });

  // D2:D7 — Claim Statuses
  data.push({ range: `'📋 Reference Data'!D2:D7`, values: claimStatuses.map(s => [s]) });

  // E2:E4 — Tax Years
  data.push({ range: `'📋 Reference Data'!E2:E4`, values: taxYears.map(y => [y]) });

  // F1:H1 — HSA/FSA table header (styled above as mutedSage)
  data.push({
    range: `'📋 Reference Data'!F2:H2`,
    values: [['Lookup Key', 'Limit ($)', 'Account Type']],
  });

  // F3:G11 — HSA/FSA limit data
  data.push({
    range: `'📋 Reference Data'!F3`,
    values: hsaFsaLimits.map(([key, limit]) => [key, limit, key.replace(/^\d{4} /,'')]),
  });

  // G12 — Note about placeholder limits
  data.push({
    range: `'📋 Reference Data'!F12:H12`,
    values: [['⚠️ PLACEHOLDER LIMITS — Verify annually with IRS Publication 969 before real tax planning', '', '']],
  });

  // I2:I6 — Insurance Claim Status
  data.push({ range: `'📋 Reference Data'!I2:I6`, values: insClaimStatuses.map(s => [s]) });

  await valuesBatchUpdate(id, data, 'reference-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
