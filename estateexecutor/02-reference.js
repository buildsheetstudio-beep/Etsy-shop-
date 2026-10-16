'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];

(async () => {
  const reqs = [];

  // Hide tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true },
    fields: 'hidden',
  }});

  // Freeze 1 row
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, gridProperties: { frozenRowCount: 1 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  await batchUpdate(id, reqs, 'ref-props');

  // ── Reference data ────────────────────────────────────────────────────────
  // Column layout (0-indexed):
  // A(0):  Asset Types
  // B(1):  Debt Types
  // C(2):  Task Phase
  // D(3):  Task Status
  // E(4):  Asset Status
  // F(5):  Debt Status
  // G(6):  Beneficiary Type
  // H(7):  Distribution Type
  // I(8):  Account Type (Digital)
  // J(9):  Account Status (Digital)
  // K(10): Comm Type
  // L(11): Comm Method
  // M(12): Service Type (Professional)
  // N(13): Fee Status
  // O(14): Doc Category
  // P(15): Doc Status
  // Q(16): Wellbeing (unused here — skip)
  // R(17): Note Topic
  // S(18): Note Status
  // T(19): Contact Role
  // U(20): Priority
  // V(21): Timeline Phase
  // W(22): Dist Status

  const data = [
    // Headers row 0
    {
      range: `'📋 Reference Data'!A1:W1`,
      values: [['Asset Types','Debt Types','Task Phase','Task Status','Asset Status','Debt Status','Beneficiary Type','Distribution Type','Account Type (Digital)','Account Status','Comm Type','Comm Method','Service Type','Fee Status','Doc Category','Doc Status','Note Topic','Note Status','Contact Role','Priority','Timeline Phase','Dist Status','']],
    },
    // A: Asset Types (rows 2-18)
    {
      range: `'📋 Reference Data'!A2:A18`,
      values: [
        ['Real Estate'],['Bank Account'],['Investment Portfolio'],['Retirement Account'],['Life Insurance'],
        ['Vehicle'],['Personal Property'],['Business Interest'],['Cryptocurrency'],['Superannuation'],
        ['Bonds'],['Trust Asset'],['Shares/Stocks'],['Collectibles'],['Jewellery'],['Other Asset'],[''],
      ],
    },
    // B: Debt Types
    {
      range: `'📋 Reference Data'!B2:B12`,
      values: [
        ['Mortgage'],['Credit Card'],['Personal Loan'],['Tax Debt'],['Medical Bill'],
        ['Utility Bill'],['Legal Fees'],['Funeral Expenses'],['Other Debt'],[''],[''],
      ],
    },
    // C: Task Phase
    {
      range: `'📋 Reference Data'!C2:C8`,
      values: [
        ['Phase 1: Immediate'],['Phase 2: First Month'],['Phase 3: 1-3 Months'],
        ['Phase 4: 3-6 Months'],['Phase 5: 6-12 Months'],['Phase 6: Finalisation'],[''],
      ],
    },
    // D: Task Status
    {
      range: `'📋 Reference Data'!D2:D7`,
      values: [['Not Started'],['In Progress'],['Completed'],['On Hold'],['Delegated'],['N/A']],
    },
    // E: Asset Status
    {
      range: `'📋 Reference Data'!E2:E8`,
      values: [['Active'],['Pending Valuation'],['Pending Transfer'],['Transferred'],['Sold'],['Closed'],['Disputed']],
    },
    // F: Debt Status
    {
      range: `'📋 Reference Data'!F2:F7`,
      values: [['Outstanding'],['In Dispute'],['Payment Plan'],['Partially Paid'],['Paid in Full'],['Written Off']],
    },
    // G: Beneficiary Type
    {
      range: `'📋 Reference Data'!G2:G7`,
      values: [['Individual'],['Trust'],['Charity'],['Minor (Trust Held)'],['Organisation'],['Other']],
    },
    // H: Distribution Type
    {
      range: `'📋 Reference Data'!H2:H8`,
      values: [['Specific Bequest'],['Residuary'],['Pecuniary Bequest'],['Trust Distribution'],['Life Interest'],['Intestacy Share'],['Other']],
    },
    // I: Account Type (Digital)
    {
      range: `'📋 Reference Data'!I2:I12`,
      values: [
        ['Email'],['Social Media'],['Banking/Finance'],['Subscription'],['Cloud Storage'],
        ['Shopping'],['Healthcare'],['Government/Legal'],['Investment'],['Insurance'],['Other'],
      ],
    },
    // J: Account Status (Digital)
    {
      range: `'📋 Reference Data'!J2:J7`,
      values: [['Active'],['To Be Closed'],['Closed'],['Memorialised'],['Transferred'],['Unknown']],
    },
    // K: Comm Type
    {
      range: `'📋 Reference Data'!K2:K8`,
      values: [['Initial Notification'],['Update'],['Distribution Notice'],['Request'],['Response'],['Meeting'],['Other']],
    },
    // L: Comm Method
    {
      range: `'📋 Reference Data'!L2:L7`,
      values: [['Email'],['Phone'],['Letter'],['In Person'],['Solicitor'],['Video Call']],
    },
    // M: Service Type (Professional)
    {
      range: `'📋 Reference Data'!M2:M10`,
      values: [['Solicitor/Lawyer'],['Accountant'],['Financial Adviser'],['Valuer'],['Real Estate Agent'],['Funeral Director'],['Counsellor'],['Other Professional'],['']],
    },
    // N: Fee Status
    {
      range: `'📋 Reference Data'!N2:N6`,
      values: [['Quoted'],['Invoiced'],['Paid'],['Disputed'],['Waived']],
    },
    // O: Doc Category
    {
      range: `'📋 Reference Data'!O2:O12`,
      values: [
        ['Will & Probate'],['Identity'],['Property'],['Financial'],['Tax'],
        ['Insurance'],['Legal'],['Correspondence'],['Funeral'],['Medical'],['Other'],
      ],
    },
    // P: Doc Status
    {
      range: `'📋 Reference Data'!P2:P7`,
      values: [['Required'],['Located'],['Certified Copy'],['Submitted'],['Returned'],['Not Applicable']],
    },
    // Q: Note Topic
    {
      range: `'📋 Reference Data'!Q2:Q10`,
      values: [['General'],['Legal'],['Financial'],['Family'],['Property'],['Tax'],['Insurance'],['Funeral'],['Other']],
    },
    // R: Note Status
    {
      range: `'📋 Reference Data'!R2:R5`,
      values: [['Open'],['In Progress'],['Resolved'],['Archived']],
    },
    // S: Contact Role
    {
      range: `'📋 Reference Data'!S2:S12`,
      values: [
        ['Beneficiary'],['Solicitor'],['Accountant'],['Financial Adviser'],['Bank Contact'],
        ['Insurance Contact'],['Property Agent'],['Funeral Director'],['Trustee'],['Family'],['Other'],
      ],
    },
    // T: Priority
    {
      range: `'📋 Reference Data'!T2:T4`,
      values: [['High'],['Medium'],['Low']],
    },
    // U: Timeline Phase
    {
      range: `'📋 Reference Data'!U2:U8`,
      values: [
        ['Phase 1: Immediate'],['Phase 2: First Month'],['Phase 3: 1-3 Months'],
        ['Phase 4: 3-6 Months'],['Phase 5: 6-12 Months'],['Phase 6: Finalisation'],[''],
      ],
    },
    // V: Dist Status
    {
      range: `'📋 Reference Data'!V2:V6`,
      values: [['Pending'],['In Progress'],['Partial'],['Complete'],['On Hold']],
    },
  ];

  await valuesBatchUpdate(id, data, 'ref');
  console.log('Reference data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
