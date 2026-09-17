'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

// ── Dropdown list definitions ────────────────────────────────────────────────
// Stored in rows 1-20, cols A-P
// Row 1 = section header, Row 2 = column headers, Rows 3+ = values
// Each list occupies one column; max list length = 10 items

const LISTS = {
  A: { header: 'Family Roles',         items: ['Parent','Guardian','Grandparent','Relative','Beneficiary','Contributor','Other'] },
  B: { header: 'Beneficiary Status',   items: ['Planning','Saving','Near Enrollment','Enrolled','Graduated','Paused','Archived'] },
  C: { header: 'Account Types',        items: ['529 Savings Plan','Prepaid Tuition Plan','Coverdell ESA','Custodial Account','Brokerage / Investment Account','Savings Account','CD','Other Education Savings'] },
  D: { header: 'Account Ownership',    items: ['Parent / Guardian','Grandparent','Beneficiary','Relative','Joint / Household','Other'] },
  E: { header: 'Contribution Sources', items: ['Parent / Guardian','Grandparent','Relative','Beneficiary','Gift','Payroll / Automatic Deposit','Bonus','Tax Refund','Other'] },
  F: { header: 'Transaction Types',    items: ['Contribution','Withdrawal','Adjustment Increase','Adjustment Decrease','Transfer In','Transfer Out','Other'] },
  G: { header: 'Contrib. Frequencies', items: ['Weekly','Biweekly','Semi-Monthly','Monthly','Quarterly','Annual','Irregular / Manual'] },
  H: { header: 'College Types',        items: ['In-State Public','Out-of-State Public','Private Nonprofit','Private For-Profit','Community College','Trade / Technical School','Custom School'] },
  I: { header: 'Cost Categories',      items: ['Tuition','Mandatory Fees','Housing','Meal Plan / Food','Books & Supplies','Transportation','Personal Expenses','Technology','Insurance','Other'] },
  J: { header: 'Growth Scenarios',     items: ['Conservative','Base','Higher Growth'] },
  K: { header: 'Goal Statuses',        items: ['Not Started','Active','Ahead of Plan','On Track','Behind Plan','Reached','Paused','Archived'] },
  L: { header: 'Milestone Types',      items: ['10% Funded','25% Funded','50% Funded','75% Funded','90% Funded','100% Funded','Custom Amount','Custom Percentage','Enrollment Date','Other'] },
  M: { header: 'Milestone Statuses',   items: ['Not Started','In Progress','Achieved','Delayed','Reassess'] },
  N: { header: 'Yes / No',             items: ['Yes','No'] },
  O: { header: 'Account Status',       items: ['Active','Paused','Closed','Transferred','Archived'] },
  P: { header: 'Planner Status',       items: ['Ahead of Plan','On Track','Behind Plan','Goal Reached','Insufficient Data'] },
};

// Benchmark college cost table starts at row 23 (0-indexed = 22)
// Federal 529 reference starts at row 33 (0-indexed = 32)
// State 529 benefit reference starts at row 43 (0-indexed = 42)

// Benchmark data: "Illustrative planning assumption — verify and update before relying on this figure."
const BENCHMARKS = [
  // [Type, Tuition, Fees, Housing, Food, Books, Transport, Personal, TuitionInfl, OtherInfl, LastVerified, Notes]
  ['In-State Public',       11500, 1100, 10200, 5100, 1150, 1800, 2000, 0.04,  0.03,  2025, 'Illustrative planning assumption — verify and update before relying on this figure.'],
  ['Out-of-State Public',   27000, 1400, 12500, 5500, 1200, 1500, 2200, 0.04,  0.03,  2025, 'Illustrative planning assumption — verify and update before relying on this figure.'],
  ['Private Nonprofit',     39500, 1700, 13800, 6300, 1350, 1000, 2500, 0.045, 0.03,  2025, 'Illustrative planning assumption — verify and update before relying on this figure.'],
  ['Private For-Profit',    21500, 1100, 11000, 5000, 1200, 1500, 2000, 0.035, 0.025, 2025, 'Illustrative planning assumption — verify and update before relying on this figure.'],
  ['Community College',      4200,  580,  8500, 4200,  950, 2200, 1600, 0.03,  0.025, 2025, 'Illustrative planning assumption — verify and update before relying on this figure.'],
  ['Trade / Technical School',7800, 750,  9200, 4600, 1450, 1600, 1800, 0.035, 0.025, 2025, 'Illustrative planning assumption — verify and update before relying on this figure.'],
  ['Custom School',              0,   0,     0,    0,    0,    0,    0,  0.04,  0.03,  2025, 'User-defined — enter school-specific values in the College Cost Estimator tab.'],
];

// Federal 529 reference (year-keyed)
const FEDERAL = [
  // [Year, Item, Amount, Source/Note, LastVerified]
  [2025, 'Annual Gift Tax Exclusion',                                '$18,000 per donor per recipient',  'IRS Rev. Proc. 2023-34 — VERIFY CURRENT VALUE',          '2025-01-01'],
  [2025, '529 5-Year Election (Superfunding) — Max per beneficiary', '$90,000 (5 × annual exclusion)',   'IRS § 529(c)(2)(B) — VERIFY CURRENT VALUE',              '2025-01-01'],
  [2025, 'K-12 Qualified Expense Annual Limit (per § 529)',          '$10,000',                          'Tax Cuts and Jobs Act § 529(c)(7)',                        '2025-01-01'],
  [2025, 'Student Loan Repayment Lifetime Limit (per § 529)',        '$10,000',                          'SECURE Act § 529(c)(9)',                                  '2025-01-01'],
  [2025, '529-to-Roth IRA Rollover Annual Limit (SECURE 2.0)',       'Up to Roth IRA annual limit',      'SECURE 2.0 Act — conditions and limits apply; VERIFY',    '2025-01-01'],
  [2024, 'Annual Gift Tax Exclusion',                                '$18,000 per donor per recipient',  'IRS Rev. Proc. 2023-34',                                  '2024-01-01'],
  [2023, 'Annual Gift Tax Exclusion',                                '$17,000 per donor per recipient',  'IRS Rev. Proc. 2022-38',                                  '2023-01-01'],
];

// State 529 benefit reference (manual — verify with official sources)
const STATE_BENEFITS = [
  // [State, PlanName, TaxDeduction, TaxCredit, MaxDeductionCredit, EligibleContrib, ResidencyRequired, BenefitYear, SourceURL, LastVerified, Notes]
  ['Washington', 'GET (Prepaid) / DreamAhead (Savings)', 'No', 'No', 'N/A — no state income tax', 'Any', 'No',  2025, 'https://www.wacollegesavings.com', '2025-01-01', 'No state income tax; no deduction/credit applicable.'],
  ['New York',   'NY 529 Direct Plan',                   'Yes','No', '$5,000 single / $10,000 married', 'Account Owner', 'Yes', 2025, 'https://www.nysaves.org', '2025-01-01', 'VERIFY current limits. Recapture rules apply on non-qualified withdrawals.'],
  ['Ohio',       'CollegeAdvantage',                     'Yes','No', 'Unlimited (full contribution deduction)', 'Account Owner', 'No', 2025, 'https://www.collegeadvantage.com', '2025-01-01', 'No cap on deduction amount — VERIFY current rules.'],
  ['Utah',       'my529',                                'Yes','No', 'Up to ~$2,340 per beneficiary per taxpayer', 'Account Owner', 'Yes', 2025, 'https://my529.org', '2025-01-01', 'VERIFY current limits. Utah also offers a credit option.'],
  ['California', 'ScholarShare 529',                     'No', 'No', 'N/A',                           'N/A',           'N/A', 2025, 'https://www.scholarshare529.com', '2025-01-01', 'California does not offer a state income tax deduction or credit for 529 contributions.'],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // ── Background wash ───────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID,0,300,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Section 1: Dropdown Lists (rows 1-21) ─────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['DROPDOWN REFERENCE LISTS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Column headers and values for each dropdown list
  const colLetters = Object.keys(LISTS);
  for (const col of colLetters) {
    const ci = col.charCodeAt(0) - 65;
    const list = LISTS[col];
    vals.push({ range: `${S}!${col}2`, values: [[list.header]] });
    fmt.push({ repeatCell: { range: gridRange(SID,1,2,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});
    list.items.forEach((item, ii) => {
      vals.push({ range: `${S}!${col}${3+ii}`, values: [[item]] });
    });
    // Color the values cells light
    fmt.push({ repeatCell: { range: gridRange(SID,2,2+list.items.length,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.white), textFormat: { fontSize: 8, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat' }});
  }

  // ── Section 2: Benchmark College Costs (row 22 = 0-indexed) ───────────────
  vals.push({ range: `${S}!A22`, values: [['BENCHMARK COLLEGE COST PLANNING ASSUMPTIONS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,21,22,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,21,22,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 21, endIndex: 22 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const benchHdr = ['College Type','Current Annual Tuition','Current Annual Fees','Current Annual Housing','Current Annual Food','Current Annual Books','Current Annual Transportation','Current Annual Personal / Other','Default Tuition Inflation %','Default Other-Cost Inflation %','Last Verified Year','Notes'];
  vals.push({ range: `${S}!A23`, values: [benchHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID,22,23,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.aubergTint), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial' },
    wrapStrategy: 'WRAP', horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 22, endIndex: 23 }, properties: { pixelSize: 34 }, fields: 'pixelSize' }});

  BENCHMARKS.forEach((bRow, bi) => {
    vals.push({ range: `${S}!A${24+bi}`, values: [bRow] });
    fmt.push({ repeatCell: { range: gridRange(SID,23+bi,24+bi,0,12), cell: { userEnteredFormat: {
      backgroundColor: hex(bi % 2 === 0 ? C.white : C.altRow), textFormat: { fontSize: 8, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat' }});
    // Format currency cols B-H (indices 1-7)
    [1,2,3,4,5,6,7].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,23+bi,24+bi,ci,ci+1), cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
    });
    // Format percentage cols I-J (indices 8-9)
    [8,9].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,23+bi,24+bi,ci,ci+1), cell: { userEnteredFormat: {
        numberFormat: { type: 'PERCENT', pattern: '0.0%' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
    });
  });

  // ── Section 3: Federal 529 Reference (row 32+) ───────────────────────────
  vals.push({ range: `${S}!A32`, values: [['FEDERAL 529 REFERENCE TABLE — VERIFY FIGURES BEFORE RELYING ON THEM']] });
  fmt.push({ mergeCells: { range: gridRange(SID,31,32,0,5), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,31,32,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 31, endIndex: 32 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A33`, values: [['Year','Reference Item','Amount / Limit','Source / Verification Note','Last Verified Date']] });
  fmt.push({ repeatCell: { range: gridRange(SID,32,33,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.copperTint), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  FEDERAL.forEach((row, ri) => {
    vals.push({ range: `${S}!A${34+ri}`, values: [row] });
    fmt.push({ repeatCell: { range: gridRange(SID,33+ri,34+ri,0,5), cell: { userEnteredFormat: {
      backgroundColor: hex(ri % 2 === 0 ? C.white : C.altRow), textFormat: { fontSize: 8, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat' }});
  });

  // ── Section 4: State 529 Benefit Reference (row 43+) ─────────────────────
  const stateStart = 43;
  vals.push({ range: `${S}!A${stateStart}`, values: [['STATE 529 BENEFIT REFERENCE — VERIFY WITH OFFICIAL SOURCES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,stateStart-1,stateStart,0,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,stateStart-1,stateStart,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.eucalTint), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: stateStart-1, endIndex: stateStart }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const stateDisclaimer = 'State 529 deductions, credits, recapture rules, eligibility requirements, and limits vary and can change. Verify current rules with the relevant state program or tax authority.';
  vals.push({ range: `${S}!A${stateStart+1}`, values: [[stateDisclaimer]] });
  fmt.push({ mergeCells: { range: gridRange(SID,stateStart,stateStart+1,0,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,stateStart,stateStart+1,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.warning), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.text), italic: true, fontFamily: 'Arial' },
    wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: stateStart, endIndex: stateStart+1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  const stateHdr = ['State','State Plan Name','State Tax Deduction?','State Tax Credit?','Max Deduction / Credit','Eligible Contributor','Residency Required?','Benefit Year','Source / Official URL','Last Verified Date','Notes'];
  vals.push({ range: `${S}!A${stateStart+2}`, values: [stateHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID,stateStart+1,stateStart+2,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.eucalTint), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial' },
    wrapStrategy: 'WRAP', horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  STATE_BENEFITS.forEach((row, ri) => {
    vals.push({ range: `${S}!A${stateStart+3+ri}`, values: [row] });
    fmt.push({ repeatCell: { range: gridRange(SID,stateStart+2+ri,stateStart+3+ri,0,11), cell: { userEnteredFormat: {
      backgroundColor: hex(ri % 2 === 0 ? C.white : C.altRow), textFormat: { fontSize: 8, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat' }});
  });

  // Column widths for reference data
  [180,120,90,100,100,100,120,120,120,90,80,300].forEach((px,ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  await valuesBatchUpdate(id, vals, '02-reference values');
  await batchUpdate(id, fmt, '02-reference format');

  // Hide Reference Data sheet now that other sheets exist
  await batchUpdate(id, [{
    updateSheetProperties: {
      properties: { sheetId: SID, hidden: true },
      fields: 'hidden'
    }
  }], '02-reference hide');
  console.log('✅  Reference Data done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
