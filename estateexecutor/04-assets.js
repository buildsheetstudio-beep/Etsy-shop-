'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const AD = sheetMap['💼 Assets & Debts'];

// 12 assets, 5 debts
// Columns: A=# B=Description C=Type D=Location/Institution E=Date of Death Value F=Current Value G=Status H=Notes I=Est. Realisation J=Owner K=Acct/Ref
const ASSETS = [
  [1, 'Family Home — 42 Wattle Street, Mosman NSW', 'Real Estate', 'Heritage Bank Mortgage', '1450000', '1450000', 'Pending Transfer', 'Listed for sale', '2025-09-01', 'Sole', 'CT/1234567'],
  [2, 'Investment Property — 8 Harbour View, Manly NSW', 'Real Estate', 'Tenant occupied', '920000', '920000', 'Pending Transfer', 'Rental income ongoing', '2025-10-01', 'Sole', 'CT/7654321'],
  [3, 'NAB Savings Account', 'Bank Account', 'NAB Mosman', '48520', '49100', 'Active', 'Interest accruing', '2025-06-01', 'Sole', 'BSB 082-001 A/C 231456789'],
  [4, 'CommBank Transaction Account', 'Bank Account', 'Commonwealth Bank', '12340', '12340', 'Pending Transfer', 'Funds to estate account', '2025-06-01', 'Sole', 'BSB 062-800 A/C 10234567'],
  [5, 'Westpac Term Deposit', 'Bank Account', 'Westpac', '75000', '75000', 'Pending Transfer', 'Maturity 2025-08-15', '2025-08-15', 'Sole', 'A/C 034-789-000123'],
  [6, 'Colonial First State Investment Portfolio', 'Investment Portfolio', 'Colonial First State', '340000', '345200', 'Pending Transfer', 'ASX equities & managed funds', '2025-09-01', 'Sole', 'CFS-2024-78901'],
  [7, 'AustralianSuper — Death Benefit (paid)', 'Superannuation', 'AustralianSuper', '385000', '385000', 'Transferred', 'Paid to estate 2025-05-28', '2025-05-28', 'Sole', 'MBR-00124567'],
  [8, 'AIA Life Insurance Policy', 'Life Insurance', 'AIA Australia', '500000', '500000', 'Active', 'Claim lodged 2025-05-05', '2025-07-01', 'Sole', 'POL-0099887'],
  [9, 'Toyota Camry 2021', 'Vehicle', 'Garage — family home', '28500', '26000', 'Active', 'Agreed sale to family member', '2025-07-15', 'Sole', 'Rego: ABC123'],
  [10, 'Jewellery & Watches Collection', 'Personal Property', 'Family home (secured)', '38000', '38000', 'Pending Valuation', 'Professional valuation required', '2025-08-01', 'Sole', 'Insured under home contents'],
  [11, 'Antique Furniture & Art', 'Collectibles', 'Family home', '22500', '22500', 'Pending Valuation', 'Contents valuer booked', '2025-08-01', 'Sole', 'See contents schedule'],
  [12, 'Share Portfolio — ASX Direct', 'Shares/Stocks', 'CommSec', '48280', '49200', 'Pending Transfer', 'Transfer to estate account', '2025-09-01', 'Sole', 'CDA-000-456789'],
];

// 5 debts
// Columns: A=# B=Description C=Type D=Creditor E=Balance Owing F=Interest Rate G=Status H=Notes I=Due/Expiry
const DEBTS = [
  [1, 'Heritage Bank Home Mortgage — 42 Wattle St', 'Mortgage', 'Heritage Bank', '185000', '4.95%', 'Outstanding', 'Will be discharged on property sale', '2025-09-01'],
  [2, 'ANZ Credit Card', 'Credit Card', 'ANZ Bank', '4200', '19.99%', 'Outstanding', 'To be paid from estate account', '2025-07-31'],
  [3, 'ATO — Final Tax Assessment 2024-25', 'Tax Debt', 'Australian Taxation Office', '18400', '0%', 'In Dispute', 'Accountant liaising with ATO', '2025-10-31'],
  [4, 'St Vincent\'s Hospital — Medical Bills', 'Medical Bill', 'St Vincent\'s Hospital Sydney', '6800', '0%', 'Outstanding', 'Insurance claim pending', '2025-08-31'],
  [5, 'Funeral Expenses — Southern Cross Funerals', 'Funeral Expenses', 'Southern Cross Funerals', '12400', '0%', 'Paid in Full', 'Paid from estate account 2025-05-02', '2025-05-02'],
];

(async () => {
  const reqs = [];

  // Freeze: 4 rows (title + disclaimer + section header + col headers), no col freeze
  // Full-width section merges conflict with column freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: AD, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: AD, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AD, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 44 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AD, dimension: 'ROWS', startIndex: 2, endIndex: 4 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AD, dimension: 'ROWS', startIndex: 4, endIndex: 60 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Column widths: A=40 B=300 C=130 D=200 E=120 F=120 G=120 H=200 I=110 J=70 K=170
  const colW = [40, 300, 130, 200, 120, 120, 120, 200, 110, 70, 170];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: AD, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // ── Row 0: Title (full width A:K) ─────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(AD, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(AD, 0, 1, 0, 11), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 1: Disclaimer — col A empty, merge B:K ────────────────────────────
  reqs.push({ repeatCell: { range: gridRange(AD, 1, 2, 0, 1), cell: {
    userEnteredFormat: { backgroundColor: hex(C.disclaimerBg) },
  }, fields: 'userEnteredFormat' }});
  reqs.push({ mergeCells: { range: gridRange(AD, 1, 2, 1, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(AD, 1, 2, 1, 11), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.disclaimerBg),
      textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 2: ASSETS section header ──────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(AD, 2, 3, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(AD, 2, 3, 0, 11), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      padding: { left: 10 },
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 3: Column headers ─────────────────────────────────────────────────
  reqs.push({ repeatCell: { range: gridRange(AD, 3, 4, 0, 11), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // ── Asset rows 4-15 (12 assets) ───────────────────────────────────────────
  for (let r = 0; r < 12; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(AD, r+4, r+5, 0, 11), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // ── Row 16: Assets totals ─────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(AD, 16, 17, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(AD, 16, 17, 0, 11), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.paleSage),
      textFormat: { foregroundColor: hex(C.mutedSage), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { top: { style: 'SOLID', color: hex(C.mutedSage), width: 2 }, bottom: { style: 'SOLID', color: hex(C.mutedSage), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 17: spacer ────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: AD, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 16 }, fields: 'pixelSize' }});

  // ── Row 18: DEBTS section header ──────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(AD, 18, 19, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(AD, 18, 19, 0, 11), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepRust),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      padding: { left: 10 },
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 19: Debt column headers ───────────────────────────────────────────
  reqs.push({ repeatCell: { range: gridRange(AD, 19, 20, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex('#9B5A5A'),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepRust), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // ── Debt rows 20-24 (5 debts) ─────────────────────────────────────────────
  for (let r = 0; r < 5; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(AD, r+20, r+21, 0, 9), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // ── Row 25: Debts total ───────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(AD, 25, 26, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(AD, 25, 26, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.errorBg),
      textFormat: { foregroundColor: hex(C.deepRust), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { top: { style: 'SOLID', color: hex(C.deepRust), width: 2 }, bottom: { style: 'SOLID', color: hex(C.deepRust), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 26: spacer ────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: AD, dimension: 'ROWS', startIndex: 26, endIndex: 27 }, properties: { pixelSize: 16 }, fields: 'pixelSize' }});

  // ── Row 27: NET ESTATE VALUE ──────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(AD, 27, 28, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(AD, 27, 28, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.paleSage),
      textFormat: { foregroundColor: hex(C.mutedSage), bold: true, fontSize: 14, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: {
        top: { style: 'SOLID', color: hex(C.mutedSage), width: 3 },
        bottom: { style: 'SOLID', color: hex(C.mutedSage), width: 3 },
      },
    },
  }, fields: 'userEnteredFormat' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AD, dimension: 'ROWS', startIndex: 27, endIndex: 28 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Number formats — currency for value columns
  reqs.push({ repeatCell: { range: gridRange(AD, 4, 16, 4, 7), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(AD, 20, 25, 4, 5), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});
  // Totals rows
  reqs.push({ repeatCell: { range: gridRange(AD, 16, 17, 4, 7), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(AD, 25, 26, 4, 5), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(AD, 27, 28, 4, 9), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  // Borders on data areas
  reqs.push({ updateBorders: { range: gridRange(AD, 3, 17, 0, 11),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});
  reqs.push({ updateBorders: { range: gridRange(AD, 19, 26, 0, 9),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  // Center align: A, C, G, I, J columns
  [0, 2, 6, 8, 9].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(AD, 4, 16, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));
  [0, 2, 5, 6, 7, 8].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(AD, 20, 25, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Wrap description col
  reqs.push({ repeatCell: { range: gridRange(AD, 4, 16, 1, 2), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat.wrapStrategy' }});
  reqs.push({ repeatCell: { range: gridRange(AD, 20, 25, 1, 2), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat.wrapStrategy' }});

  await batchUpdate(id, reqs, 'assets-fmt');

  // ── Values ────────────────────────────────────────────────────────────────
  const vals = [];

  vals.push({ range: `'💼 Assets & Debts'!A1`, values: [['💼 ASSETS & DEBTS — Estate of Margaret Eleanor Williams']] });
  vals.push({
    range: `'💼 Assets & Debts'!B2`,
    values: [['IMPORTANT DISCLAIMER: The values shown are estimates based on information available at the date of death and may change. This sheet does not constitute a formal estate account or legal valuation. Engage a qualified professional for formal valuations and legal advice.']],
  });
  vals.push({ range: `'💼 Assets & Debts'!A3`, values: [['💎 ASSETS']] });
  vals.push({
    range: `'💼 Assets & Debts'!A4:K4`,
    values: [['#', 'Description', 'Type', 'Location / Institution', 'DoD Value ($)', 'Current Value ($)', 'Status', 'Notes', 'Est. Realisation Date', 'Owner', 'Acct / Ref No.']],
  });

  // Asset data
  const assetRows = ASSETS.map(a => [a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9], a[10]]);
  vals.push({ range: `'💼 Assets & Debts'!A5`, values: assetRows });

  // Assets totals row
  vals.push({
    range: `'💼 Assets & Debts'!A17:K17`,
    values: [['TOTAL ASSETS', '', '', '', '=SUM(E5:E16)', '=SUM(F5:F16)', '', '', '', '', '']],
  });

  vals.push({ range: `'💼 Assets & Debts'!A19`, values: [['💳 DEBTS & LIABILITIES']] });
  vals.push({
    range: `'💼 Assets & Debts'!A20:I20`,
    values: [['#', 'Description', 'Type', 'Creditor', 'Balance Owing ($)', 'Interest Rate', 'Status', 'Notes', 'Due / Expiry Date']],
  });

  const debtRows = DEBTS.map(d => [d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8]]);
  vals.push({ range: `'💼 Assets & Debts'!A21`, values: debtRows });

  // Debts total
  vals.push({
    range: `'💼 Assets & Debts'!A26:I26`,
    values: [['TOTAL DEBTS', '', '', '', '=SUM(E21:E25)', '', '', '', '']],
  });

  // Net estate value
  vals.push({
    range: `'💼 Assets & Debts'!A28:I28`,
    values: [['NET ESTATE VALUE', '', '', '', '=F17-E26', '', '', '', '']],
  });

  await valuesBatchUpdate(id, vals, 'assets-vals');
  console.log('Assets & Debts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
