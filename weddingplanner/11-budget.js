'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BP = sheetMap['Budget & Payments'];
const S  = "'Budget & Payments'";

const PAYMENT_HEADERS = [
  '#', 'Date', 'Category', 'Vendor Name', 'Service Description',
  'Invoice #', 'Invoice Amount', 'Payment Amount', 'Payment Method',
  'Paid By', 'Payment Date', 'Receipt Ref', 'Running Total (auto)', 'Notes', 'Verified', 'Refund',
];

const budgetCategories = [
  ['Venue',                   13000, 12500],
  ['Catering',                18000, 9000],
  ['Photography',              4800, 2400],
  ['Videography',              3200, 1600],
  ['Florals',                  5500, 2750],
  ['Music / Entertainment',    2800, 1400],
  ['Hair & Makeup',            1800,  900],
  ['Attire & Accessories',     3500,    0],
  ['Invitations & Stationery', 1400, 1400],
  ['Transportation',           1600,  800],
  ['Cake & Desserts',          2200,  550],
  ['Officiant',                 800,  400],
  ['Decorations & Rentals',    2600, 1300],
  ['Favors & Gifts',            600,  300],
  ['Honeymoon',                8500, 4000],
  ['Accommodations',              0,    0],
  ['Miscellaneous',            1000,    0],
];

const payments = [
  // [date, category, vendor, service, invoice, invoiceAmt, paymentAmt, method, paidBy, payDate, receipt, notes, verified, refund]
  ['11/20/2025','Venue',                'Westlake Estate','Ceremony + Reception deposit', 'WE-1001', 12500, 5000, 'Bank Transfer', 'Marcus Thornton', '11/20/2025', 'RCPT-001', '', true, false],
  ['11/20/2025','Coordinator / Planner','Rivera Planning','Full-service planning deposit', 'RP-2025-11', 4200, 1050, 'Credit Card', 'Amara Osei', '11/20/2025', 'RCPT-002', '', true, false],
  ['2/15/2026', 'Photography',          'Golden Hour Studio','Booking deposit',           'GH-2026-01', 4800, 1200, 'Credit Card', 'Amara Osei', '2/15/2026', 'RCPT-003', '', true, false],
  ['2/25/2026', 'Videography',          'LifeFrame Films', 'Booking deposit',             'LF-2026-01', 3200, 800,  'Credit Card', 'Marcus Thornton', '2/25/2026', 'RCPT-004', '', true, false],
  ['3/5/2026',  'Florals',              'Bloom & Wild',    'Floral design deposit',       'BW-2026-01', 5500, 1100, 'Check',       'Amara Osei', '3/5/2026', 'RCPT-005', '', true, false],
  ['3/15/2026', 'Music / Entertainment','Vibe Events DJ',  'DJ booking deposit',          'VE-2026-03', 2800, 700,  'Credit Card', 'Marcus Thornton', '3/15/2026', 'RCPT-006', '', true, false],
  ['3/20/2026', 'Catering',             'Harvest Table',   'Catering deposit (20%)',      'HT-2026-01', 18000, 3600,'Bank Transfer','Amara Osei', '3/20/2026', 'RCPT-007', '', true, false],
  ['3/22/2026', 'Decorations & Rentals','Event Elegance',  'Rental deposit',              'EE-2026-01', 2600, 650,  'Credit Card', 'Marcus Thornton', '3/22/2026', 'RCPT-008', '', true, false],
  ['3/28/2026', 'Transportation',       'Luxury Limo ATX', 'Transport deposit',           'LL-2026-01', 1600, 400,  'Credit Card', 'Amara Osei', '3/28/2026', 'RCPT-009', '', true, false],
  ['4/5/2026',  'Hair & Makeup',        'Luxe Bridal',     'Beauty package deposit',      'LB-2026-01', 1800, 450,  'Venmo',       'Amara Osei', '4/5/2026', 'RCPT-010', '', true, false],
  ['4/10/2026', 'Honeymoon',            'Golden Gate Travel','Honeymoon deposit',          'GGT-2026-01', 8500, 2000,'Credit Card', 'Marcus Thornton', '4/10/2026', 'RCPT-011', '', true, false],
  ['4/12/2026', 'Music / Entertainment','ClickBox ATX',    'Photo booth deposit',         'CB-2026-01', 900, 225,   'Credit Card', 'Amara Osei', '4/12/2026', 'RCPT-012', '', true, false],
  ['4/20/2026', 'Invitations & Stationery','Paper & Petal','Final stationery balance',    'PP-2026-02', 1400, 1400, 'Check',       'Amara Osei', '4/20/2026', 'RCPT-013', 'Final invoice — paid in full', true, false],
  ['4/20/2026', 'Cake & Desserts',      'Sugar & Spire',   'Cake deposit',                'SS-2026-01', 2200, 550,  'Zelle',       'Marcus Thornton', '4/20/2026', 'RCPT-014', '', true, false],
  ['3/20/2026', 'Lighting',             'Luminary Events', 'Lighting deposit',            'LU-2026-01', 1800, 450,  'Credit Card', 'Amara Osei', '3/20/2026', 'RCPT-015', '', true, false],
  ['3/15/2026', 'Coordinator / Planner','Rivera Planning', '2nd installment',             'RP-2026-03', 4200, 1050, 'Credit Card', 'Amara Osei', '3/15/2026', 'RCPT-016', '', true, false],
  ['4/5/2026',  'Photography',          'Golden Hour Studio','2nd installment',            'GH-2026-02', 4800, 1200, 'Bank Transfer','Marcus Thornton','4/5/2026','RCPT-017','',true,false],
  ['4/5/2026',  'Videography',          'LifeFrame Films', '2nd installment',             'LF-2026-02', 3200, 800,  'Bank Transfer','Marcus Thornton','4/5/2026','RCPT-018','',true,false],
  ['4/10/2026', 'Florals',              'Bloom & Wild',    '2nd installment',             'BW-2026-02', 5500, 1650, 'Check',       'Amara Osei', '4/10/2026', 'RCPT-019', '', true, false],
  ['4/15/2026', 'Catering',             'Harvest Table',   'Progress payment',            'HT-2026-02', 18000, 5400,'Bank Transfer','Amara Osei','4/15/2026','RCPT-020','',true,false],
  ['4/15/2026', 'Officiant',            'Rev. David Nguyen','Ceremony fee partial',       'DN-2026-01', 800, 400,   'Check',       'Marcus Thornton', '4/15/2026', 'RCPT-021', '', true, false],
  ['4/15/2026', 'Decorations & Rentals','Event Elegance',  '2nd installment',             'EE-2026-02', 2600, 650,  'Credit Card', 'Marcus Thornton', '4/15/2026', 'RCPT-022', '', true, false],
  ['5/1/2026',  'Transportation',       'Luxury Limo ATX', '2nd installment',             'LL-2026-02', 1600, 400,  'Credit Card', 'Amara Osei', '5/1/2026', 'RCPT-023', '', true, false],
  ['5/1/2026',  'Music / Entertainment','Vibe Events DJ',  '2nd installment',             'VE-2026-04', 2800, 700,  'Credit Card', 'Marcus Thornton', '5/1/2026', 'RCPT-024', '', true, false],
  ['5/5/2026',  'Hair & Makeup',        'Luxe Bridal',     '2nd installment',             'LB-2026-02', 1800, 450,  'Venmo',       'Amara Osei', '5/5/2026', 'RCPT-025', '', true, false],
  ['5/10/2026', 'Honeymoon',            'Golden Gate Travel','2nd payment',               'GGT-2026-02', 8500, 2000,'Credit Card','Marcus Thornton','5/10/2026','RCPT-026','',true,false],
  ['5/15/2026', 'Lighting',             'Luminary Events', '2nd installment',             'LU-2026-02', 1800, 450,  'Credit Card', 'Amara Osei', '5/15/2026', 'RCPT-027', '', true, false],
  ['5/20/2026', 'Favors & Gifts',       'Artisan Gifting', 'Favors partial payment',      'AG-2026-01', 600, 300,   'PayPal',      'Amara Osei', '5/20/2026', 'RCPT-028', '', true, false],
  ['5/25/2026', 'Photography',          'Golden Hour Studio','Final payment',              'GH-2026-03', 4800, 2400, 'Bank Transfer','Marcus Thornton','5/25/2026','RCPT-029','Paid in full',true,false],
  ['5/25/2026', 'Videography',          'LifeFrame Films', 'Final payment',               'LF-2026-03', 3200, 1600, 'Bank Transfer','Marcus Thornton','5/25/2026','RCPT-030','Paid in full',true,false],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['BUDGET & PAYMENTS TRACKER']] });
  reqs.push({ mergeCells: { range: gridRange(BP, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BP, 0, 1, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BP, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Budget overview KPIs row 2 ───────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [[
    'Total Budget:','', `=IFERROR('Wedding Setup'!B20,0)`, '',
    'Total Allocated:','', `=IFERROR(SUM(B5:B22),0)`, '',
    'Total Paid:','', `=IFERROR(SUM(H40:H400),0)`, '',
    'Balance Remaining:','', `=IFERROR('Wedding Setup'!B20-SUM(H40:H400),0)`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(BP, 1, 2, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' }});

  // ── SECTION 1: CATEGORY BUDGETS ───────────────────────────────────────────
  data.push({ range: `${S}!A3`, values: [['  SECTION 1 — CATEGORY BUDGETS']] });
  reqs.push({ mergeCells: { range: gridRange(BP, 2, 3, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BP, 2, 3, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  const budgetHdrs = ['Category', 'Budgeted', 'Estimated Actual', 'Variance (auto)', '% Used (auto)', 'Notes'];
  data.push({ range: `${S}!A4`, values: [budgetHdrs] });
  reqs.push({ repeatCell: { range: gridRange(BP, 3, 4, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  budgetCategories.forEach((bc, i) => {
    const row = 5 + i;
    const [cat, budget, actual] = bc;
    data.push({ range: `${S}!A${row}`, values: [[cat]] });
    data.push({ range: `${S}!B${row}`, values: [[budget]] });
    data.push({ range: `${S}!C${row}`, values: [[actual]] });
    // D: Variance
    data.push({ range: `${S}!D${row}`, values: [[`=IFERROR(B${row}-C${row},0)`]] });
    // E: % Used
    data.push({ range: `${S}!E${row}`, values: [[`=IFERROR(C${row}/B${row},0)`]] });
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(BP, row - 1, row, 0, 6), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  });

  const budgetEnd = 5 + budgetCategories.length;

  // Totals row
  data.push({ range: `${S}!A${budgetEnd}`, values: [['TOTAL']] });
  data.push({ range: `${S}!B${budgetEnd}`, values: [[`=IFERROR(SUM(B5:B${budgetEnd-1}),0)`]] });
  data.push({ range: `${S}!C${budgetEnd}`, values: [[`=IFERROR(SUM(C5:C${budgetEnd-1}),0)`]] });
  data.push({ range: `${S}!D${budgetEnd}`, values: [[`=IFERROR(B${budgetEnd}-C${budgetEnd},0)`]] });
  data.push({ range: `${S}!E${budgetEnd}`, values: [[`=IFERROR(C${budgetEnd}/B${budgetEnd},0)`]] });
  reqs.push({ repeatCell: { range: gridRange(BP, budgetEnd - 1, budgetEnd, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) },
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' }});

  // Format B, C, D as currency; E as percent
  reqs.push({ repeatCell: { range: gridRange(BP, 4, budgetEnd, 1, 4), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(BP, 4, budgetEnd, 4, 5), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // Formula cols D, E
  reqs.push({ repeatCell: { range: gridRange(BP, 4, budgetEnd - 1, 3, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // CF: over budget → red; > 80% → amber
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(BP, 4, budgetEnd - 1, 4, 5)],
      booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] },
        format: { backgroundColor: hex(C.attention), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
    }, index: 0,
  }});
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(BP, 4, budgetEnd - 1, 4, 5)],
      booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.8' }] },
        format: { backgroundColor: hex(C.warning) } },
    }, index: 0,
  }});

  // ── SECTION 2: PAYMENT LOG ────────────────────────────────────────────────
  const payStart = budgetEnd + 2;
  data.push({ range: `${S}!A${payStart - 1}`, values: [['  SECTION 2 — PAYMENT LOG']] });
  reqs.push({ mergeCells: { range: gridRange(BP, payStart - 2, payStart - 1, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BP, payStart - 2, payStart - 1, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Payment summary row
  data.push({ range: `${S}!A${payStart}`, values: [[
    'Total Payments:','', `=IFERROR(COUNTA(B${payStart+2}:B${payStart+200}),"")`, '',
    'Total Paid:','', `=IFERROR(SUM(H${payStart+2}:H${payStart+200}),0)`, '',
    'Verified:','', `=IFERROR(SUMPRODUCT((O${payStart+2}:O${payStart+200}=TRUE)*1),"")`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(BP, payStart - 1, payStart, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow), textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Payment headers
  data.push({ range: `${S}!A${payStart + 1}`, values: [PAYMENT_HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(BP, payStart, payStart + 1, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BP, dimension: 'ROWS', startIndex: payStart, endIndex: payStart + 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // Payment data
  let runningTotal = 0;
  payments.forEach((p, i) => {
    const row = payStart + 2 + i;
    const [date, cat, vendor, service, invoice, invoiceAmt, payAmt, method, paidBy, payDate, receipt, notes, verified, refund] = p;
    runningTotal += payAmt;
    data.push({ range: `${S}!A${row}`, values: [[i + 1]] });
    data.push({ range: `${S}!B${row}`, values: [[date]] });
    data.push({ range: `${S}!C${row}`, values: [[cat]] });
    data.push({ range: `${S}!D${row}`, values: [[vendor]] });
    data.push({ range: `${S}!E${row}`, values: [[service]] });
    data.push({ range: `${S}!F${row}`, values: [[invoice]] });
    data.push({ range: `${S}!G${row}`, values: [[invoiceAmt]] });
    data.push({ range: `${S}!H${row}`, values: [[payAmt]] });
    data.push({ range: `${S}!I${row}`, values: [[method]] });
    data.push({ range: `${S}!J${row}`, values: [[paidBy]] });
    data.push({ range: `${S}!K${row}`, values: [[payDate]] });
    data.push({ range: `${S}!L${row}`, values: [[receipt]] });
    // M: Running total
    data.push({ range: `${S}!M${row}`, values: [[i === 0 ? `=IFERROR(H${row},0)` : `=IFERROR(M${row-1}+H${row},0)`]] });
    data.push({ range: `${S}!N${row}`, values: [[notes]] });
    data.push({ range: `${S}!O${row}`, values: [[verified]] });
    data.push({ range: `${S}!P${row}`, values: [[refund]] });
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(BP, row - 1, row, 0, 16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  });

  const payEnd = payStart + 2 + payments.length;

  // Checkboxes: O (Verified), P (Refund)
  [14, 15].forEach(col => {
    reqs.push({ setDataValidation: {
      range: gridRange(BP, payStart + 1, payEnd, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Payment Method dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(BP, payStart + 1, payStart + 200, 8, 9),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Credit Card', 'Debit Card', 'Bank Transfer', 'Check', 'Cash', 'Venmo', 'Zelle', 'PayPal', 'Other',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Currency format: G, H, M
  [6, 7, 12].forEach(col => {
    reqs.push({ repeatCell: { range: gridRange(BP, payStart + 1, payEnd, col, col + 1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Running total is formula
  reqs.push({ repeatCell: { range: gridRange(BP, payStart + 1, payEnd, 12, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Payment total row
  data.push({ range: `${S}!G${payEnd}`, values: [['TOTAL PAID']] });
  data.push({ range: `${S}!H${payEnd}`, values: [[`=IFERROR(SUM(H${payStart+2}:H${payEnd-1}),0)`]] });
  reqs.push({ repeatCell: { range: gridRange(BP, payEnd - 1, payEnd, 6, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) },
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const widths = [35, 80, 130, 160, 180, 90, 90, 90, 100, 120, 80, 80, 100, 180, 60, 60];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: BP, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze row 1+budget headers
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BP, gridProperties: { frozenRowCount: 4 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BP, tabColor: hex(C.warning) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'budget-data');
  await batchUpdate(id, reqs, 'budget-format');

  console.log('✓ Budget & Payments built with', budgetCategories.length, 'categories and', payments.length, 'payments');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
