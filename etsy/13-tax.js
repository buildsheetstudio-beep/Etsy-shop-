'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, EXPENSE_CATS, TAX_GROUP_MAP, TAX_GROUPS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Tax Deduction Summary'];
const S  = "'Tax Deduction Summary'";

const yr = 2026;

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 3, endIndex: 200 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,50],[1,210],[2,140],[3,140],[4,140],[5,140],[6,140]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:G1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 17 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Summary card row 2-3
  [[0,2],[2,4],[4,7]].forEach(([c0,c1], idx) => {
    const colors = [C.success, C.primary, C.attention];
    reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(colors[idx]),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 2, 3, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { foregroundColor: hex(colors[idx]), bold: true, fontSize: 20 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 3, 4, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 3, 4, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });

  // Section helper
  function sectionHdr(r, color) {
    reqs.push({ mergeCells: { range: gridRange(SH, r, r+1, 0, 7), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 7), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function colHdr(r) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 7), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function dataBlock(r1, r2) {
    for (let r = r1; r < r2; r++) {
      const bg = r % 2 === 0 ? C.panel : C.altRow;
      reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 7), cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    }
  }
  function totalRow(r, color) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 7), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }

  // By tax group: row 5 hdr, 6 col hdr, 7-20 data, 21 total
  sectionHdr(4, C.primary);
  colHdr(5);
  dataBlock(6, 20);
  totalRow(20, C.primary);

  // By expense category: row 22 hdr, 23 col hdr, 24-53 data, 54 total
  sectionHdr(21, C.secondary);
  colHdr(22);
  dataBlock(23, 53);
  totalRow(53, C.secondary);

  // Notes section
  sectionHdr(54, C.mutedBlue);
  dataBlock(55, 65);

  // Currency format
  [2,3,4,5,6].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 55, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Percent col F (5) in group table
  reqs.push({ repeatCell: { range: gridRange(SH, 6, 21, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(SH, 23, 54, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 4, 65, 0, 7),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze row 2
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'tax-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['🧾 Tax Deduction Summary — ${yr} Schedule C Deductions']] });

  // KPI cards
  const totalDeductible = `=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$I$6:$I$500="Yes")*'Expense Log'!$H$6:$H$500)`;
  const totalExpenses   = `=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*'Expense Log'!$H$6:$H$500)`;
  const totalRevenue    = `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)`;

  data.push({ range: `${S}!A2`, values: [['  Total Deductible Expenses']] });
  data.push({ range: `${S}!A3`, values: [[totalDeductible]] });
  data.push({ range: `${S}!A4`, values: [['Tax-deductible business expenses']] });

  data.push({ range: `${S}!C2`, values: [['  Net Business Income']] });
  data.push({ range: `${S}!C3`, values: [[`=${totalRevenue}-${totalDeductible}`]] });
  data.push({ range: `${S}!C4`, values: [['Revenue minus deductible expenses']] });

  data.push({ range: `${S}!E2`, values: [['  Deductible % of Total Expenses']] });
  data.push({ range: `${S}!E3`, values: [[`=IFERROR(${totalDeductible}/${totalExpenses},0)`]] });
  data.push({ range: `${S}!E4`, values: [['% of expenses that are deductible']] });

  // Tax group summary
  data.push({ range: `${S}!A5`, values: [['  📊 Deductions by Tax Group (Schedule C Categories)']] });
  data.push({ range: `${S}!A6:G6`, values: [['#','Tax Group','Q1','Q2','Q3','Q4 (Est.)','Full Year']] });

  // Unique tax groups (from TAX_GROUPS constant)
  TAX_GROUPS.forEach((group, i) => {
    const r = 7 + i;
    const cats = EXPENSE_CATS.filter(c => TAX_GROUP_MAP[c] === group);
    const catList = cats.map(c => `"${c}"`).join(',');
    if (!catList) {
      data.push({ range: `${S}!A${r}:G${r}`, values: [[i+1, group, 0, 0, 0, 0, 0]] });
      return;
    }
    const qFn = (months) =>
      `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Expense Log'!$A$6:$A$500),{${months}},0)))*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*(ISNUMBER(MATCH('Expense Log'!$E$6:$E$500,{${catList}},0)))*('Expense Log'!$I$6:$I$500="Yes")*'Expense Log'!$H$6:$H$500)`;
    data.push({ range: `${S}!A${r}:G${r}`, values: [[
      i+1, group,
      qFn('1,2,3'), qFn('4,5,6'), qFn('7,8,9'), qFn('10,11,12'),
      `=SUM(C${r}:F${r})`,
    ]] });
  });

  // Total row 21
  data.push({ range: `${S}!A21:G21`, values: [['','TOTAL','=SUM(C7:C20)','=SUM(D7:D20)','=SUM(E7:E20)','=SUM(F7:F20)','=SUM(G7:G20)']] });

  // By expense category
  data.push({ range: `${S}!A22`, values: [['  📋 Deductions by Expense Category']] });
  data.push({ range: `${S}!A23:G23`, values: [['#','Expense Category','Deductible Amount','Tax Group','Q1','Q2','Full Year']] });

  EXPENSE_CATS.forEach((cat, i) => {
    const r = 24 + i;
    const taxGroup = TAX_GROUP_MAP[cat] || 'Other / Review Required';
    data.push({ range: `${S}!A${r}:G${r}`, values: [[
      i+1, cat,
      `=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*('Expense Log'!$I$6:$I$500="Yes")*'Expense Log'!$H$6:$H$500)`,
      taxGroup,
      `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Expense Log'!$A$6:$A$500),{1,2,3},0)))*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*('Expense Log'!$I$6:$I$500="Yes")*'Expense Log'!$H$6:$H$500)`,
      `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Expense Log'!$A$6:$A$500),{4,5,6},0)))*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*('Expense Log'!$I$6:$I$500="Yes")*'Expense Log'!$H$6:$H$500)`,
      `=IFERROR(C${r}/C$54,0)`,
    ]] });
  });
  data.push({ range: `${S}!A54:G54`, values: [['','TOTAL',`=SUM(C24:C53)`, '', `=SUM(E24:E53)`,`=SUM(F24:F53)`, '100%']] });

  // Important tax notes
  data.push({ range: `${S}!A55`, values: [['  📌 Important Tax Notes & Reminders']] });
  const notes = [
    ['Self-Employment Tax', '15.3% on net self-employment income (Schedule SE) — also deductible (50% deduction)'],
    ['Quarterly Estimated Taxes', 'Due: Apr 15, Jun 15, Sep 15, Jan 15 — use Form 1040-ES to calculate'],
    ['Home Office Deduction', 'Simplified method: $5/sq ft up to 300 sq ft = $1,500 max. Regular method: % of home used'],
    ['Mileage Rate (2026)', 'IRS standard mileage: 67¢/mile for business travel (check IRS.gov for current rate)'],
    ['Receipt Requirement', 'Keep all receipts for expenses over $75; digital copies acceptable'],
    ['Section 179 / Bonus Depreciation', 'Equipment and software may qualify for immediate full deduction in year of purchase'],
    ['Retirement Contributions', 'SEP-IRA: up to 25% of net self-employment income; Solo 401(k) may allow more'],
    ['Health Insurance Premiums', 'Self-employed health insurance premiums deductible on Schedule 1 (not Schedule C)'],
    ['Consult a CPA', 'These figures are for bookkeeping purposes only — consult a licensed CPA for tax advice'],
  ];
  notes.forEach(([topic, note], i) => {
    data.push({ range: `${S}!A${56+i}:G${56+i}`, values: [[`• ${topic}`, note, '', '', '', '', '']] });
  });

  await valuesBatchUpdate(id, data, 'tax-values');

  // Merge note rows across cols
  for (let r = 55; r < 65; r++) {
    reqs.push({ mergeCells: { range: gridRange(SH, r, r+1, 1, 7), mergeType: 'MERGE_ALL' } });
  }
  // Percent for card E
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, 4, 7), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  await batchUpdate(id, reqs, 'tax-post-format');

  console.log('Tax Deduction Summary complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
