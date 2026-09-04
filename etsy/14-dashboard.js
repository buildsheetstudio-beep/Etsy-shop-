'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Dashboard'];
const S  = "'Dashboard'";

const yr = 2026;

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 70 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 4 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 4, endIndex: 100 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Column widths — 12 cols (A-L)
  const colW = [30,200,110,110,110,110,30,200,110,110,110,110];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:L1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 22 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // KPI section — 3 rows of cards (rows 2-4, 5-7, 8-10, indices 1-9)
  // Left panel A-F, right panel H-L, col G is divider
  // Row 2: labels, Row 3: values, Row 4: subtitles
  const kpiGroups = [
    // [c0, c1, color, label, valueFormula, subtitle]
    [0,  3, C.success, 'Gross Revenue (YTD)', `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)`, 'All income sources 2026'],
    [3,  6, C.primary, 'Net Revenue (YTD)', `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$N$6:$N$500)`, 'After Etsy fees & refunds'],
    [7, 10, C.attention, 'Total Expenses (YTD)', `=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*'Expense Log'!$H$6:$H$500)`, 'All business costs 2026'],
    [10, 13, C.warning, 'Net Profit (YTD)', `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$N$6:$N$500)-SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*'Expense Log'!$H$6:$H$500)`, 'Revenue minus all expenses'],
  ];

  kpiGroups.forEach(([c0, c1, color, label, formula, sub]) => {
    // Label row (index 1)
    reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    // Value row (index 2)
    reqs.push({ mergeCells: { range: gridRange(SH, 2, 3, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { foregroundColor: hex(color), bold: true, fontSize: 24 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    // Subtitle row (index 3)
    reqs.push({ mergeCells: { range: gridRange(SH, 3, 4, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 3, 4, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });

  // Second row of KPI cards (rows 5-7, indices 4-6)
  const kpiGroups2 = [
    [0,  3, C.mutedBlue,  'Total Orders (YTD)', `=COUNTIFS(YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received")`, 'Etsy + off-Etsy orders'],
    [3,  6, C.secondary, 'Avg Order Value', `=IFERROR(SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$P$6:$P$500="Received")*'Income & Orders'!$H$6:$H$500)/COUNTIFS(YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received"),0)`, 'Per completed order'],
    [7, 10, C.mutedGold, 'Profit Margin %', `=IFERROR((SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$N$6:$N$500)-SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*'Expense Log'!$H$6:$H$500))/SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500),0)`, 'Goal: 70%+ margin'],
    [10, 13, C.success,  'Revenue Goal Progress', `=IFERROR(SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)/85000,0)`, 'YTD vs $85K annual goal'],
  ];

  kpiGroups2.forEach(([c0, c1, color, label, formula, sub]) => {
    reqs.push({ mergeCells: { range: gridRange(SH, 4, 5, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 4, 5, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 5, 6, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 6, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { foregroundColor: hex(C.primary), bold: true, fontSize: 22 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 6, 7, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 6, 7, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });

  // Divider column G (index 6)
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 100, 6, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' }});

  // Main content sections - left panel (cols A-F, 0-5) and right panel (cols H-L, 7-12)

  // Left: Recent Income summary (rows 8-20, indices 8-20)
  reqs.push({ mergeCells: { range: gridRange(SH, 7, 8, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 7, 8, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.success),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  reqs.push({ repeatCell: { range: gridRange(SH, 8, 9, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  for (let r = 9; r < 22; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 6), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Right: Expense summary (rows 8-20)
  reqs.push({ mergeCells: { range: gridRange(SH, 7, 8, 7, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 7, 8, 7, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.attention),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  reqs.push({ repeatCell: { range: gridRange(SH, 8, 9, 7, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  for (let r = 9; r < 22; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 7, 12), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Monthly trend section (rows 23-36)
  reqs.push({ mergeCells: { range: gridRange(SH, 22, 23, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 22, 23, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  reqs.push({ repeatCell: { range: gridRange(SH, 23, 24, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  for (let r = 24; r < 37; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 12), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Total row 37
  reqs.push({ repeatCell: { range: gridRange(SH, 36, 37, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Currency format for KPI values
  [[0,3],[3,6],[7,10],[10,13]].forEach(([c0,c1]) => {
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, c0, c1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Second row KPI — avg order value currency, margin %, progress %
  [[0,3],[3,6]].forEach(([c0,c1]) => {
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 6, c0, c1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  [[3,6]].forEach(([c0,c1]) => {
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 6, c0, c1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  [[7,10],[10,13]].forEach(([c0,c1]) => {
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 6, c0, c1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Data area currency
  [2,3,4,5].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 9, 37, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
    reqs.push({ repeatCell: { range: gridRange(SH, 9, 37, c+7, c+8), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 7, 37, 0, 12),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // CF on profit margin card — color the text green if ≥70%, red if <60%
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SH, 5, 6, 7, 10)],
    booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.7' }] }, format: { textFormat: { foregroundColor: hex(C.success) } } },
  }, index: 0 }});
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SH, 5, 6, 7, 10)],
    booleanRule: { condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0.6' }] }, format: { textFormat: { foregroundColor: hex(C.attention) } } },
  }, index: 1 }});

  // Freeze rows 1
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'dashboard-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['🌿 Ultimate Etsy Seller Bookkeeping System — Willow Paper Studio · ${yr}']] });

  // KPI card labels & values
  kpiGroups.forEach(([c0,,color, label, formula, sub]) => {
    const col = String.fromCharCode(65+c0);
    data.push({ range: `${S}!${col}2`, values: [[`  ${label}`]] });
    data.push({ range: `${S}!${col}3`, values: [[formula]] });
    data.push({ range: `${S}!${col}4`, values: [[sub]] });
  });

  kpiGroups2.forEach(([c0,,color, label, formula, sub]) => {
    const col = String.fromCharCode(65+c0);
    data.push({ range: `${S}!${col}5`, values: [[`  ${label}`]] });
    data.push({ range: `${S}!${col}6`, values: [[formula]] });
    data.push({ range: `${S}!${col}7`, values: [[sub]] });
  });

  // Recent income section
  data.push({ range: `${S}!A8`, values: [['  💰 Recent Income — Last 10 Orders']] });
  data.push({ range: `${S}!A9:F9`, values: [['Date','Source','Item','Gross','Net','Status']] });

  // FILTER last 10 income rows by date descending
  for (let i = 0; i < 10; i++) {
    const row = 10 + i;
    const n = i + 1;
    data.push({ range: `${S}!A${row}:F${row}`, values: [[
      `=IFERROR(INDEX(SORT('Income & Orders'!$B$6:$S$500,1,-1),${n},1),"")`,
      `=IFERROR(INDEX(SORT('Income & Orders'!$B$6:$S$500,1,-1),${n},5),"")`,
      `=IFERROR(INDEX(SORT('Income & Orders'!$B$6:$S$500,1,-1),${n},6),"")`,
      `=IFERROR(INDEX(SORT('Income & Orders'!$B$6:$S$500,1,-1),${n},7),"")`,
      `=IFERROR(INDEX(SORT('Income & Orders'!$B$6:$S$500,1,-1),${n},13),"")`,
      `=IFERROR(INDEX(SORT('Income & Orders'!$B$6:$S$500,1,-1),${n},15),"")`,
    ]] });
  }

  // Recent expenses section
  data.push({ range: `${S}!H8`, values: [['  💸 Recent Expenses — Last 10']] });
  data.push({ range: `${S}!H9:L9`, values: [['Date','Category','Description','Amount','Status']] });

  for (let i = 0; i < 10; i++) {
    const row = 10 + i;
    const n = i + 1;
    data.push({ range: `${S}!H${row}:L${row}`, values: [[
      `=IFERROR(INDEX(SORT('Expense Log'!$A$6:$O$500,1,-1),${n},1),"")`,
      `=IFERROR(INDEX(SORT('Expense Log'!$A$6:$O$500,1,-1),${n},5),"")`,
      `=IFERROR(INDEX(SORT('Expense Log'!$A$6:$O$500,1,-1),${n},7),"")`,
      `=IFERROR(INDEX(SORT('Expense Log'!$A$6:$O$500,1,-1),${n},8),"")`,
      `=IFERROR(INDEX(SORT('Expense Log'!$A$6:$O$500,1,-1),${n},11),"")`,
    ]] });
  }

  // Monthly trend section
  data.push({ range: `${S}!A23`, values: [['  📅 Monthly Performance Trend — ${yr}']] });
  data.push({ range: `${S}!A24:L24`, values: [['Month','Gross Rev','Expenses','Net Profit','Margin %','vs Goal','','Orders','Avg Order','Top Product','Etsy Fees','Fee Rate %']] });

  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  MONTHS.forEach((m, mi) => {
    const r = 25 + mi;
    const mo = mi + 1;
    data.push({ range: `${S}!A${r}:L${r}`, values: [[
      m,
      `=SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=${mo})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)`,
      `=SUMPRODUCT((MONTH('Expense Log'!$A$6:$A$500)=${mo})*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*'Expense Log'!$H$6:$H$500)`,
      `=B${r}-C${r}`,
      `=IFERROR(D${r}/B${r},0)`,
      `=B${r}-7083.33`,
      '',
      `=COUNTIFS(MONTH('Income & Orders'!$B$6:$B$500),${mo},YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received")`,
      `=IFERROR(B${r}/H${r},0)`,
      `=IFERROR(INDEX(SORT(UNIQUE(FILTER('Income & Orders'!$G$6:$G$500,(MONTH('Income & Orders'!$B$6:$B$500)=${mo})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$H$6:$H$500)>0)),1,1),1,1),"")`,
      `=SUMPRODUCT((MONTH('Expense Log'!$A$6:$A$500)=${mo})*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*(ISNUMBER(MATCH('Expense Log'!$E$6:$E$500,{"Etsy Transaction Fees","Etsy Listing Fees","Etsy Payment Processing Fees","Etsy Ads","Offsite Ads"},0)))*'Expense Log'!$H$6:$H$500)`,
      `=IFERROR(K${r}/B${r},0)`,
    ]] });
  });

  // Total row
  data.push({ range: `${S}!A37:L37`, values: [[
    'FULL YEAR',
    '=SUM(B25:B36)','=SUM(C25:C36)','=SUM(D25:D36)',
    `=IFERROR(D37/B37,0)`, '=B37-85000', '',
    '=SUM(H25:H36)',
    `=IFERROR(B37/H37,0)`, '',
    '=SUM(K25:K36)',
    `=IFERROR(K37/B37,0)`,
  ]] });

  // Date format for recent income/expense date columns
  reqs.push({ repeatCell: { range: gridRange(SH, 9, 22, 0, 1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yy' } } }, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(SH, 9, 22, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yy' } } }, fields: 'userEnteredFormat.numberFormat' }});
  // Margin % and fee rate % in monthly trend
  [4,11].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 24, 38, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Avg order currency
  reqs.push({ repeatCell: { range: gridRange(SH, 24, 38, 8, 9), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});

  await valuesBatchUpdate(id, data, 'dashboard-values');
  await batchUpdate(id, reqs, 'dashboard-post-format');

  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
