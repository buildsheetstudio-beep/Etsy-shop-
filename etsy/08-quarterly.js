'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, EXPENSE_CATS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Quarterly Summary'];
const S  = "'Quarterly Summary'";

const QUARTERS = ['Q1','Q2','Q3','Q4'];
const QUARTER_MONTHS = { Q1:[1,2,3], Q2:[4,5,6], Q3:[7,8,9], Q4:[10,11,12] };
const yr = 2026;

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 3, endIndex: 200 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,50],[1,210],[2,140],[3,140],[4,140],[5,140],[6,120],[7,130]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:H1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Subheader row 2 A2:H2
  reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Column header row 3 A3:H3
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Main comparison table rows 4-9 (indices 3-8)
  for (let r = 3; r < 9; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Section headers for income/expense breakdown
  function sectionHdr(r, color) {
    reqs.push({ mergeCells: { range: gridRange(SH, r, r+1, 0, 8), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function colHdr(r) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function dataBlock(r1, r2) {
    for (let r = r1; r < r2; r++) {
      const bg = r % 2 === 0 ? C.panel : C.altRow;
      reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 8), cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    }
  }
  function totalRow(r, color) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }

  // Income breakdown: row 10 hdr, 11 col hdr, 12-21 data, 22 total
  sectionHdr(9, C.success);
  colHdr(10);
  dataBlock(11, 21);
  totalRow(21, C.success);

  // Expense breakdown: row 23 hdr, 24 col hdr, 25-54 data, 55 total
  sectionHdr(22, C.attention);
  colHdr(23);
  dataBlock(24, 54);
  totalRow(54, C.attention);

  // Net Profit: row 56 hdr, 57-60 data
  sectionHdr(55, C.primary);
  dataBlock(56, 61);

  // Currency format
  [2,3,4,5,6,7].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 3, 62, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Margin % col H (7) in net profit
  reqs.push({ repeatCell: { range: gridRange(SH, 56, 62, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 2, 62, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze row 3
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'quarterly-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['📊 Quarterly Summary — ${yr} Performance by Quarter']] });
  data.push({ range: `${S}!A2`, values: [['All four quarters · Year ${yr} · Willow Paper Studio']] });
  data.push({ range: `${S}!A3:H3`, values: [['Metric','Q1 (Jan-Mar)','Q2 (Apr-Jun)','Q3 (Jul-Sep)','Q4 (Oct-Dec)','Full Year','Goal','vs Goal']] });

  // Summary table
  const metrics = [
    ['Gross Revenue'],
    ['Total Expenses'],
    ['Net Profit'],
    ['Profit Margin %'],
    ['Order Count'],
    ['Avg Order Value'],
  ];

  // QUARTER formula helper
  function qFilter(sheet, dateCol, valCol, q) {
    const months = QUARTER_MONTHS[q];
    return `SUMPRODUCT((ISNUMBER(MATCH(MONTH(${sheet}!${dateCol}6:${dateCol}500),{${months.join(',')}},0)))*(YEAR(${sheet}!${dateCol}6:${dateCol}500)=${yr})*${sheet}!${valCol}6:${valCol}500)`;
  }

  data.push({ range: `${S}!A4:H4`, values: [[
    'Gross Revenue',
    `=${qFilter("'Income & Orders'","$B$","$H$","Q1")}`,
    `=${qFilter("'Income & Orders'","$B$","$H$","Q2")}`,
    `=${qFilter("'Income & Orders'","$B$","$H$","Q3")}`,
    `=${qFilter("'Income & Orders'","$B$","$H$","Q4")}`,
    '=SUM(C4:F4)',
    `='Business Setup'!B13`,
    '=F4-G4',
  ]] });

  data.push({ range: `${S}!A5:H5`, values: [[
    'Total Expenses',
    `=${qFilter("'Expense Log'","$A$","$H$","Q1")}`,
    `=${qFilter("'Expense Log'","$A$","$H$","Q2")}`,
    `=${qFilter("'Expense Log'","$A$","$H$","Q3")}`,
    `=${qFilter("'Expense Log'","$A$","$H$","Q4")}`,
    '=SUM(C5:F5)',
    '',
    '',
  ]] });

  data.push({ range: `${S}!A6:H6`, values: [['Net Profit','=C4-C5','=D4-D5','=E4-E5','=F4-F5','=SUM(C6:F6)',`='Business Setup'!B14`,'=F6-G6']] });
  data.push({ range: `${S}!A7:H7`, values: [['Profit Margin %','=IFERROR(C6/C4,0)','=IFERROR(D6/D4,0)','=IFERROR(E6/E4,0)','=IFERROR(F6/F4,0)','=IFERROR(F6/F4,0)','=Business Setup\'!B18','=IFERROR(F7-G7,0)']] });

  // Order count
  function qOrders(q) {
    const months = QUARTER_MONTHS[q];
    return `COUNTPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{${months.join(',')}},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$P$6:$P$500="Received"))`;
  }
  data.push({ range: `${S}!A8:H8`, values: [[
    'Order Count',
    `=COUNTIFS(MONTH('Income & Orders'!$B$6:$B$500)>0,ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{1,2,3},0)),YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received")`,
    `=COUNTIFS(MONTH('Income & Orders'!$B$6:$B$500)>0,ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{4,5,6},0)),YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received")`,
    `=COUNTIFS(MONTH('Income & Orders'!$B$6:$B$500)>0,ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{7,8,9},0)),YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received")`,
    `=COUNTIFS(MONTH('Income & Orders'!$B$6:$B$500)>0,ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{10,11,12},0)),YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received")`,
    '=SUM(C8:F8)', '', '',
  ]] });
  data.push({ range: `${S}!A9:H9`, values: [['Avg Order Value','=IFERROR(C4/C8,0)','=IFERROR(D4/D8,0)','=IFERROR(E4/E8,0)','=IFERROR(F4/F8,0)','=IFERROR(F4/F8,0)','','=IFERROR(F9-G9,0)']] });

  // Quarterly % margin format — row 7
  [2,3,4,5,6,7].forEach(c => {
    data.push({ range: `${S}!${String.fromCharCode(65+c)}7`, values: [[`=IFERROR(${String.fromCharCode(65+c)}6/${String.fromCharCode(65+c)}4,0)`]] });
  });

  // Income by source
  data.push({ range: `${S}!A10`, values: [['  💰 Income Breakdown by Source']] });
  data.push({ range: `${S}!A11:H11`, values: [['#','Source','Q1','Q2','Q3','Q4','Full Year','YTD %']] });
  const SOURCES = ['Etsy Order','Etsy Shipping Collected','Etsy Refund Adjustment','Off-Etsy Sale','Wholesale','Custom Order','Licensing','Affiliate Income','Other Business Income'];
  SOURCES.forEach((src, i) => {
    const r = 12 + i;
    data.push({ range: `${S}!A${r}:H${r}`, values: [[
      i+1, src,
      ...QUARTERS.map(q => {
        const months = QUARTER_MONTHS[q];
        return `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{${months.join(',')}},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$F$6:$F$500="${src}")*'Income & Orders'!$H$6:$H$500)`;
      }),
      `=SUM(C${r}:F${r})`,
      `=IFERROR(G${r}/G$22,0)`,
    ]] });
  });
  data.push({ range: `${S}!A22:H22`, values: [['','TOTAL','=SUM(C12:C21)','=SUM(D12:D21)','=SUM(E12:E21)','=SUM(F12:F21)','=SUM(G12:G21)','100%']] });

  // Expenses by category
  data.push({ range: `${S}!A23`, values: [['  💸 Expense Breakdown by Category']] });
  data.push({ range: `${S}!A24:H24`, values: [['#','Category','Q1','Q2','Q3','Q4','Full Year','YTD %']] });
  EXPENSE_CATS.slice(0, 28).forEach((cat, i) => {
    const r = 25 + i;
    data.push({ range: `${S}!A${r}:H${r}`, values: [[
      i+1, cat,
      ...QUARTERS.map(q => {
        const months = QUARTER_MONTHS[q];
        return `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Expense Log'!$A$6:$A$500),{${months.join(',')}},0)))*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*'Expense Log'!$H$6:$H$500)`;
      }),
      `=SUM(C${r}:F${r})`,
      `=IFERROR(G${r}/G$55,0)`,
    ]] });
  });
  data.push({ range: `${S}!A55:H55`, values: [['','TOTAL','=SUM(C25:C53)','=SUM(D25:D53)','=SUM(E25:E53)','=SUM(F25:F53)','=SUM(G25:G53)','100%']] });

  // Net profit section
  data.push({ range: `${S}!A56`, values: [['  📊 Quarterly Profit Summary']] });
  data.push({ range: `${S}!A57:H57`, values: [['','Metric','Q1','Q2','Q3','Q4','Full Year','Margin %']] });
  data.push({ range: `${S}!A58:H58`, values: [['','Net Revenue','=C4','=D4','=E4','=F4','=SUM(C58:F58)','=IFERROR(G58/G58,1)']] });
  data.push({ range: `${S}!A59:H59`, values: [['','Total Expenses','=C5','=D5','=E5','=F5','=SUM(C59:F59)','=IFERROR(G59/G58,0)']] });
  data.push({ range: `${S}!A60:H60`, values: [['','Net Profit','=C6','=D6','=E6','=F6','=SUM(C60:F60)','=IFERROR(G60/G58,0)']] });
  data.push({ range: `${S}!A61:H61`, values: [['','Profit Margin','=C7','=D7','=E7','=F7','=IFERROR(G60/G58,0)','']] });

  await valuesBatchUpdate(id, data, 'quarterly-values');
  console.log('Quarterly Summary complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
