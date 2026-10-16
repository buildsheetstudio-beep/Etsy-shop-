'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS, EXPENSE_CATS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Monthly Dashboard'];
const S  = "'Monthly Dashboard'";

// Monthly Dashboard: pick a month (B1), see full breakdown
// Rows: 1=title, 2=filter, 3-6=KPI cards, 7=blank, 8=income section hdr,
//       9-20=income by source, 21=blank, 22=expense section hdr, 23-52=expense by cat,
//       53=blank, 54=net profit section, 55-57=summary

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 2, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 7, endIndex: 100 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,40],[1,200],[2,140],[3,140],[4,140],[5,140],[6,140],[7,140],[8,140]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:I1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Filter row A2:I2
  reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 11 },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  // Filter input cell C2
  reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, 2, 4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
    textFormat: { foregroundColor: hex(C.primary), bold: true, fontSize: 12 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, 2, 4), mergeType: 'MERGE_ALL' } });

  // KPI cards rows 3-6 (indices 2-5) — 4 cards, 2 cols each
  const kpiColors = [C.success, C.primary, C.attention, C.warning];
  [[0,2],[2,4],[5,7],[7,9]].forEach(([c0,c1], idx) => {
    reqs.push({ mergeCells: { range: gridRange(SH, 2, 3, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(kpiColors[idx]),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 3, 5, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 3, 5, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { foregroundColor: hex(kpiColors[idx]), bold: true, fontSize: 22 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 5, 6, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 6, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });

  // KPI card 5 (col 4, center) — Orders
  reqs.push({ mergeCells: { range: gridRange(SH, 2, 3, 4, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, 4, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ mergeCells: { range: gridRange(SH, 3, 5, 4, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 3, 5, 4, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.panel),
    textFormat: { foregroundColor: hex(C.mutedBlue), bold: true, fontSize: 22 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ mergeCells: { range: gridRange(SH, 5, 6, 4, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 5, 6, 4, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Section headers & data table style
  function sectionHdr(r, c0, c1, color) {
    reqs.push({ mergeCells: { range: gridRange(SH, r, r+1, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function tableHdr(r, c0, c1) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function dataRows(r1, r2, c0, c1) {
    for (let r = r1; r < r2; r++) {
      const bg = r % 2 === 0 ? C.panel : C.altRow;
      reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, c0, c1), cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    }
  }

  // Income section: row 7 hdr, 8 col hdr, 9-19 data, 20 total
  sectionHdr(7, 0, 9, C.success);
  tableHdr(8, 0, 9);
  dataRows(9, 20, 0, 9);
  // Income total row
  reqs.push({ repeatCell: { range: gridRange(SH, 20, 21, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.success),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Expense section: row 22 hdr, 23 col hdr, 24-53 data, 54 total
  sectionHdr(22, 0, 9, C.attention);
  tableHdr(23, 0, 9);
  dataRows(24, 54, 0, 9);
  reqs.push({ repeatCell: { range: gridRange(SH, 54, 55, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.attention),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Net Profit section row 56
  sectionHdr(56, 0, 9, C.primary);
  dataRows(57, 62, 0, 9);

  // Currency format for cols C-I (2-8) in data ranges
  [2,3,4,5,6,7,8].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 9, 62, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
    reqs.push({ repeatCell: { range: gridRange(SH, 3, 5, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Percent format for margin col I (8) in summary
  reqs.push({ repeatCell: { range: gridRange(SH, 58, 62, 8, 9), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Borders
  [gridRange(SH, 7, 22, 0, 9), gridRange(SH, 22, 56, 0, 9), gridRange(SH, 56, 62, 0, 9)].forEach(range => {
    reqs.push({ updateBorders: { range,
      innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
      innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
      bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
    }});
  });

  // Freeze rows 2
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'monthly-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['📅 Monthly Dashboard — Income & Expense Breakdown']] });
  data.push({ range: `${S}!A2`, values: [['  View Month ▶']] });
  data.push({ range: `${S}!C2`, values: [['July']] }); // default to current month

  // KPI cards labels + formulas
  const mo = `C$2`; // month name in C2
  const yr = 2026;

  data.push({ range: `${S}!A3`, values: [['  Gross Revenue']] });
  data.push({ range: `${S}!A4`, values: [[`=SUMPRODUCT((TEXT('Income & Orders'!$B$6:$B$500,"mmmm")=${mo})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)`]] });
  data.push({ range: `${S}!A6`, values: [['vs prior month']] });

  data.push({ range: `${S}!C3`, values: [['  Total Expenses']] });
  data.push({ range: `${S}!C4`, values: [[`=SUMPRODUCT((TEXT('Expense Log'!$A$6:$A$500,"mmmm")=${mo})*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*'Expense Log'!$H$6:$H$500)`]] });
  data.push({ range: `${S}!C6`, values: [['operating costs']] });

  data.push({ range: `${S}!E3`, values: [['  Orders']] });
  data.push({ range: `${S}!E4`, values: [[`=COUNTIFS(TEXT('Income & Orders'!$B$6:$B$500,"mmmm"),${mo},YEAR('Income & Orders'!$B$6:$B$500),${yr},'Income & Orders'!$P$6:$P$500,"Received")`]] });
  data.push({ range: `${S}!E6`, values: [['completed orders']] });

  data.push({ range: `${S}!F3`, values: [['  Net Profit']] });
  data.push({ range: `${S}!F4`, values: [['=A4-C4']] });
  data.push({ range: `${S}!F6`, values: [['revenue minus expenses']] });

  data.push({ range: `${S}!H3`, values: [['  Profit Margin']] });
  data.push({ range: `${S}!H4`, values: [['=IFERROR(F4/A4,0)']] });
  data.push({ range: `${S}!H6`, values: [['goal: 70%+']] });

  // Income section
  data.push({ range: `${S}!A8`, values: [['  💰 Income Breakdown by Source']] });
  data.push({ range: `${S}!A9:I9`, values: [['#','Source','This Month','Last Month','3-Month Avg','YTD','Goal','vs Goal','% of Revenue']] });

  const SOURCES = ['Etsy Order','Etsy Shipping Collected','Off-Etsy Sale','Custom Order','Licensing','Affiliate Income','Other Business Income'];
  SOURCES.forEach((src, i) => {
    const r = 10 + i;
    data.push({ range: `${S}!A${r}:I${r}`, values: [[
      i+1, src,
      `=SUMPRODUCT((TEXT('Income & Orders'!$B$6:$B$500,"mmmm")=${mo})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$F$6:$F$500="${src}")*'Income & Orders'!$H$6:$H$500)`,
      `=SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=MONTH(DATEVALUE(${mo}&" 1")-1)*(YEAR('Income & Orders'!$B$6:$B$500)=${yr}))*('Income & Orders'!$F$6:$F$500="${src}")*'Income & Orders'!$H$6:$H$500)`,
      `=IFERROR(AVERAGE(C${r},D${r}),0)`,
      `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$F$6:$F$500="${src}")*'Income & Orders'!$H$6:$H$500)`,
      '',
      `=IFERROR(F${r}-G${r},0)`,
      `=IFERROR(C${r}/C$21,0)`,
    ]] });
  });
  // Income total row 21
  data.push({ range: `${S}!A21:I21`, values: [['','TOTAL',
    '=SUM(C10:C17)', '=SUM(D10:D17)', '=IFERROR(AVERAGE(E10:E17),0)', '=SUM(F10:F17)',
    '=Business Setup\'!B14', '=F21-G21', '100.0%',
  ]] });

  // Expense section
  data.push({ range: `${S}!A23`, values: [['  💸 Expense Breakdown by Category']] });
  data.push({ range: `${S}!A24:I24`, values: [['#','Category','This Month','Last Month','3-Month Avg','YTD','Tax Deductible?','Tax Group','% of Expenses']] });

  EXPENSE_CATS.slice(0, 28).forEach((cat, i) => {
    const r = 25 + i;
    data.push({ range: `${S}!A${r}:I${r}`, values: [[
      i+1, cat,
      `=SUMPRODUCT((TEXT('Expense Log'!$A$6:$A$500,"mmmm")=${mo})*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*'Expense Log'!$H$6:$H$500)`,
      `=SUMPRODUCT((MONTH('Expense Log'!$A$6:$A$500)=IFERROR(MONTH(DATEVALUE(${mo}&" 1"))-1,12))*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*'Expense Log'!$H$6:$H$500)`,
      `=IFERROR(AVERAGE(C${r},D${r}),0)`,
      `=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*'Expense Log'!$H$6:$H$500)`,
      'Yes',
      `=IFERROR(VLOOKUP(B${r},'Reference Data'!$N$2:$O${1+EXPENSE_CATS.length},2,FALSE),"")`,
      `=IFERROR(C${r}/C$55,0)`,
    ]] });
  });
  // Expense total row 55
  data.push({ range: `${S}!A55:I55`, values: [['','TOTAL',
    '=SUM(C25:C53)', '=SUM(D25:D53)', '=IFERROR(AVERAGE(E25:E53),0)', '=SUM(F25:F53)',
    '', '', '100.0%',
  ]] });

  // Net Profit section
  data.push({ range: `${S}!A57`, values: [['  📊 Profit Summary']] });
  data.push({ range: `${S}!A58:I58`, values: [['','Metric','This Month','vs Goal','YTD','YTD Goal','YTD Variance','','Margin %']] });
  data.push({ range: `${S}!A59:I59`, values: [['','Net Revenue','=A4','=A4-Business Setup\'!B16','=F21','=Business Setup\'!B13','=E59-F59','','=IFERROR(C59/C21,0)']] });
  data.push({ range: `${S}!A60:I60`, values: [['','Total Expenses','=C4','','=F55','','','','=IFERROR(C60/C21,0)']] });
  data.push({ range: `${S}!A61:I61`, values: [['','Net Profit','=F4','=F4-Business Setup\'!B17','=E59-E60','=Business Setup\'!B14','=E61-F61','','=IFERROR(C61/C59,0)']] });
  data.push({ range: `${S}!A62:I62`, values: [['','Etsy Fees (Total)','=SUMPRODUCT((TEXT(\'Expense Log\'!$A$6:$A$500,"mmmm")=C$2)*(YEAR(\'Expense Log\'!$A$6:$A$500)=2026)*(ISNUMBER(MATCH(\'Expense Log\'!$E$6:$E$500,{"Etsy Transaction Fees","Etsy Listing Fees","Etsy Payment Processing Fees","Etsy Ads","Offsite Ads"},0)))*\'Expense Log\'!$H$6:$H$500)','','=SUMPRODUCT((YEAR(\'Expense Log\'!$A$6:$A$500)=2026)*(ISNUMBER(MATCH(\'Expense Log\'!$E$6:$E$500,{"Etsy Transaction Fees","Etsy Listing Fees","Etsy Payment Processing Fees","Etsy Ads","Offsite Ads"},0)))*\'Expense Log\'!$H$6:$H$500)','','','','=IFERROR(C62/C59,0)']] });

  await valuesBatchUpdate(id, data, 'monthly-values');

  // Dropdown for month filter C2
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 1, 2, 2, 4),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$J$2:$J$13` }] }, showCustomUi: true, strict: true },
  }});
  await batchUpdate(id, dvReqs, 'monthly-validation');

  // CF for KPI margin — red if <60%, green if >70%
  const cfReqs = [];
  cfReqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SH, 3, 5, 7, 8)],
    booleanRule: { condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0.6' }] }, format: { textFormat: { foregroundColor: hex(C.attention) } } },
  }, index: 0 }});
  cfReqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SH, 3, 5, 7, 8)],
    booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.7' }] }, format: { textFormat: { foregroundColor: hex(C.success) } } },
  }, index: 1 }});
  await batchUpdate(id, cfReqs, 'monthly-cf');

  console.log('Monthly Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
