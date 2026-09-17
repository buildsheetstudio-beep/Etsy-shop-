'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Multi-Year Comparison'];
const S  = "'Multi-Year Comparison'";

const YEARS = [2023, 2024, 2025, 2026];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 200 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Column widths
  [[0,220],[1,130],[2,130],[3,130],[4,130],[5,130],[6,130]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:G1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedGold),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 18 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Column header row 2
  reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, 0, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  function sectionHdr(r, color) {
    reqs.push({ mergeCells: { range: gridRange(SH, r, r+1, 0, 7), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 7), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function dataBlock(r1, r2) {
    for (let r = r1; r < r2; r++) {
      const bg = r % 2 === 0 ? C.panel : C.altRow;
      reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 7), cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    }
  }
  function totalRow(r) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 7), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }

  // Main table rows 3-16
  dataBlock(2, 17);

  // Monthly trend section
  sectionHdr(17, C.secondary);
  reqs.push({ repeatCell: { range: gridRange(SH, 18, 19, 0, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
  dataBlock(19, 32);
  totalRow(32);

  // Currency format
  [1,2,3,4,5,6].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 34, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Percent format for margin and growth rows
  [1,2,3,4,5,6].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 6, 10, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 1, 34, 0, 7),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze row 2
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'multiyear-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['📈 Multi-Year Comparison — 2023 to 2026']] });
  data.push({ range: `${S}!A2:G2`, values: [['Metric','2023','2024','2025','2026','YoY Growth (2025→2026)','3-Year CAGR']] });

  // Historical data (static for 2023-2025, formula for 2026)
  const historicalRevenue    = [32400, 48750, 67200];
  const historicalExpenses   = [12800, 18300, 22400];
  const historicalOrders     = [290, 425, 590];
  const historicalAvgOrder   = [111.72, 114.71, 113.90];

  // 2026 live formulas
  const rev26   = `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=2026)*'Income & Orders'!$H$6:$H$500)`;
  const exp26   = `=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=2026)*'Expense Log'!$H$6:$H$500)`;
  const ord26   = `=COUNTIFS(YEAR('Income & Orders'!$B$6:$B$500),2026,'Income & Orders'!$P$6:$P$500,"Received")`;
  const avgord26 = `=IFERROR(E3/E7,0)`;

  data.push({ range: `${S}!A3:G3`, values: [['Gross Revenue',...historicalRevenue,rev26,'=IFERROR((E3-D3)/D3,0)','=IFERROR((E3/B3)^(1/3)-1,0)']] });
  data.push({ range: `${S}!A4:G4`, values: [['Total Expenses',...historicalExpenses,exp26,'=IFERROR((E4-D4)/D4,0)','=IFERROR((E4/B4)^(1/3)-1,0)']] });
  data.push({ range: `${S}!A5:G5`, values: [['Net Profit','=B3-B4','=C3-C4','=D3-D4','=E3-E4','=IFERROR((E5-D5)/D5,0)','=IFERROR((E5/B5)^(1/3)-1,0)']] });
  data.push({ range: `${S}!A6:G6`, values: [['Profit Margin %','=IFERROR(B5/B3,0)','=IFERROR(C5/C3,0)','=IFERROR(D5/D3,0)','=IFERROR(E5/E3,0)','=IFERROR(E6-D6,0)','']] });
  data.push({ range: `${S}!A7:G7`, values: [['Order Count',...historicalOrders,ord26,'=IFERROR((E7-D7)/D7,0)','=IFERROR((E7/B7)^(1/3)-1,0)']] });
  data.push({ range: `${S}!A8:G8`, values: [['Avg Order Value',...historicalAvgOrder,'=IFERROR(E3/E7,0)','=IFERROR((E8-D8)/D8,0)','']] });
  data.push({ range: `${S}!A9:G9`, values: [['Revenue per Listing',3240,4875,6720,'=IFERROR(E3/20,0)','=IFERROR((E9-D9)/D9,0)','']] });
  data.push({ range: `${S}!A10:G10`, values: [['Etsy Fee % of Revenue',0.148,0.142,0.138,'=IFERROR(SUMPRODUCT((YEAR(\'Expense Log\'!$A$6:$A$500)=2026)*(ISNUMBER(MATCH(\'Expense Log\'!$E$6:$E$500,{"Etsy Transaction Fees","Etsy Listing Fees","Etsy Payment Processing Fees","Etsy Ads"},0)))*\'Expense Log\'!$H$6:$H$500)/E3,0)','=IFERROR(E10-D10,0)','']] });

  // Growth narrative row
  data.push({ range: `${S}!A12`, values: [['  Growth Highlights']] });
  data.push({ range: `${S}!A13:G13`, values: [['Revenue Growth 2023→2026','','','','','=IFERROR((E3-B3)/B3,0)','']] });
  data.push({ range: `${S}!A14:G14`, values: [['Expense Growth 2023→2026','','','','','=IFERROR((E4-B4)/B4,0)','']] });
  data.push({ range: `${S}!A15:G15`, values: [['Profit Growth 2023→2026','','','','','=IFERROR((E5-B5)/B5,0)','']] });
  data.push({ range: `${S}!A16:G16`, values: [['Best Year (Revenue)','=IF(MAX(B3,C3,D3,E3)=B3,"2023",IF(MAX(B3,C3,D3,E3)=C3,"2024",IF(MAX(B3,C3,D3,E3)=D3,"2025","2026")))','','','','','']] });
  data.push({ range: `${S}!A17:G17`, values: [['  📅 Revenue by Month — Year over Year']] });
  data.push({ range: `${S}!A18:G18`, values: [['Month','2023','2024','2025','2026','MoM Growth','vs 2025']] });

  const monthlyRev23 = [2100,1900,2800,2600,3100,2400,2800,2700,3200,3100,2800,3900];
  const monthlyRev24 = [3200,2900,4100,3900,4600,3700,4100,4000,4700,4700,4300,5500];
  const monthlyRev25 = [4500,4000,5600,5300,6400,5100,5600,5500,6500,6400,5800,7000];

  MONTHS.forEach((m, mi) => {
    const r = 19 + mi;
    data.push({ range: `${S}!A${r}:G${r}`, values: [[
      m,
      monthlyRev23[mi],
      monthlyRev24[mi],
      monthlyRev25[mi],
      `=SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=${mi+1})*(YEAR('Income & Orders'!$B$6:$B$500)=2026)*'Income & Orders'!$H$6:$H$500)`,
      `=IFERROR((E${r}-E${r-1})/E${r-1},0)`,
      `=IFERROR((E${r}-D${r})/D${r},0)`,
    ]] });
  });
  data.push({ range: `${S}!A32:G32`, values: [['TOTAL','=SUM(B19:B31)','=SUM(C19:C31)','=SUM(D19:D31)','=SUM(E19:E31)','=IFERROR((E32-D32)/D32,0)','=IFERROR((E32-D32)/D32,0)']] });

  await valuesBatchUpdate(id, data, 'multiyear-values');
  console.log('Multi-Year Comparison complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
