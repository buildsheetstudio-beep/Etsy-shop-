'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Etsy Fee Analysis'];
const S  = "'Etsy Fee Analysis'";

(async () => {
  const reqs = [];
  const yr = 2026;

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 200 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Column widths
  [[0,210],[1,110],[2,110],[3,110],[4,110],[5,110],[6,110],[7,100]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:H1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 17 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // KPI cards row 2-4 (indices 1-3) — 4 cards
  [[0,2],[2,4],[4,6],[6,8]].forEach(([c0,c1], idx) => {
    const colors = [C.attention, C.warning, C.primary, C.mutedBlue];
    reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(colors[idx]),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 2, 4, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 4, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { foregroundColor: hex(colors[idx]), bold: true, fontSize: 20 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 4, 5, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 4, 5, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });

  // Section header helper
  function sectionHdr(r, color) {
    reqs.push({ mergeCells: { range: gridRange(SH, r, r+1, 0, 8), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
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
        textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    }
  }

  // Fee breakdown section: row 6 hdr, 7 col hdr, 8-16 data, 17 total
  sectionHdr(5, C.attention);
  colHdr(6);
  dataBlock(7, 17);
  reqs.push({ repeatCell: { range: gridRange(SH, 17, 18, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.attention),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Monthly fee trend section: row 19 hdr, 20 col hdr, 21-33 data, 34 total
  sectionHdr(18, C.secondary);
  colHdr(19);
  dataBlock(20, 33);
  reqs.push({ repeatCell: { range: gridRange(SH, 33, 34, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Currency KPI cards
  [0,1,2,3].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 4, c*2, c*2+2), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Currency data
  [1,2,3,4,5,6,7].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 6, 35, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Percent format col G (fee rate)
  reqs.push({ repeatCell: { range: gridRange(SH, 7, 18, 6, 7), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.00%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(SH, 20, 34, 6, 7), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.00%' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 5, 35, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze row 2
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'fees-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['💳 Etsy Fee Analysis — Fee Breakdown & Optimization']] });

  // KPI cards
  const totalFeeFormula = `=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*(ISNUMBER(MATCH('Expense Log'!$E$6:$E$500,{"Etsy Transaction Fees","Etsy Listing Fees","Etsy Payment Processing Fees","Etsy Ads","Offsite Ads"},0)))*'Expense Log'!$H$6:$H$500)`;
  const grossRevFormula = `=SUMPRODUCT((YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)`;

  data.push({ range: `${S}!A2`, values: [['  Total Etsy Fees (2026)']] });
  data.push({ range: `${S}!A3`, values: [[totalFeeFormula]] });
  data.push({ range: `${S}!A5`, values: [['Etsy platform costs YTD']] });

  data.push({ range: `${S}!C2`, values: [['  Fees as % of Revenue']] });
  data.push({ range: `${S}!C3`, values: [[`=IFERROR(${totalFeeFormula}/${grossRevFormula},0)`]] });
  data.push({ range: `${S}!C5`, values: [['aim: <15% of gross revenue']] });

  data.push({ range: `${S}!E2`, values: [['  Transaction Fees (2026)']] });
  data.push({ range: `${S}!E3`, values: [[`=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="Etsy Transaction Fees")*'Expense Log'!$H$6:$H$500)`]] });
  data.push({ range: `${S}!E5`, values: [['6.5% per transaction']] });

  data.push({ range: `${S}!G2`, values: [['  Listing Fees (2026)']] });
  data.push({ range: `${S}!G3`, values: [[`=SUMPRODUCT((YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="Etsy Listing Fees")*'Expense Log'!$H$6:$H$500)`]] });
  data.push({ range: `${S}!G5`, values: [['$0.20 per listing per 4 months']] });

  // Fee breakdown section
  data.push({ range: `${S}!A6`, values: [['  📊 Fee Breakdown by Type']] });
  data.push({ range: `${S}!A7:H7`, values: [['Fee Type','Jan–Mar','Apr–Jun','Jul–Sep','Oct–Dec','Full Year','% of Revenue','vs Prior Year']] });

  const FEE_ROWS = [
    ['Etsy Transaction Fees',   'Etsy Transaction Fees'],
    ['Etsy Listing Fees',       'Etsy Listing Fees'],
    ['Payment Processing Fees', 'Etsy Payment Processing Fees'],
    ['Etsy Ads',                'Etsy Ads'],
    ['Offsite Ads',             'Offsite Ads'],
  ];
  const Q_MONTHS = { Q1:[1,2,3], Q2:[4,5,6], Q3:[7,8,9], Q4:[10,11,12] };

  FEE_ROWS.forEach(([label, cat], i) => {
    const r = 8 + i;
    const qCols = ['Q1','Q2','Q3','Q4'].map(q => {
      const months = Q_MONTHS[q];
      return `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Expense Log'!$A$6:$A$500),{${months.join(',')}},0)))*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*'Expense Log'!$H$6:$H$500)`;
    });
    data.push({ range: `${S}!A${r}:H${r}`, values: [[
      label, ...qCols,
      `=SUM(B${r}:E${r})`,
      `=IFERROR(F${r}/${grossRevFormula},0)`,
      '',
    ]] });
  });

  // Sub-total rows for non-deductible items
  data.push({ range: `${S}!A13:H13`, values: [['Regulatory Op Fee', 0, 0, 0, 0, 0, '=IFERROR(F13/'+grossRevFormula.slice(1)+',0)', '']] });
  data.push({ range: `${S}!A14:H14`, values: [['Currency Conversion', 0, 0, 0, 0, 0, '', '']] });
  data.push({ range: `${S}!A15:H15`, values: [['Refunds Issued',
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{1,2,3},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$M$6:$M$500<0)*'Income & Orders'!$M$6:$M$500)`,
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{4,5,6},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$M$6:$M$500<0)*'Income & Orders'!$M$6:$M$500)`,
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{7,8,9},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$M$6:$M$500<0)*'Income & Orders'!$M$6:$M$500)`,
    `=SUMPRODUCT((ISNUMBER(MATCH(MONTH('Income & Orders'!$B$6:$B$500),{10,11,12},0)))*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$M$6:$M$500<0)*'Income & Orders'!$M$6:$M$500)`,
    '=SUM(B15:E15)','','',
  ]] });
  data.push({ range: `${S}!A16:H16`, values: [['Net Listing Credits', 0, 0, 0, 0, 0, '', '']] });
  data.push({ range: `${S}!A17:H17`, values: [['Etsy Credits Used', 0, 0, 0, 0, 0, '', '']] });

  // Fee total row
  data.push({ range: `${S}!A18:H18`, values: [[
    'TOTAL FEES',
    '=SUM(B8:B17)','=SUM(C8:C17)','=SUM(D8:D17)','=SUM(E8:E17)',
    '=SUM(F8:F17)',
    `=IFERROR(F18/${grossRevFormula},0)`,
    '',
  ]] });

  // Monthly fee trend
  data.push({ range: `${S}!A19`, values: [['  📅 Monthly Fee Trend (2026)']] });
  data.push({ range: `${S}!A20:H20`, values: [['Month','Trans. Fee','Listing Fee','Pay Process.','Etsy Ads','Offsite Ads','Total Fees','Fee Rate %']] });

  MONTHS.forEach((m, mi) => {
    const r = 21 + mi;
    const mo = mi + 1;
    const feeTypes = ['Etsy Transaction Fees','Etsy Listing Fees','Etsy Payment Processing Fees','Etsy Ads','Offsite Ads'];
    const feeCols = feeTypes.map(cat =>
      `=SUMPRODUCT((MONTH('Expense Log'!$A$6:$A$500)=${mo})*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*'Expense Log'!$H$6:$H$500)`
    );
    const revMo = `SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=${mo})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)`;
    data.push({ range: `${S}!A${r}:H${r}`, values: [[
      m, ...feeCols,
      `=SUM(B${r}:F${r})`,
      `=IFERROR(G${r}/${revMo},0)`,
    ]] });
  });
  data.push({ range: `${S}!A34:H34`, values: [['TOTAL','=SUM(B21:B33)','=SUM(C21:C33)','=SUM(D21:D33)','=SUM(E21:E33)','=SUM(F21:F33)','=SUM(G21:G33)',`=IFERROR(G34/${grossRevFormula},0)`]] });

  await valuesBatchUpdate(id, data, 'fees-values');

  // Percent format for fees as % card
  const pvReqs = [];
  pvReqs.push({ repeatCell: { range: gridRange(SH, 2, 4, 2, 4), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  await batchUpdate(id, pvReqs, 'fees-pct-format');

  console.log('Etsy Fee Analysis complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
