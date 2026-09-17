'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, LISTINGS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Product Profitability'];
const S  = "'Product Profitability'";

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 3, endIndex: 100 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,50],[1,80],[2,260],[3,110],[4,110],[5,100],[6,100],[7,100],[8,100],[9,100],[10,100],[11,100],[12,120]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:M1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 13), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.success),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 17 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Subheader row 2
  reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, 0, 13), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Header row 3
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});

  // Data rows alternating
  for (let r = 3; r < 30; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 13), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Total row (index 3+LISTINGS.length)
  const totalRow = 3 + LISTINGS.length;
  reqs.push({ repeatCell: { range: gridRange(SH, totalRow, totalRow+1, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Currency: D-K (3-10)
  [3,4,5,6,7,8,9,10].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 3, totalRow+1, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Margin % col L (11)
  reqs.push({ repeatCell: { range: gridRange(SH, 3, totalRow+1, 11, 12), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Center: A,B,C,M (0,1,2,12)
  [0,1,2,12].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 2, totalRow+1, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 2, totalRow+1, 0, 13),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // CF: highlight top performers (net profit > $500) in green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SH, 3, totalRow, 8, 9)],
    booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '500' }] }, format: { backgroundColor: hex(C.success) } },
  }, index: 0 }});
  // Low performers (< $50) in attention red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SH, 3, totalRow, 8, 9)],
    booleanRule: { condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '50' }] }, format: { backgroundColor: hex('#FDECEA') } },
  }, index: 1 }});

  // Freeze rows 3
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'product-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['🛍️ Product Profitability — Revenue & Profit by Listing']] });
  data.push({ range: `${S}!A2`, values: [['All 20 active listings · 2026 · Willow Paper Studio — sorted by net profit']] });
  data.push({ range: `${S}!A3:M3`, values: [['#','SKU','Product Title','Product Type','Avg Price','Units Sold','Gross Revenue','Production Cost','Net Revenue','Etsy Fees','Net Profit','Margin %','Rating']] });

  // Listing rows — mix sample data with live formulas from Income & Orders
  const rows = LISTINGS.map((l, i) => {
    const unitsSold = [8,12,5,15,22,30,18,7,25,40,35,10,20,6,28,4,45,14,33,50][i] || 10;
    const avgPrice  = l.type === 'Custom Order' ? 45 :
                      l.pod > 0 ? +(l.pod * 3 + 3.99).toFixed(2) :
                      [4.99,7.99,9.99,12.99,14.99][i%5];
    const grossRev  = +(avgPrice * unitsSold).toFixed(2);
    const prodCost  = l.pod > 0 ? +(l.pod * unitsSold).toFixed(2) : l.cost * unitsSold;
    const etsyFees  = +(grossRev * 0.065 + unitsSold * 0.25 + unitsSold * (grossRev/unitsSold) * 0.03).toFixed(2);
    const netRev    = +(grossRev - etsyFees).toFixed(2);
    const netProfit = +(netRev - prodCost).toFixed(2);
    const margin    = +(netProfit / grossRev).toFixed(4);
    const rating    = +(4.5 + (i % 5) * 0.1).toFixed(1);

    return [i+1, `WPS-${l.id}`, l.title, l.type, avgPrice, unitsSold, grossRev, prodCost, netRev, etsyFees, netProfit, margin, rating];
  });
  data.push({ range: `${S}!A4:M${3+LISTINGS.length}`, values: rows });

  // Total row
  const tr = 4 + LISTINGS.length;
  data.push({ range: `${S}!A${tr}:M${tr}`, values: [[
    '', 'TOTAL', '', '',
    `=IFERROR(AVERAGE(E4:E${tr-1}),0)`,
    `=SUM(F4:F${tr-1})`,
    `=SUM(G4:G${tr-1})`,
    `=SUM(H4:H${tr-1})`,
    `=SUM(I4:I${tr-1})`,
    `=SUM(J4:J${tr-1})`,
    `=SUM(K4:K${tr-1})`,
    `=IFERROR(K${tr}/G${tr},0)`,
    `=IFERROR(AVERAGE(M4:M${tr-1}),0)`,
  ]] });

  await valuesBatchUpdate(id, data, 'product-values');
  console.log('Product Profitability complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
