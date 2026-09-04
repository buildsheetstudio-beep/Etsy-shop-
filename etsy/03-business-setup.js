'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, SHOP_TYPES, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BS = sheetMap['Business Setup'];
const S  = "'Business Setup'";

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: BS, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BS, dimension: 'ROWS', startIndex: 1, endIndex: 200 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,240],[1,280],[2,200],[3,200],[4,160],[5,160],[6,160],[7,160]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: BS, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row A1:H1
  reqs.push({ mergeCells: { range: gridRange(BS, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BS, 0, 1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Section header style helper
  function sectionHeader(r) {
    reqs.push({ mergeCells: { range: gridRange(BS, r, r+1, 0, 8), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(BS, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }

  // Label/input style helpers
  function labelRow(r, cols) {
    reqs.push({ repeatCell: { range: gridRange(BS, r, r+1, 0, cols), cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 10 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }
  function inputRows(r1, r2, c1, c2) {
    for (let r = r1; r < r2; r++) {
      reqs.push({ repeatCell: { range: gridRange(BS, r, r+1, c1, c2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
      reqs.push({ repeatCell: { range: gridRange(BS, r, r+1, 0, c1), cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    }
  }

  // Section 1: Shop Profile (rows 2-12)
  sectionHeader(1);
  inputRows(2, 12, 1, 4);

  // Section 2: Financial Goals (rows 13-20)
  sectionHeader(12);
  inputRows(13, 21, 1, 4);

  // Section 3: Fee Reference (rows 22-30)
  sectionHeader(21);
  for (let r = 22; r < 30; r++) {
    const bg = r % 2 === 0 ? C.bg : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(BS, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Section 4: Annual Goals by Month (rows 31-45)
  sectionHeader(30);
  // header row
  reqs.push({ repeatCell: { range: gridRange(BS, 31, 32, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  // Month data rows
  for (let r = 32; r < 45; r++) {
    const bg = r % 2 === 0 ? C.bg : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(BS, r, r+1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)' }});
  }

  // Currency format for fee reference and goals
  [1,2,3,4,5,6,7].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(BS, 22, 30, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.00%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Goals currency
  [1,2,3,4,5,6,7].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(BS, 13, 21, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  [2,3,4,5,6,7].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(BS, 32, 45, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Borders
  reqs.push({ updateBorders: { range: gridRange(BS, 1, 45, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze row 1
  reqs.push({ updateSheetProperties: { properties: { sheetId: BS, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'bs-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['🌿 Business Setup — Willow Paper Studio']] });

  data.push({ range: `${S}!A2`, values: [['  🏪 Shop Profile']] });

  const profileLabels = [
    ['Shop Name', 'Willow Paper Studio'],
    ['Owner Name', 'Emma Willow'],
    ['Etsy Shop URL', 'https://www.etsy.com/shop/WillowPaperStudio'],
    ['Shop Type', 'Digital Products'],
    ['Business Year Start', 'January'],
    ['Tax Year', '2026'],
    ['Home Currency', 'USD'],
    ['Etsy Seller Since', '2021'],
    ['Primary Niche', 'Printable Planners, Wall Art & Stationery'],
    ['Secondary Niche', 'Print on Demand Art Prints'],
  ];
  profileLabels.forEach(([label, val], i) => {
    data.push({ range: `${S}!A${3+i}:B${3+i}`, values: [[label, val]] });
  });

  data.push({ range: `${S}!A13`, values: [['  💰 Financial Goals']] });
  const goals = [
    ['Annual Revenue Goal', 85000],
    ['Annual Profit Goal', 60000],
    ['Monthly Revenue Target', 7083.33],
    ['Monthly Profit Target', 5000],
    ['Target Profit Margin', 0.706],
    ['Etsy Ads Monthly Budget', 150],
    ['Emergency Fund Target', 10000],
    ['Reinvestment % of Profit', 0.15],
  ];
  goals.forEach(([label, val], i) => {
    data.push({ range: `${S}!A${14+i}:B${14+i}`, values: [[label, val]] });
  });

  data.push({ range: `${S}!A22`, values: [['  📊 Etsy Fee Reference (Current Rates)']] });
  const fees = [
    ['Transaction Fee', 0.065, '', 'Applied to item price + shipping'],
    ['Listing Fee (per listing per 4 months)', 0.20, '', 'Fixed USD amount'],
    ['Payment Processing — US', 0.03, 0.25, '3% + $0.25 per transaction'],
    ['Offsite Ads Fee (>$10K sales/yr)', 0.12, '', 'If opted out: N/A'],
    ['Offsite Ads Fee (<$10K sales/yr)', 0.15, '', 'Optional — can opt out'],
    ['Regulatory Operating Fee (some regions)', 0.0045, '', '0.45% on certain sales'],
    ['Currency Conversion Fee', 0.025, '', '2.5% if currency mismatch'],
    ['VAT/Sales Tax', 0, '', 'Collected & remitted by Etsy'],
  ];
  fees.forEach(([label, rate, fixed, note], i) => {
    data.push({ range: `${S}!A${23+i}:D${23+i}`, values: [[label, rate, fixed||'', note||'']] });
  });

  data.push({ range: `${S}!A31`, values: [['  📅 Monthly Revenue & Profit Goals — 2026']] });
  data.push({ range: `${S}!A32:H32`, values: [['Month','Revenue Goal','Profit Goal','Revenue Actual','Profit Actual','Variance $','Variance %','Notes']] });
  MONTHS.forEach((m, i) => {
    data.push({ range: `${S}!A${33+i}:H${33+i}`, values: [[
      m,
      7083.33,
      5000,
      `=IFERROR(SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=${i+1})*(YEAR('Income & Orders'!$B$6:$B$500)=2026)*('Income & Orders'!$H$6:$H$500)),0)`,
      `=IFERROR(D${33+i}-SUMPRODUCT((MONTH('Expense Log'!$B$6:$B$500)=${i+1})*(YEAR('Expense Log'!$B$6:$B$500)=2026)*('Expense Log'!$L$6:$L$500)),0)`,
      `=D${33+i}-B${33+i}`,
      `=IFERROR((D${33+i}-B${33+i})/B${33+i},0)`,
      '',
    ]] });
  });

  await valuesBatchUpdate(id, data, 'bs-values');

  // Data validation for Shop Type
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(BS, 5, 6, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$A$2:$A${1+SHOP_TYPES.length}` }] }, showCustomUi: true, strict: true },
  }});
  dvReqs.push({ setDataValidation: {
    range: gridRange(BS, 10, 11, 1, 2),
    rule: { condition: { type: 'ONE_OF_LIST', values: MONTHS.map(m => ({ userEnteredValue: m })) }, showCustomUi: true, strict: true },
  }});
  await batchUpdate(id, dvReqs, 'bs-validation');

  console.log('Business Setup complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
