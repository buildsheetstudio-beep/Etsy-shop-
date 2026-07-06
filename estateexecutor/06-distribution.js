'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DIST = sheetMap['🏦 Estate Distribution'];

// Cols: A=# B=Item/Description C=Beneficiary D=Type E=Will Clause F=Est Value G=Amount Paid H=Date Paid I=Status J=Notes
const ENTRIES = [
  [1, 'Residuary Estate Share (40%) — after debts & expenses', 'Robert James Williams', 'Residuary', 'Clause 6(a)', 818216, 0, '', 'Pending', 'First interim distribution expected Q4 2025'],
  [2, 'Residuary Estate Share (30%) — after debts & expenses', 'Catherine Anne Williams', 'Residuary', 'Clause 6(b)', 613662, 0, '', 'Pending', 'To be paid by bank transfer'],
  [3, 'Residuary Estate Share (20%) — after debts & expenses', 'Thomas Edward Williams', 'Residuary', 'Clause 6(c)', 409108, 0, '', 'Pending', ''],
  [4, 'Charitable Bequest (10%) — Margaret Grace Williams Charitable Trust', 'Margaret Grace Williams Charitable Trust', 'Pecuniary Bequest', 'Clause 8', 204554, 0, '', 'Pending', 'Arts & community grants — per Will directions'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DIST, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  reqs.push({ updateDimensionProperties: { range: { sheetId: DIST, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DIST, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DIST, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DIST, dimension: 'ROWS', startIndex: 3, endIndex: 30 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Cols: A=40 B=300 C=220 D=130 E=90 F=130 G=130 H=110 I=110 J=250
  const colW = [40, 300, 220, 130, 90, 130, 130, 110, 110, 250];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: DIST, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title
  reqs.push({ mergeCells: { range: gridRange(DIST, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DIST, 0, 1, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Summary bar
  reqs.push({ mergeCells: { range: gridRange(DIST, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DIST, 1, 2, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Headers
  reqs.push({ repeatCell: { range: gridRange(DIST, 2, 3, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Data rows
  for (let r = 0; r < 4; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(DIST, r+3, r+4, 0, 10), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Totals row
  reqs.push({ repeatCell: { range: gridRange(DIST, 7, 8, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.paleSage),
      textFormat: { foregroundColor: hex(C.mutedSage), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { top: { style: 'SOLID', color: hex(C.mutedSage), width: 2 }, bottom: { style: 'SOLID', color: hex(C.mutedSage), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Per-beneficiary summary block (rows 9-13)
  reqs.push({ updateDimensionProperties: { range: { sheetId: DIST, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 16 }, fields: 'pixelSize' }});
  reqs.push({ mergeCells: { range: gridRange(DIST, 9, 10, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DIST, 9, 10, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      padding: { left: 10 },
    },
  }, fields: 'userEnteredFormat' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DIST, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  // Summary headers row 10
  reqs.push({ repeatCell: { range: gridRange(DIST, 10, 11, 0, 5), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});
  // Summary data rows 11-14
  for (let r = 0; r < 4; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(DIST, r+11, r+12, 0, 5), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Currency formats
  reqs.push({ repeatCell: { range: gridRange(DIST, 3, 8, 5, 7), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(DIST, 11, 15, 2, 5), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  // Borders
  reqs.push({ updateBorders: { range: gridRange(DIST, 2, 8, 0, 10),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});
  reqs.push({ updateBorders: { range: gridRange(DIST, 10, 15, 0, 5),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'dist-fmt');

  const vals = [];
  vals.push({ range: `'🏦 Estate Distribution'!A1`, values: [['🏦 ESTATE DISTRIBUTION — Estate of Margaret Eleanor Williams (Ref: PROB-2025-0442)']] });
  vals.push({ range: `'🏦 Estate Distribution'!A2`, values: [['Net Estate Approx. $2,045,540 | 4 Beneficiaries | Executor: Robert James Williams | Status: Distribution Pending']] });
  vals.push({ range: `'🏦 Estate Distribution'!A3:J3`, values: [['#', 'Item / Description', 'Beneficiary', 'Distribution Type', 'Will Clause', 'Est. Value ($)', 'Amount Paid ($)', 'Date Paid', 'Status', 'Notes']] });

  const distRows = ENTRIES.map(e => e);
  vals.push({ range: `'🏦 Estate Distribution'!A4`, values: distRows });

  vals.push({ range: `'🏦 Estate Distribution'!A8:J8`, values: [['TOTALS', '', '', '', '', '=SUM(F4:F7)', '=SUM(G4:G7)', '', '', '']] });

  // Per-beneficiary summary
  vals.push({ range: `'🏦 Estate Distribution'!A10`, values: [['📊 Per-Beneficiary Summary']] });
  vals.push({ range: `'🏦 Estate Distribution'!A11:E11`, values: [['Beneficiary', 'Share %', 'Est. Amount ($)', 'Paid ($)', 'Remaining ($)']] });
  const benNames = ['Robert James Williams', 'Catherine Anne Williams', 'Thomas Edward Williams', 'Margaret Grace Williams Charitable Trust'];
  const benShares = ['40%', '30%', '20%', '10%'];
  const summaryRows = benNames.map((name, i) => [
    name,
    benShares[i],
    `=IFERROR(SUMIF($C$4:$C$7,"${name}",$F$4:$F$7),0)`,
    `=IFERROR(SUMIF($C$4:$C$7,"${name}",$G$4:$G$7),0)`,
    `=C${12+i}-D${12+i}`,
  ]);
  vals.push({ range: `'🏦 Estate Distribution'!A12`, values: summaryRows });

  await valuesBatchUpdate(id, vals, 'dist-vals');
  console.log('Estate Distribution complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
