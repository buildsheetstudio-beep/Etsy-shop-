'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const BEN = sheetMap['👨‍👩‍👧 Beneficiaries'];
const DIST = sheetMap['🏦 Estate Distribution'];

// 4 beneficiaries
// Cols: A=# B=Full Name C=Relationship D=Type E=DoB F=Contact Phone G=Contact Email H=Address I=TFN/ABN (last 3) J=Bank/BSB K=Share % L=Est Share $ M=Total Received $ N=Notes
const BENS = [
  [1, 'Robert James Williams', 'Son (Executor)', 'Individual', '1975-03-15', '+61 2 9345 6789', 'robert.williams@email.com.au', '12 Elm Street, Mosman NSW 2088', 'XXX-XXX-789', 'NAB BSB: 082-001 / 234567890', '40%', '', '', 'Also serving as executor'],
  [2, 'Catherine Anne Williams', 'Daughter', 'Individual', '1978-09-22', '+61 2 9876 5432', 'catherine.williams@email.com.au', '7 Rose Garden, Killara NSW 2071', 'XXX-XXX-456', 'WBC BSB: 032-100 / 109876543', '30%', '', '', 'Lives interstate; communication by phone/email'],
  [3, 'Thomas Edward Williams', 'Son', 'Individual', '1982-07-08', '+61 417 234 567', 'thomas.williams@gmail.com', '22 Banksia Avenue, Dee Why NSW 2099', 'XXX-XXX-123', 'CBA BSB: 062-800 / 44556677', '20%', '', '', ''],
  [4, 'Margaret Grace Williams Charitable Trust', 'Charitable Bequest', 'Trust', '', '+61 2 9000 1234', 'grants@mwcharitabletrust.org', 'Suite 4, 88 Collins Street Melbourne VIC 3000', 'ABN 12 345 678 901', 'Trust Account via Trustee', '10%', '', '', 'Per clause 8 of the Will — arts & community grants'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BEN, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: BEN, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BEN, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BEN, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BEN, dimension: 'ROWS', startIndex: 3, endIndex: 30 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Col widths: A=40 B=220 C=150 D=130 E=100 F=140 G=220 H=250 I=110 J=200 K=80 L=120 M=120 N=220
  const colW = [40, 220, 150, 130, 100, 140, 220, 250, 110, 200, 80, 120, 120, 220];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: BEN, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title row
  reqs.push({ mergeCells: { range: gridRange(BEN, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BEN, 0, 1, 0, 14), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Subtitle — estate summary bar
  reqs.push({ mergeCells: { range: gridRange(BEN, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BEN, 1, 2, 0, 14), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Column headers
  reqs.push({ repeatCell: { range: gridRange(BEN, 2, 3, 0, 14), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Beneficiary rows
  for (let r = 0; r < 4; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(BEN, r+3, r+4, 0, 14), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
        wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Totals row
  reqs.push({ repeatCell: { range: gridRange(BEN, 7, 8, 0, 14), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.paleSage),
      textFormat: { foregroundColor: hex(C.mutedSage), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: {
        top: { style: 'SOLID', color: hex(C.mutedSage), width: 2 },
        bottom: { style: 'SOLID', color: hex(C.mutedSage), width: 2 },
      },
    },
  }, fields: 'userEnteredFormat' }});

  // Center cols A, D, E, K
  [0, 3, 4, 10].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(BEN, 3, 8, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Currency format for L, M cols
  reqs.push({ repeatCell: { range: gridRange(BEN, 3, 8, 11, 13), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  // Borders
  reqs.push({ updateBorders: { range: gridRange(BEN, 2, 8, 0, 14),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'ben-fmt');

  const vals = [];
  vals.push({ range: `'👨‍👩‍👧 Beneficiaries'!A1`, values: [['👨‍👩‍👧 BENEFICIARIES — Estate of Margaret Eleanor Williams (Ref: PROB-2025-0442)']] });
  vals.push({ range: `'👨‍👩‍👧 Beneficiaries'!A2`, values: [['Estate Value (Current): =\'💼 Assets & Debts\'!F17&"  |  Total Debts: "&\'💼 Assets & Debts\'!E26&"  |  Net Estate: "&\'💼 Assets & Debts\'!F28']] });
  // Simpler subtitle
  vals.push({ range: `'👨‍👩‍👧 Beneficiaries'!A2`, values: [['4 Beneficiaries Named | Net Estate Approx. $2,045,540 | Executor: Robert James Williams | Ref: PROB-2025-0442']] });
  vals.push({
    range: `'👨‍👩‍👧 Beneficiaries'!A3:N3`,
    values: [['#', 'Full Name', 'Relationship', 'Type', 'Date of Birth', 'Contact Phone', 'Contact Email', 'Address', 'TFN / ABN (masked)', 'Bank / BSB', 'Will Share %', 'Est. Share Amount ($)', 'Total Received ($)', 'Notes']],
  });

  const benRows = BENS.map(b => b);
  // Add SUMIF for M column (Total Received from Distribution tab)
  benRows[0][12] = `=IFERROR(SUMIF('🏦 Estate Distribution'!$C$4:$C$30,$B4,'🏦 Estate Distribution'!$G$4:$G$30),0)`;
  benRows[1][12] = `=IFERROR(SUMIF('🏦 Estate Distribution'!$C$4:$C$30,$B5,'🏦 Estate Distribution'!$G$4:$G$30),0)`;
  benRows[2][12] = `=IFERROR(SUMIF('🏦 Estate Distribution'!$C$4:$C$30,$B6,'🏦 Estate Distribution'!$G$4:$G$30),0)`;
  benRows[3][12] = `=IFERROR(SUMIF('🏦 Estate Distribution'!$C$4:$C$30,$B7,'🏦 Estate Distribution'!$G$4:$G$30),0)`;
  // Est Share = Net estate * share %
  const netEstate = 2045540;
  benRows[0][11] = `=IFERROR(TEXT(${netEstate}*0.40,"$#,##0.00"),"—")`;
  benRows[1][11] = `=IFERROR(TEXT(${netEstate}*0.30,"$#,##0.00"),"—")`;
  benRows[2][11] = `=IFERROR(TEXT(${netEstate}*0.20,"$#,##0.00"),"—")`;
  benRows[3][11] = `=IFERROR(TEXT(${netEstate}*0.10,"$#,##0.00"),"—")`;

  vals.push({ range: `'👨‍👩‍👧 Beneficiaries'!A4`, values: benRows });

  // Totals
  vals.push({
    range: `'👨‍👩‍👧 Beneficiaries'!A8:N8`,
    values: [['TOTALS', '', '', '', '', '', '', '', '', '', '=SUMIF(K4:K7,"*%",K4:K7)', '—', '=SUM(M4:M7)', '']],
  });

  await valuesBatchUpdate(id, vals, 'ben-vals');
  console.log('Beneficiaries complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
