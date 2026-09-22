'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const VT = sheetMap['🏪 Vendor & Supplier Tracker'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: VT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: VT, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: VT, dimension: 'ROWS', startIndex: 3, endIndex: 45 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths: A(30) B(Vendor 180) C(Category 120) D(Event 180) E(Contact 140) F(Phone 110) G(Email 200) H(Contract $ 110) I(Deposit 110) J(Deposit Paid 90) K(Balance Due 110) L(Pay Status 120) M(Contract Signed 90) N(Review Date 100) O(Rating 60) P(Notes 200)
  const colW = [30, 180, 120, 180, 140, 110, 200, 110, 110, 90, 110, 120, 90, 100, 60, 200];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: VT, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row — merge A:O only (leave P-R for summary)
  reqs.push({ mergeCells: { range: gridRange(VT, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(VT, 0, 1, 0, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmBrown),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary block rows 2-3 header label
  reqs.push({ repeatCell: {
    range: gridRange(VT, 1, 2, 0, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(VT, 2, 3, 0, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleRose),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(VT, 3, 4, 0, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmBrown),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows alternating
  for (let r = 4; r < 45; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(VT, r, r+1, 0, 16),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Currency format H, I, K (contract, deposit, balance)
  [7, 8, 10].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(VT, 4, 45, c, c+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  });

  // Formula col K (balance due) — paleRose
  reqs.push({ repeatCell: { range: gridRange(VT, 4, 45, 10, 11), cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose) } }, fields: 'userEnteredFormat.backgroundColor' }});

  // Date format N (index 13)
  reqs.push({ repeatCell: {
    range: gridRange(VT, 4, 45, 13, 14),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Checkbox cols J, M (indices 9, 12)
  [9, 12].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(VT, 4, 45, c, c+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose), horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    }});
  });

  // Center cols A, B, C, L, O
  [0, 2, 11, 14].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(VT, 3, 45, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(VT, 3, 45, 0, 16),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'vendor-format');

  const data = [];
  data.push({ range: `'🏪 Vendor & Supplier Tracker'!A1`, values: [['🏪 Vendor & Supplier Tracker — All Events']] });

  // Summary labels row 2
  data.push({ range: `'🏪 Vendor & Supplier Tracker'!A2:F2`, values: [['Total Vendors','Paid in Full','Outstanding','Overdue','Total Contracted','Total Paid']] });
  // Summary values row 3
  data.push({ range: `'🏪 Vendor & Supplier Tracker'!A3:F3`, values: [[
    '=COUNTA(B5:B45)',
    '=COUNTIF(L5:L45,"Paid in Full")',
    '=COUNTIF(L5:L45,"Not Paid")',
    '=COUNTIF(L5:L45,"Overdue")',
    '=SUM(H5:H45)',
    '=SUM(I5:I45)',
  ]] });

  data.push({ range: `'🏪 Vendor & Supplier Tracker'!A4:P4`, values: [['#','Vendor Name','Category','Event','Contact Person','Phone','Email','Contract ($)','Deposit Paid','Deposit?','Balance Due','Payment Status','Contract Signed','Review Date','Rating','Notes']] });

  const vendors = [
    [1,'Grand Convention Centre','Venue','Tech Summit 2025','David Wong','02 9000 1111','events@gcc.com.au',8500,8500,true,'=H5-I5','Paid in Full',true,'2025-09-15',5,'Excellent venue, great AV'],
    [2,'Fresh Table Catering Co.','Catering','Tech Summit 2025','Maria Santos','0411 200 001','maria@freshtable.com',4200,4200,true,'=H6-I6','Paid in Full',true,'2025-09-15',4,''],
    [3,'Print & Digital Agency','Marketing','Tech Summit 2025','Jake Liu','0411 200 002','jake@printdigital.com',2000,2000,true,'=H7-I7','Paid in Full',true,'',4,'On time delivery'],
    [4,'ProAV Solutions','Audio/Visual','Tech Summit 2025','Sam Briggs','0411 200 003','sam@proav.com.au',3500,0,false,'=H8-I8','Not Paid',true,'2025-09-14',0,'Payment due day before'],
    [5,'Riverside Ballroom','Venue','Charity Gala 2025','Anne-Marie Cole','02 9000 2222','events@riverside.com',12000,6000,true,'=H9-I9','Deposit Paid',true,'2025-10-22',0,'Balance due 30 Sep'],
    [6,'Prestige Catering','Catering','Charity Gala 2025','Raj Kapoor','0411 200 004','raj@prestigecatering.com',9000,0,false,'=H10-I10','Not Paid',false,'2025-10-15',0,'Contract pending'],
    [7,'The Jazz Collective','Entertainment','Charity Gala 2025','Nico Ferreira','0411 200 005','nico@thejazzcollective.com',4500,0,false,'=H11-I11','Not Paid',false,'',0,''],
    [8,'Bloom & Drape Florist','Decor','Charity Gala 2025','Sophie Green','0411 200 006','sophie@bloomdrape.com',3500,1750,true,'=H12-I12','Deposit Paid',true,'2025-10-20',0,''],
    [9,'The Grand Hotel','Venue','Annual Awards Night','Michael Osei','02 9000 3333','events@grandhotel.com',9500,4750,true,'=H13-I13','Deposit Paid',true,'2025-11-28',0,''],
    [10,'Harbour View Restaurant','Venue','5-Year Anniversary Dinner','Clara Perez','02 9000 4444','events@harbourview.com',4800,2400,true,'=H14-I14','Deposit Paid',true,'2025-11-07',0,''],
  ];

  data.push({ range: `'🏪 Vendor & Supplier Tracker'!A5:P14`, values: vendors });

  await valuesBatchUpdate(id, data, 'vendor-values');
  console.log('Vendor Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
