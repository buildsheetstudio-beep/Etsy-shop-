'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GF = sheetMap['🎁 Gift Tracker'];

// Layout:
// Row 1: Title (A1:K1)
// Row 2: Summary strip
// Row 3: Col headers
// Rows 4-23: 20 gift entries (one per guest from guest list)
// Cols: A=#(formula), B=Guest Name, C=Gift Description, D=Category, E=Est Value, F=Received?(checkbox),
//       G=Thank You Sent?(checkbox), H=Thank You Date, I=Notes,
//       (spacer J), K=Category summary label, L=Category count
// Category summary at K3:L12 (10 categories)

const NCOLS = 9; // A-I (main table)

const GIFTS = [
  // [guestName, giftDesc, category, estValue, received, tyDone, tyDate, notes]
  ['Sarah Mitchell','Personalised keepsake book','Keepsake',45,true,false,'','Guest of honour — receive at party'],
  ['Margaret Mitchell','Baby clothing bundle (0-3 months)',           'Clothing',   80, true, true, '25 Jul 2025',''],
  ['Claire Mitchell','Gift card ($100) — Target',                    'Gift Card',  100,true, true, '25 Jul 2025',''],
  ['Diana Thompson','Pram accessories set',                          'Nursery',    120,true, true, '28 Jul 2025',''],
  ['Lucy Chen','Awaiting gift',                                      'Other',      0,  false,false,'','Not yet received'],
  ['Emma Clarke','Nursing pillow + 3 covers',                        'Feeding',    95, true, true, '22 Jul 2025','Co-host gift'],
  ['Mia Patel','Awaiting',                                           'Other',      0,  false,false,'','Not yet received'],
  ['Zoe Roberts','Baby monitor — VTech',                             'Nursery',    130,true, true, '28 Jul 2025',''],
  ['Amy Johnson','Gift card ($50) — Baby Bunting',                   'Gift Card',  50, true, true, '28 Jul 2025',''],
  ['Natalie Wong','Nursery décor set — cloud mobile + prints',       'Nursery',    110,true, true, '28 Jul 2025',''],
  ['Rachel Green','Awaiting',                                        'Other',      0,  false,false,'','Not yet received'],
  ['Sophie Taylor','(Declined — not attending)',                     'Other',      0,  false,false,'','Travelling'],
  ['Olivia Brown','Baby bath set — Gaia Natural',                    'Bath & Care',75, true, false,'','Thank-you to send'],
  ['Hannah Davis','(No response — gift unknown)',                    'Other',      0,  false,false,'',''],
  ['Kate Wilson','Baby clothing bundle (3-6 months)',                'Clothing',   70, true, true, '28 Jul 2025',''],
  ['Jessica Lee','Breastfeeding bundle — nipple shields, pads, balm','Feeding',    85, true, true, '28 Jul 2025',''],
  ['Lily Thompson','Baby book collection x5',                        'Toys & Books',60,true, true, '22 Jul 2025','Co-host gift'],
  ['Anna Scott','(RSVP pending — gift unknown)',                     'Other',      0,  false,false,'',''],
  ['Beth Carter','(No response — gift unknown)',                     'Other',      0,  false,false,'',''],
  ['Chloe Evans','Keepsake blanket — personalised',                  'Keepsake',   85, true, false,'','Thank-you pending'],
];

(async () => {
  const reqs = [];

  // Column widths: A=#(50), B=Name(180), C=Gift(220), D=Category(110), E=Value(80),
  //               F=Rcvd(70), G=TY(70), H=TY Date(100), I=Notes(180),
  //               J=spacer(40), K=Cat label(150), L=Cat count(60)
  [50,180,220,110,80,70,70,100,180,40,150,60].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GF, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GF, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GF, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GF, dimension: 'ROWS', startIndex: 3, endIndex: 23 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Title row A1:I1
  reqs.push({ mergeCells: { range: gridRange(GF, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GF, 0, 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 2
  reqs.push({ repeatCell: {
    range: gridRange(GF, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  [[0,2],[2,4],[4,6],[6,9]].forEach(([c1,c2]) => {
    reqs.push({ mergeCells: { range: gridRange(GF, 1, 2, c1, c2), mergeType: 'MERGE_ALL' }});
  });

  // Col headers row 3
  reqs.push({ repeatCell: {
    range: gridRange(GF, 2, 3, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows 4-23 (indices 3-22)
  for (let r = 3; r < 23; r++) {
    const bg = (r % 2 === 1) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(GF, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // # col A — center, formula bg
  reqs.push({ repeatCell: {
    range: gridRange(GF, 3, 23, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', backgroundColor: hex(C.ivory) }},
    fields: 'userEnteredFormat(horizontalAlignment,backgroundColor)',
  }});

  // Est Value col E — currency
  reqs.push({ repeatCell: {
    range: gridRange(GF, 3, 23, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT' }},
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }});

  // Checkboxes cols F(5) and G(6) — center
  [5, 6].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(GF, 3, 23, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', backgroundColor: hex(C.ivory) }},
      fields: 'userEnteredFormat(horizontalAlignment,backgroundColor)',
    }});
    reqs.push({ setDataValidation: {
      range: gridRange(GF, 3, 23, c, c + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // TY Date col H — center, text
  reqs.push({ repeatCell: {
    range: gridRange(GF, 3, 23, 7, 8),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'TEXT' } }},
    fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
  }});

  // Borders on main table
  reqs.push({ updateBorders: {
    range: gridRange(GF, 2, 23, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Category summary block (K3:L12, indices 10-11, rows 2-12)
  reqs.push({ repeatCell: {
    range: gridRange(GF, 2, 12, 10, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Summary header
  reqs.push({ repeatCell: {
    range: gridRange(GF, 2, 3, 10, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateBorders: {
    range: gridRange(GF, 2, 12, 10, 12),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'gifts-format');

  // Values
  const vals = [];
  vals.push({ range: "'🎁 Gift Tracker'!A1", values: [['  🎁  Gift Tracker']] });

  // Summary row 2
  vals.push({ range: "'🎁 Gift Tracker'!A2", values: [['=COUNTA(B4:B200)&" gifts logged"']] });
  vals.push({ range: "'🎁 Gift Tracker'!C2", values: [['=COUNTIF(F4:F200,TRUE)&" received"']] });
  vals.push({ range: "'🎁 Gift Tracker'!E2", values: [['=COUNTIF(G4:G200,TRUE)&" thank-yous sent"']] });
  vals.push({ range: "'🎁 Gift Tracker'!G2", values: [['=IFERROR("Est. total: $"&TEXT(SUMIF(F4:F200,TRUE,E4:E200),"#,##0"),"—")']] });

  // Col headers
  vals.push({ range: "'🎁 Gift Tracker'!A3", values: [['#','GUEST NAME','GIFT DESCRIPTION','CATEGORY','EST. VALUE','RCVD?','TY SENT?','TY DATE','NOTES']] });

  // Gift rows with # formula
  const giftRows = GIFTS.map(([name, desc, cat, val, rcvd, ty, tyDate, notes], i) => {
    const row = 4 + i;
    return [`=IF(B${row}="","",ROW()-3)`, name, desc, cat, val, rcvd ? 'TRUE' : 'FALSE', ty ? 'TRUE' : 'FALSE', tyDate, notes];
  });
  vals.push({ range: "'🎁 Gift Tracker'!A4", values: giftRows });

  // Category summary K3:L12
  vals.push({ range: "'🎁 Gift Tracker'!K3", values: [['GIFT CATEGORY','COUNT']] });
  const CATS = ['Clothing','Nursery','Feeding','Toys & Books','Gift Card','Bath & Care','Keepsake','Handmade','Other'];
  CATS.forEach((cat, i) => {
    vals.push({ range: `'🎁 Gift Tracker'!K${4+i}`, values: [[cat, `=COUNTIF($D$4:$D$200,"${cat}")`]] });
  });

  await valuesBatchUpdate(id, vals, 'gifts-values');
  console.log('Gift Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
