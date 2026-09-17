'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DEC = sheetMap['🌸 Decor, Food & Drinks'];

// Cols A:J (10 cols, indices 0-9)
// A=# B=Item Name C=Vendor(INDEX/MATCH from Budget) D=Category E=Qty F=Unit Cost G=Total H=Status I=Responsible J=Notes
// Budget col B = Vendor, Budget col C = Category (indices 1,2 → B,C in A1)
// Decor col C formula: INDEX(Budget!$B$4:$B$13, MATCH(Decor!D{row}, Budget!$C$4:$C$13, 0))

const ITEMS = [
  // [name, category, qty, unitCost, status, responsible, notes]
  ['Pillar candle centerpieces (tall)',   'Decor',    8,   45,  'Confirmed',  'Bloom & Blossom',  'White & cream, varied heights'],
  ['Rose gold charger plates',            'Decor',   50,    4,  'Ordered',    'Party Supplies Plus','Set of 50, 12" diameter'],
  ['Fairy light curtain backdrop',        'Decor',    2,   80,  'Confirmed',  'Grand Ballroom Events','Venue supplied, 2 panels'],
  ['Personalized place cards',            'Decor',   50,   1.5, 'Ordered',   'Paper & Craft Studio','Calligraphy style'],
  ['Floral table runners (per table)',    'Decor',    8,   35,  'Confirmed',  'Bloom & Blossom',  'Eucalyptus & blush roses'],
  ['3-course plated dinner',             'Food',    35,   55,  'Confirmed',  'Delicious Bites Co','Entrée options: chicken, fish, beef, veg'],
  ['Chicken entrée (main)',               'Food',    20,   0,   'Confirmed',  'Delicious Bites Co','Included in dinner package'],
  ['Beef entrée (main)',                  'Food',     8,   0,   'Confirmed',  'Delicious Bites Co','Included in dinner package'],
  ['Vegetarian / vegan entrée',           'Food',     4,   0,   'Confirmed',  'Delicious Bites Co','Included in dinner package'],
  ['Canapés & appetizers',               'Food',   100,   0,   'Confirmed',  'Delicious Bites Co','Passed during cocktail hour'],
  ['Champagne welcome toast',            'Drinks',  35,   12,  'Ordered',    'Delicious Bites Co','Moët & Chandon, 1 glass/person'],
  ['House wine (red & white)',           'Drinks',  10,   22,  'Ordered',    'Delicious Bites Co','2 bottles per table'],
  ['Signature mocktail',                 'Drinks',  35,    6,  'Confirmed',  'Delicious Bites Co','Rose lemonade w/ garnish'],
  ['Sparkling water & soft drinks',      'Drinks',   5,   18,  'Confirmed',  'Delicious Bites Co','5 assorted bottles'],
  ['3-tier engagement cake',             'Desserts', 1,  350,  'Ordered',    'Sweet Endings Bakery','Vanilla & raspberry, fondant floral'],
  ['Petit fours & macarons',             'Desserts',50,    2.5,'Ordered',    'Sweet Endings Bakery','Dessert table display'],
  ['Engagement favors — candle jars',    'Favors',  30,    8,  'Ordered',    'Party Supplies Plus','Personalized labels, rosé scent'],
  ['Custom linen napkins',               'Supplies',50,    3,  'Confirmed',  'Party Supplies Plus','Blush rose color, monogrammed'],
  ['Event program cards',               'Supplies', 30,    1,  'Ordered',    'Paper & Craft Studio','Double-sided, envelopes'],
  ['Photo booth props & frame',          'Decor',    1,   45,  'Ordered',    'Party Supplies Plus','Includes 20 prop items'],
];

// Budget tab category col is C (index 2), vendor name col is B (index 1)
// In A1 notation (Budget data rows 4-13): Budget!$B$4:$B$13, Budget!$C$4:$C$13
const BUD_VENDOR_RANGE  = `'💰 Budget & Vendor Tracker'!$B$4:$B$13`;
const BUD_CAT_RANGE     = `'💰 Budget & Vendor Tracker'!$C$4:$C$13`;

function fmtCell(bg, textColor, bold, size, align) {
  return {
    backgroundColor: hex(bg),
    textFormat: { foregroundColor: hex(textColor), bold: !!bold, fontSize: size || 10 },
    horizontalAlignment: align || 'LEFT',
    verticalAlignment: 'MIDDLE',
  };
}

(async () => {
  const reqs = [];

  // Row 0: full-width title
  reqs.push({ mergeCells: { range: gridRange(DEC, 0, 1, 0, 10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DEC, 0, 1, 0, 10),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 1: stats
  reqs.push({ mergeCells: { range: gridRange(DEC, 1, 2, 0, 10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DEC, 1, 2, 0, 10),
    cell: { userEnteredFormat: fmtCell(C.roseGold, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 2: headers
  reqs.push({ repeatCell: {
    range: gridRange(DEC, 2, 3, 0, 10),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows
  for (let r = 0; r < 20; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(DEC, r + 3, r + 4, 0, 10),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Center A, E, H; right-align F, G
    reqs.push({ repeatCell: {
      range: gridRange(DEC, r + 3, r + 4, 0, 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DEC, r + 3, r + 4, 4, 5),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DEC, r + 3, r + 4, 5, 7),
      cell: { userEnteredFormat: { horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  }

  // Currency format for F, G
  const currFmt = { type: 'CURRENCY', pattern: '"$"#,##0.00' };
  reqs.push({ repeatCell: {
    range: gridRange(DEC, 3, 23, 5, 7),
    cell: { userEnteredFormat: { numberFormat: currFmt } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DEC, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DEC, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Col widths: A=40 B=250 C=170 D=110 E=60 F=100 G=100 H=110 I=160 J=200
  const cw = [40, 250, 170, 110, 60, 100, 100, 110, 160, 200];
  cw.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DEC, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Borders
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(DEC, 2, 23, 0, 10),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  await batchUpdate(id, reqs, 'decor-format');

  // Values
  const data = [];
  data.push({ range: `'🌸 Decor, Food & Drinks'!A1`, values: [['🌸 DECOR, FOOD & DRINKS']] });
  data.push({
    range: `'🌸 Decor, Food & Drinks'!A2`,
    values: [[`=IFERROR("Total Items: "&COUNTA(B4:B200)&"  |  Confirmed: "&COUNTIF(H4:H200,"Confirmed")&"  |  Total Cost: $"&TEXT(SUM(G4:G200),"#,##0"),"")`]],
  });
  data.push({
    range: `'🌸 Decor, Food & Drinks'!A3:J3`,
    values: [['#','Item Name','Vendor (auto)','Category','Qty','Unit Cost','Total','Status','Responsible','Notes']],
  });

  const rows = ITEMS.map((item, i) => {
    const row = i + 4;
    const vendorFormula = `=IFERROR(INDEX(${BUD_VENDOR_RANGE}, MATCH(D${row},${BUD_CAT_RANGE},0)),"")`;
    const totalFormula  = `=IFERROR(E${row}*F${row},"")`;
    return [i + 1, item[0], vendorFormula, item[1], item[2], item[3], totalFormula, item[4], item[5], item[6]];
  });
  data.push({ range: `'🌸 Decor, Food & Drinks'!A4:J23`, values: rows });

  await valuesBatchUpdate(id, data, 'decor-values');
  console.log('Decor, Food & Drinks complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
