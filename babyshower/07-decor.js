'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DC = sheetMap['🎀 Décor, Food & Drinks'];

// Layout:
// Row 1: Title (A1:J1)
// Row 2: Décor section col headers
// Rows 3-17: 15 décor items
// Row 18: spacer (section divider)
// Row 19: Food & Drinks section header
// Row 20: Food col headers
// Rows 21-35: 15 food items
//
// Décor cols (A-J): A=Item, B=Category, C=Colour/Style, D=Qty, E=Supplier, F=Cost, G=Status, H=Notes, (9 total = A-I)
// Food cols (A-J): A=Item, B=Food Type, C=Dietary Tag, D=Qty/Portions, E=Supplier, F=Cost, G=Status, H=Serves, I=Notes

const NCOLS = 9; // A-I

const DECOR = [
  // [item, category, colour/style, qty, supplier, cost, status, notes]
  ['Balloon arch — lilac & blush','Balloons','Lilac, blush, white',1,'Bloom & Co',120,'Ordered','To be installed morning of party'],
  ['Fresh floral centrepieces x4','Florals','Garden roses, baby\'s breath, lilac',4,'Bloom & Co',180,'Ordered','One per table'],
  ['Welcome sign — acrylic','Signage','Gold lettering on acrylic',1,'Etsy — LoveLaserCo',65,'Arrived',''],
  ['Baby shower banner','Signage','Lilac & gold',1,'Kmart',12,'Arrived',''],
  ['Table runner — blush linen','Table Setting','Blush linen',4,'Target',48,'Arrived',''],
  ['Linen napkins x20','Table Setting','White & blush',20,'IKEA',35,'Arrived',''],
  ['Candle votives x10','Lighting','Clear glass, cream candle',10,'Kmart',25,'Arrived',''],
  ['Nappy cake centrepiece','Backdrop','Lilac ribbon',1,'DIY — Claire Mitchell',0,'Done','Homemade by Claire'],
  ['Backdrop arch — floral & tulle','Backdrop','White, lilac, blush',1,'Bloom & Co',220,'Ordered','For photo area'],
  ['Chair sashes x20','Table Setting','Lilac organza',20,'Kmart',40,'Arrived',''],
  ['Gift table display items','Table Setting','Mixed florals & ribbons',1,'DIY',15,'Done',''],
  ['Party favour boxes x20','Favours','White kraft, lilac ribbon',20,'Etsy — BoxedWithLove',90,'Ordered',''],
  ['Balloon weights x5','Balloons','Gold',5,'Party City',8,'Arrived',''],
  ['Photo props basket','Photo','"Oh Baby" props set',1,'Kmart',18,'Arrived',''],
  ['Miscellaneous décor (ribbon, tape, pins)','Misc','Various',1,'Various',22,'Done',''],
];

const FOOD = [
  // [item, type, dietary, qty, supplier, cost, status, serves, notes]
  ['Finger sandwiches assorted','Appetiser','Standard',60,'The Pantry Caterers',0,'Confirmed',20,'Included in catering package'],
  ['Scones with jam & cream','Dessert','Standard',40,'The Pantry Caterers',0,'Confirmed',20,'Included in catering package'],
  ['Macarons — lilac & blush','Dessert','Vegetarian',40,'Sweet Creations',0,'Confirmed',20,'Part of cake order'],
  ['Fresh fruit platter','Appetiser','Vegan, Gluten Free',2,'The Pantry Caterers',0,'Confirmed',20,'Included in catering package'],
  ['Gluten-free platter','Appetiser','Gluten Free',1,'The Pantry Caterers',35,'Confirmed',5,'For Lucy, Kate & Chloe'],
  ['Vegan bites platter','Appetiser','Vegan',1,'The Pantry Caterers',0,'Confirmed',3,'For Mia Patel'],
  ['Custom nappy cake (display)','Cake','Vegetarian',1,'Sweet Creations',280,'Confirmed',0,'Display only — not eaten'],
  ['Mini cupcakes x24','Dessert','Vegetarian',24,'Sweet Creations',0,'Confirmed',20,'Included in cake order'],
  ['Sparkling wine — Chandon','Drinks','Standard',3,'Dan Murphy\'s',95,'Confirmed',20,'For toast & drinks'],
  ['Sparkling apple juice (non-alcoholic)','Non-Alcoholic','Standard',4,'Woolworths',28,'Confirmed',20,'For mocktails & non-drinkers'],
  ['Elderflower lemonade','Non-Alcoholic','Vegan, Gluten Free',4,'Woolworths',32,'Confirmed',20,''],
  ['Iced tea (herbal blend)','Non-Alcoholic','Vegan, Gluten Free',2,'Woolworths',16,'Confirmed',20,''],
  ['Water — still & sparkling','Non-Alcoholic','Standard',4,'Woolworths',24,'Confirmed',20,''],
  ['Cheese & antipasto board','Appetiser','Standard',2,'The Pantry Caterers',0,'Confirmed',20,'Included in catering package'],
  ['Chocolate dipped strawberries','Dessert','Vegetarian',30,'Homemade — Natalie Wong','Guest Contribution',0,'Confirmed',20,'Natalie is bringing these'],
];

(async () => {
  const reqs = [];

  // Column widths
  [200, 130, 140, 70, 160, 80, 90, 80, 200].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DC, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DC, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DC, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DC, dimension: 'ROWS', startIndex: 2, endIndex: 17 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DC, dimension: 'ROWS', startIndex: 17, endIndex: 20 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DC, dimension: 'ROWS', startIndex: 20, endIndex: 35 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Title row
  reqs.push({ mergeCells: { range: gridRange(DC, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DC, 0, 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Décor section header row 2 (index 1)
  reqs.push({ repeatCell: {
    range: gridRange(DC, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Décor data rows 3-17 (indices 2-16)
  for (let r = 2; r < 17; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(DC, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Food section header row 19 (index 18)
  reqs.push({ repeatCell: {
    range: gridRange(DC, 18, 19, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Spacer row 18 (index 17) — champagne
  reqs.push({ repeatCell: {
    range: gridRange(DC, 17, 18, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.champagne) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Food col headers row 20 (index 19)
  reqs.push({ repeatCell: {
    range: gridRange(DC, 19, 20, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Food data rows 21-35 (indices 20-34)
  for (let r = 20; r < 35; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(DC, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Cost col F (index 5) currency format
  [gridRange(DC, 2, 17, 5, 6), gridRange(DC, 20, 35, 5, 6)].forEach(range => {
    reqs.push({ repeatCell: {
      range,
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT' }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    }});
  });

  // Qty/Serves center
  [3, 7].forEach(c => {
    [gridRange(DC, 2, 17, c, c + 1), gridRange(DC, 20, 35, c, c + 1)].forEach(range => {
      reqs.push({ repeatCell: {
        range,
        cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
        fields: 'userEnteredFormat.horizontalAlignment',
      }});
    });
  });

  // Wrap for item col A, notes col I
  [0, 8].forEach(c => {
    [gridRange(DC, 2, 17, c, c + 1), gridRange(DC, 20, 35, c, c + 1)].forEach(range => {
      reqs.push({ repeatCell: {
        range,
        cell: { userEnteredFormat: { wrapStrategy: 'CLIP' }},
        fields: 'userEnteredFormat.wrapStrategy',
      }});
    });
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(DC, 1, 17, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});
  reqs.push({ updateBorders: {
    range: gridRange(DC, 18, 35, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'decor-format');

  // Values
  const vals = [];
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A1", values: [['  🎀  Décor, Food & Drinks']] });

  // Décor section header
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A2", values: [['🌸  DÉCOR & DECORATIONS  —  15 items']] });

  // Décor col headers
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A3", values: [['ITEM','CATEGORY','COLOUR / STYLE','QTY','SUPPLIER','COST','STATUS','HIRED?','NOTES']] });

  // Décor rows
  const decorRows = DECOR.map(([item, cat, style, qty, supplier, cost, status, notes]) =>
    [item, cat, style, qty, supplier, cost, status, '', notes]
  );
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A4", values: decorRows });

  // Spacer row total
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A18", values: [['DÉCOR TOTAL', '', '', '', '', `=SUM(F4:F17)`, '', '', '']] });

  // Food section header
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A19", values: [['🍽️  FOOD & DRINKS  —  15 items']] });

  // Food col headers
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A20", values: [['ITEM','FOOD TYPE','DIETARY','QTY / PORTIONS','SUPPLIER','COST','STATUS','SERVES','NOTES']] });

  // Food rows
  const foodRows = FOOD.map(([item, type, dietary, qty, supplier, cost, status, serves, notes]) =>
    [item, type, dietary, qty, supplier, cost, status, serves, notes]
  );
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A21", values: foodRows });

  // Food total row
  vals.push({ range: "'🎀 Décor, Food & Drinks'!A36", values: [['FOOD & DRINKS TOTAL', '', '', '', '', `=SUMIF(F21:F35,"<>Guest Contribution")`, '', '', '']] });

  await valuesBatchUpdate(id, vals, 'decor-values');
  console.log('Décor, Food & Drinks complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
