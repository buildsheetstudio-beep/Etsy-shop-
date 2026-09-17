'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GFT = sheetMap['🎁 Gift Tracker & Wishlist'];

const WISHLIST = [
  ['KitchenAid Stand Mixer — Pistachio','Homewares','Must Have',349,'https://kitchenaid.com','Would love the 5qt Artisan in Pistachio'],
  ['Le Creuset Cast Iron Casserole — Sage','Homewares','Must Have',270,'','28cm round, sage green'],
  ['Linen Duvet Set — Blush King','Homewares','Would Love',185,'','100% French linen'],
  ['Personalised Recipe Book','Books & Stationery','Would Love',45,'','Hand-bound with name on cover'],
  ['Jo Malone Candle Set','Homewares','Would Love',120,'','Peony & Blush Suede collection'],
  ['Wine Tasting Experience (2 people)','Experience','Nice to Have',95,'','Local vineyard tour'],
  ['Monogrammed Dressing Gowns (x2)','Clothing & Accessories','Would Love',130,'','Matching his & hers'],
  ['Artwork — Abstract Floral Print','Homewares','Nice to Have',80,'','Framed A2, neutral tones'],
  ['Polaroid Camera + Film Bundle','Keepsake','Nice to Have',75,'','Fujifilm Instax Wide'],
  ['Gourmet Hamper','Food & Drink','Nice to Have',90,'','Artisan cheeses, crackers, wine'],
  ['Silk Pillowcases (2pk)','Homewares','Would Love',65,'','Ivory, queen size'],
  ['Cookbook — Ottolenghi Simple','Books & Stationery','Nice to Have',28,'','Paperback'],
  ['Spa Day Voucher','Experience','Must Have',150,'','Local luxury spa'],
  ['Personalised Champagne Flute Set','Keepsake','Would Love',55,'','Set of 2, engraved'],
  ['Custom Illustration — Portrait','Keepsake','Nice to Have',120,'','Watercolour couple portrait'],
];

const GIFTS = [
  ['Charlotte Harrison','2026-10-04','KitchenAid Stand Mixer','Homewares',349,'Credit Card','Yes','Thank You Sent','Sophie helped carry it'],
  ['Sophie Clarke','2026-10-04','Le Creuset Casserole Dish','Homewares',270,'Bank Transfer','Yes','Thank You Sent',''],
  ['Lily Thompson','2026-10-04','Personalised Champagne Flutes','Keepsake',55,'Cash','Yes','Thank You Sent','Beautifully wrapped'],
  ['Emma Grant','2026-10-04','Jo Malone Candle Set','Homewares',120,'Credit Card','Yes','Thank You Sent',''],
  ['Ava Fitzgerald','2026-10-04','Linen Duvet Set','Homewares',185,'Gift','No','Pending',''],
  ['Isabella Moore','2026-10-04','Polaroid Camera + Film','Keepsake',75,'Cash','No','Pending',''],
  ['Grace Sullivan','2026-10-04','Spa Day Voucher','Experience',150,'Bank Transfer','Yes','Thank You Sent',''],
  ['Olivia Bennett','2026-10-04','Silk Pillowcases','Homewares',65,'Credit Card','No','Pending',''],
  ['Amelia Carter','2026-10-04','Cookbook — Ottolenghi','Books & Stationery',28,'Cash','No','Pending',''],
  ['Phoebe Walsh','2026-10-04','Monogrammed Dressing Gowns','Clothing & Accessories',130,'Bank Transfer','Yes','Thank You Sent',''],
  ['Ruby Chambers','2026-10-04','Wine Tasting Experience','Experience',95,'Gift','No','Pending',''],
  ['Isla Reid','2026-10-04','Personalised Recipe Book','Books & Stationery',45,'Cash','No','Not Started',''],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🎁 Gift Tracker & Wishlist'";

  // Column widths
  const colW = [200,130,130,120,90,120,70,130,170];
  colW.forEach((w,i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: GFT, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Row 0: title
  reqs.push({ mergeCells: { range: gridRange(GFT,0,1,1,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GFT,0,1,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GFT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['🎁 Gift Tracker & Wishlist']] });

  // Row 1: Wishlist section header (0-indexed 1)
  reqs.push({ mergeCells: { range: gridRange(GFT,1,2,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GFT,1,2,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['BRIDE\'S WISHLIST']] });

  // Row 2: wishlist col headers (0-indexed 2)
  const wlHeaders = ['ITEM','CATEGORY','PRIORITY','APPROX PRICE','LINK / STORE','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(GFT,2,3,0,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A3`, values: [wlHeaders] });

  // Wishlist rows 4-18 (0-indexed 3-17)
  for (let i = 0; i < WISHLIST.length; i++) {
    const ri  = 3 + i;
    const row = 4 + i;
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(GFT, ri, ri+1, 0, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(GFT, ri, ri+1, 3, 4),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [WISHLIST[i]] });
  }

  // Divider (0-indexed 18)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GFT, dimension: 'ROWS', startIndex: 18, endIndex: 19 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(GFT,18,19,0,9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Row 20: Gift Log section header (0-indexed 19)
  reqs.push({ mergeCells: { range: gridRange(GFT,19,20,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GFT,19,20,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A20`, values: [['GIFT LOG']] });

  // Row 21: summary stats (0-indexed 20)
  reqs.push({ repeatCell: {
    range: gridRange(GFT,20,21,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.parchment),
      textFormat: { foregroundColor: hex(C.muted), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A21`, values: [[
    'Total Gifts:',
    `=COUNTA(A23:A200)`,
    'Total Value:',
    `=IFERROR(SUM(E23:E200),"")`,
    'Thank Yous Sent:',
    `=COUNTIF(H23:H200,"Thank You Sent")`,
    'Pending:',
    `=COUNTIF(H23:H200,"Pending")`,
    '',
  ]]});
  reqs.push({ repeatCell: {
    range: gridRange(GFT,20,21,3,4),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }}},
    fields: 'userEnteredFormat(numberFormat)',
  }});

  // Row 22: gift log col headers (0-indexed 21)
  const giftHeaders = ['GUEST NAME','DATE RECEIVED','GIFT DESCRIPTION','CATEGORY','VALUE','PAYMENT METHOD','RECEIVED?','THANK YOU?','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(GFT,21,22,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A22`, values: [giftHeaders] });

  // Gift rows 23-34 (0-indexed 22-33)
  for (let i = 0; i < GIFTS.length; i++) {
    const ri  = 22 + i;
    const row = 23 + i;
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(GFT, ri, ri+1, 0, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(GFT, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [GIFTS[i]] });
  }

  // Gift category chart data: cols K-L (indices 10-11), rows 0-9 (0-indexed)
  // Categories and COUNTIF
  const giftCats = ['Homewares','Clothing & Accessories','Gift Card','Experience','Food & Drink','Books & Stationery','Keepsake','Handmade','Other'];
  vals.push({ range: `${TAB}!K1`, values: [['Category']] });
  vals.push({ range: `${TAB}!L1`, values: [['Count']] });
  for (let i = 0; i < giftCats.length; i++) {
    vals.push({ range: `${TAB}!K${2+i}`, values: [[giftCats[i]]] });
    vals.push({ range: `${TAB}!L${2+i}`, values: [[`=COUNTIF($D$23:$D$200,"${giftCats[i]}")`]] });
  }

  await batchUpdate(id, reqs, 'gifts-format');
  await valuesBatchUpdate(id, vals, 'gifts-values');
  console.log('Gift Tracker & Wishlist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
