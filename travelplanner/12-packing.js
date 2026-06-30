const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 1128916394; // Packing Checklist

// ~40 packing items organised by category
const items = [
  // Category, Item, Quantity, Packed?, Notes
  // Clothing
  ['Clothing','T-shirts / tops',7,false,'Lightweight, mix of smart-casual'],
  ['Clothing','Trousers / shorts',3,false,'1 smart pair for restaurants'],
  ['Clothing','Summer dress / smart outfit',2,false,'For evenings out'],
  ['Clothing','Underwear',8,false,'Pack +1 extra'],
  ['Clothing','Socks',6,false,''],
  ['Clothing','Light cardigan / jacket',1,false,'European evenings can be cool'],
  ['Clothing','Swimwear',2,false,'Hotel pool or beach days'],
  ['Clothing','Pyjamas',1,false,''],
  // Footwear
  ['Footwear','Comfortable walking shoes',1,false,'Break in before trip!'],
  ['Footwear','Sandals / flip flops',1,false,''],
  ['Footwear','Smart shoes / heels',1,false,'For dinner reservations'],
  // Toiletries & Health
  ['Toiletries & Health','Toothbrush & toothpaste',1,false,''],
  ['Toiletries & Health','Shampoo & conditioner (travel)',1,false,'Check airline liquids rules'],
  ['Toiletries & Health','Deodorant',1,false,''],
  ['Toiletries & Health','Sunscreen (SPF 50)',1,false,'Hot European summer!'],
  ['Toiletries & Health','After-sun lotion',1,false,''],
  ['Toiletries & Health','Insect repellent',1,false,''],
  ['Toiletries & Health','Prescription medications',1,false,'Pack extra + keep in hand luggage'],
  ['Toiletries & Health','Paracetamol / ibuprofen',1,false,''],
  ['Toiletries & Health','Plasters / first aid kit',1,false,''],
  ['Toiletries & Health','Hand sanitiser',1,false,''],
  // Electronics & Tech
  ['Electronics & Tech','Phone & charger',1,false,''],
  ['Electronics & Tech','Universal travel adaptor',1,false,'EU plug type C'],
  ['Electronics & Tech','Portable power bank',1,false,'Min 10,000mAh'],
  ['Electronics & Tech','Earphones / headphones',1,false,'For flights/trains'],
  ['Electronics & Tech','Camera',1,false,'Or use phone camera'],
  ['Electronics & Tech','Kindle / tablet',1,false,'Optional'],
  // Documents & Finance
  ['Documents & Finance','Passport',1,false,'Check expiry date!'],
  ['Documents & Finance','Travel insurance documents',1,false,'Print + digital copy'],
  ['Documents & Finance','Hotel confirmations',1,false,'Printed copies'],
  ['Documents & Finance','Flight tickets',1,false,''],
  ['Documents & Finance','Travel credit card (no fees)',1,false,'e.g. Chase, Starling, Revolut'],
  ['Documents & Finance','Cash (GBP + some EUR)','',false,'Get EUR before departure'],
  ['Documents & Finance','EHIC / GHIC card',1,false,'EU healthcare card'],
  // Entertainment
  ['Entertainment','Guidebooks / maps',1,false,'Or download offline'],
  ['Entertainment','Journal / travel diary',1,false,''],
  // Miscellaneous
  ['Miscellaneous','Reusable water bottle',1,false,'Save money, eco-friendly'],
  ['Miscellaneous','Tote bag / day bag',1,false,'For sightseeing days'],
  ['Miscellaneous','Umbrella / packable rain mac',1,false,'Just in case!'],
  ['Miscellaneous','Luggage locks',2,false,'TSA approved'],
  ['Miscellaneous','Ziplock bags (various sizes)',1,false,'For liquids + organisation'],
];

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '🧳 PACKING CHECKLIST' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 2: Progress bar ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 4),
      cell: {
        userEnteredValue: { formulaValue: '=IFERROR(TEXT(COUNTIF(D3:D200,TRUE),"0")&" of "&TEXT(COUNTA(B3:B200),"0")&" items packed","0 of 0 items packed")' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.champagne),
          textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 4, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 4, 8),
      cell: {
        userEnteredValue: { formulaValue: '=IFERROR(REPT("▰",ROUND(COUNTIF(D3:D200,TRUE)/COUNTA(B3:B200)*10,0))&REPT("▱",10-ROUND(COUNTIF(D3:D200,TRUE)/COUNTA(B3:B200)*10,0))&"  "&TEXT(COUNTIF(D3:D200,TRUE)/COUNTA(B3:B200)*100,"0.0")&"% packed","▱▱▱▱▱▱▱▱▱▱  0.0% packed")' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleNavy),
          textFormat: { foregroundColor: hex(COLOR.midnightBlue), bold: true, fontSize: 12, fontFamily: 'Courier New' },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 3: Column headers ──
  const headers = ['#','Category','Item','✓ Packed','Qty','Notes'];
  [0,1,2,3,4,5].forEach((_, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 2, 3, c, c+1),
        cell: {
          userEnteredValue: { stringValue: headers[c] },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });
  // Merge cols 5-8 for Notes
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 5, 8), mergeType: 'MERGE_ALL' } });

  // Category colors
  const catColors = {
    'Clothing': COLOR.lightBlue,
    'Footwear': COLOR.lightGold,
    'Toiletries & Health': COLOR.lightGreen,
    'Electronics & Tech': COLOR.lightAmber,
    'Documents & Finance': COLOR.lightRed,
    'Entertainment': COLOR.champagne,
    'Miscellaneous': COLOR.softGrey,
  };

  // ── Data rows ──
  let prevCat = null;
  items.forEach((item, i) => {
    const row = 3 + i;
    const cat = item[0];
    const catBg = catColors[cat] || COLOR.white;

    if (cat !== prevCat) {
      // Category header row — insert before data if category changes
      // We'll handle this by alternating color, not inserting extra rows
      prevCat = cat;
    }

    // # auto
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(C${row+1}="","",ROW()-2)` },
          userEnteredFormat: {
            backgroundColor: hex(catBg),
            textFormat: { bold: true, foregroundColor: hex(COLOR.deepNavy) },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Category (col 1)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 1, 2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(catBg),
            textFormat: { foregroundColor: hex(COLOR.midnightBlue), bold: true, fontSize: 9 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });

    // Item name (col 2)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 2, 3),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(catBg),
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });

    // Packed checkbox (col 3) — will add BOOLEAN validation later
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 3, 4),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(catBg),
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });

    // Qty (col 4)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 4, 5),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(catBg),
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });

    // Notes (cols 5-7 merged)
    reqs.push({ mergeCells: { range: gridRange(SID, row, row+1, 5, 8), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 5, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(catBg),
            textFormat: { foregroundColor: hex(COLOR.midnightBlue), italic: true, fontSize: 9 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  });

  // Column widths
  const widths = [35,140,220,70,60,380];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 + items.length + 10 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'packing-structure');

  // ── Values ──
  const data = [];
  data.push(vr(`'Packing Checklist'!B4`, items.map(i => [i[0], i[1], i[2], i[4]])));
  // Packed column (col D = 4, so D4 = 0-indexed col 3)
  data.push(vr(`'Packing Checklist'!D4`, items.map(i => [i[3]])));

  await valuesBatchUpdate(id, data, 'packing-values');

  // BOOLEAN checkbox validation for packed column
  const valReqs = [{
    setDataValidation: {
      range: gridRange(SID, 3, 3 + items.length + 50, 3, 4),
      rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true },
    },
  }];
  await batchUpdate(id, valReqs, 'packing-checkboxes');

  console.log('Packing Checklist built with', items.length, 'items');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
