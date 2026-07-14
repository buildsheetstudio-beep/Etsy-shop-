'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GRT = sheetMap['🎁 Gift Recipient Tracker'];

// Columns (0-indexed):
//   A(0): Recipient Name   B(1): Relationship    C(2): Gift Idea
//   D(3): Store/Source     E(4): Budget ($)      F(5): Actual Cost ($)
//   G(6): Diff ($)         H(7): Bought?         I(8): Wrapped?
//   J(9): Delivered/Sent?  K(10): Delivery Method L(11): Notes
//   M(12): Priority

const headers = ['Recipient Name','Relationship','Gift Idea','Store / Source','Budget ($)','Actual Cost ($)','Over/Under ($)','Bought?','Wrapped?','Delivered?','Delivery Method','Notes','Priority'];

const sampleData = [
  ['Sarah (wife)','Spouse/Partner','Cashmere sweater set','Nordstrom',150,142,'',true,true,true,'In Person','She mentioned wanting this in October','High'],
  ['Dad','Parent','Electric drill set','Home Depot',80,74.99,'',true,true,true,'In Person','','High'],
  ['Mom','Parent','Cookbook + cooking class','Amazon',60,67.50,'',true,false,false,'Ship to Home','Cookbook arrived, class gift card pending','High'],
  ['Jake (son)','Child','LEGO Technic set','Target',65,62.99,'',true,true,false,'In Person','Hiding in closet','High'],
  ['Emma (daughter)','Child','Art supply kit','Michaels',55,58.00,'',true,true,false,'In Person','','High'],
  ['Grandma Rose','Grandparent','Photo book + frame set','Shutterfly',45,49.99,'',true,false,false,'Ship to Home','Ships Dec 15','High'],
  ['Grandpa Joe','Grandparent','Fishing lure set','Bass Pro',40,38.50,'',true,true,true,'In Person','','Medium'],
  ['Sister Lily','Sibling','Spa day gift card','Spa Luxe',70,70,'',true,false,false,'In Person','','High'],
  ['Brother Mike','Sibling','Wireless earbuds','Best Buy',80,79.99,'',true,false,false,'Ship to Home','','Medium'],
  ['Aunt Carol','Parent','Wine & cheese basket','Harry & David',55,62,'',true,false,false,'Ship to Home','Over budget — upgraded','Medium'],
  ['Uncle Dave','Parent','Golf gloves + balls','Golf Galaxy',45,43,'',true,true,true,'In Person','','Medium'],
  ['BFF Tina','Friend','Candle gift set','Etsy',35,32,'',true,true,true,'In Person','Hand-delivered at party','Medium'],
  ['Neighbor Sue','Neighbor','Holiday cookie tin','Homemade',20,18.50,'',true,false,false,'In Person','Baking Dec 22','Low'],
  ['Boss Karen','Coworker','Plant + planter set','Trader Joe\'s',30,27,'',true,true,true,'In Person','Left on desk','Medium'],
  ['Team gift (work)','Coworker','Coffee shop gift cards x5','Starbucks',50,50,'',true,false,false,'In Person','Giving at holiday party','Medium'],
  ['Cousin Rachel','Sibling','Yoga mat + blocks','Amazon',40,37.49,'',true,false,false,'Ship to Home','','Low'],
  ['Cousin Tom','Sibling','Video game','GameStop',60,59.99,'',true,false,false,'Ship to Home','','Low'],
  ['Baby nephew Leo','Child','Plush toy set','Amazon',25,22,'',true,false,false,'Ship to Home','','High'],
  ['Dog sitter Pam','Friend','Gift basket','TJ Maxx',30,28,'',true,true,true,'In Person','','Low'],
  ['Mail carrier','Other','Gift card','Visa',15,15,'',true,false,false,'In Person','Leave in mailbox','Low'],
  ['Teacher Ms. Parks','Other','Classroom supply box','Staples',25,24.50,'',true,true,false,'In Person','','Medium'],
  ['Hairdresser','Other','Cash tip + card','N/A',40,40,'',true,false,false,'In Person','','Medium'],
  ['Plumber (holiday tip)','Other','Cash','N/A',25,25,'',false,false,false,'In Person','Check if we have cash','Low'],
  ['Secret Santa draw','Friend','Board game','Amazon',20,22,'',true,false,false,'Ship to Home','Over by $2','Low'],
  ['Office raffle prize','Coworker','Candle set','TJ Maxx',15,14.25,'',true,false,false,'In Person','','Low'],
];

(async () => {
  // Diff formula for rows 2-26 (idx 1-25)
  const diffFormula = (r) => `=IFERROR(F${r}-E${r},"")`;

  // Build value rows — insert formula string for col G, override booleans
  const valueRows = sampleData.map((row, i) => {
    const r = i + 2;
    return [
      row[0], row[1], row[2], row[3],
      row[4], row[5],
      diffFormula(r),
      row[7], row[8], row[9],
      row[10], row[11], row[12],
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'🎁 Gift Recipient Tracker'!A1", values: [headers] },
    { range: "'🎁 Gift Recipient Tracker'!A2", values: valueRows },
  ], 'grt-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];
  const numRows = sampleData.length;

  // Header row (row 1, idx 0)
  reqs.push({
    repeatCell: {
      range: gridRange(GRT, 0, 1, 0, 13),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepCranberry),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // Data rows alternating
  for (let r = 1; r <= numRows; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(GRT, r, r + 1, 0, 13),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 1 ? C.snowWhite : C.altRow),
            textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Col G (Diff) — formula col, subtle bg
  reqs.push({
    repeatCell: {
      range: gridRange(GRT, 1, numRows + 1, 6, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.formulaBg),
          textFormat: { italic: true },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00;[Red]"-$"#,##0.00' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // Col E, F — currency format
  reqs.push({
    repeatCell: {
      range: gridRange(GRT, 1, numRows + 1, 4, 6),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Cols H, I, J (checkboxes) — center
  reqs.push({
    repeatCell: {
      range: gridRange(GRT, 1, numRows + 1, 7, 10),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment)',
    },
  });

  // Col A — bold recipient names
  reqs.push({
    repeatCell: {
      range: gridRange(GRT, 1, numRows + 1, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat(textFormat)',
    },
  });

  // Column widths
  const widths = [150, 120, 180, 130, 90, 100, 90, 70, 70, 80, 120, 160, 80];
  widths.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: GRT, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row height for header
  reqs.push({ updateDimensionProperties: { range: { sheetId: GRT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'grt-format');
  console.log('Gift Recipient Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
