'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WSH = sheetMap['💡 Gift Idea Wishlist'];

// Columns: A Recipient  B Gift Idea  C Est. Price ($)  D Where to Buy  E Priority  F Notes  G Selected?

const headers = ['Recipient','Gift Idea / Brainstorm','Est. Price ($)','Where to Buy','Priority','Notes','Selected?'];

const sampleData = [
  ['Sarah (wife)','Cashmere sweater set',145,'Nordstrom','High','She mentioned this in October',true],
  ['Sarah (wife)','Silk robe',85,'Amazon','Medium','Backup idea',false],
  ['Sarah (wife)','Jewelry organizer',50,'Amazon','Low','Nice but not exciting',false],
  ['Dad','Electric drill set',75,'Home Depot','High','His old one broke',true],
  ['Dad','Recliner cushion',40,'Amazon','Medium','',false],
  ['Mom','Cookbook + cooking class',65,'Amazon/Local','High','Two-part gift — both ordered',true],
  ['Mom','Herb garden kit',35,'Amazon','Low','Backup',false],
  ['Jake (son)','LEGO Technic set',60,'Target','High','He circled in catalog',true],
  ['Jake (son)','RC car',50,'Target','Medium','Second choice',false],
  ['Emma (daughter)','Art supply kit',55,'Michaels','High','She asked for this',true],
  ['Emma (daughter)','Craft jewelry kit',30,'Amazon','Low','Good stocking stuffer option',false],
  ['Grandma Rose','Photo book + frame set',50,'Shutterfly','High','Family photos from 2025',true],
  ['Grandpa Joe','Fishing lure set',40,'Bass Pro','High','He fishes every weekend',true],
  ['Sister Lily','Spa day gift card',70,'Spa Luxe','High','Her favorite place',true],
  ['Brother Mike','Wireless earbuds',80,'Best Buy','High','He lost his last pair',true],
  ['Brother Mike','Portable charger',30,'Amazon','Medium','Good backup',false],
  ['Aunt Carol','Wine & cheese basket',55,'Harry & David','Medium','Upgraded to $62 version',true],
  ['Uncle Dave','Golf gloves + balls',45,'Golf Galaxy','Medium','',true],
  ['BFF Tina','Candle gift set',32,'Etsy','Medium','From her favorite shop',true],
  ['Neighbor Sue','Holiday cookie tin',20,'Homemade','Low','Baking Dec 22',true],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'💡 Gift Idea Wishlist'!A1", values: [headers] },
    { range: "'💡 Gift Idea Wishlist'!A2", values: sampleData },
  ], 'wishlist-values');

  const reqs = [];
  const numRows = sampleData.length;

  // Header row
  reqs.push({
    repeatCell: {
      range: gridRange(WSH, 0, 1, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepGreen),
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
        range: gridRange(WSH, r, r + 1, 0, 7),
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

  // Col A — bold recipient
  reqs.push({
    repeatCell: {
      range: gridRange(WSH, 1, numRows + 1, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat(textFormat)',
    },
  });

  // Col C — currency format
  reqs.push({
    repeatCell: {
      range: gridRange(WSH, 1, numRows + 1, 2, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Col G — center checkboxes
  reqs.push({
    repeatCell: {
      range: gridRange(WSH, 1, numRows + 1, 6, 7),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment)',
    },
  });

  // Column widths
  [140, 200, 90, 150, 80, 180, 80].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: WSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: WSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'wishlist-format');
  console.log('Gift Idea Wishlist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
