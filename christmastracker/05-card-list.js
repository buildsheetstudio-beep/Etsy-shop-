'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CDL = sheetMap['💌 Holiday Card & Mailing List'];

// Columns: A Name  B Address  C Card Status  D Card Received?  E Notes

const headers = ['Name','Mailing Address','Card Status','Card Received?','Notes'];

const sampleData = [
  ['Grandma Rose','4512 Maple Lane, Portland OR 97201','Sent',true,'Sent Dec 5'],
  ['Grandpa Joe','4512 Maple Lane, Portland OR 97201','Sent',true,'Same address as Grandma'],
  ['Uncle Frank & Aunt Carol','221 Birchwood Dr, Salem OR 97301','Sent',false,''],
  ['Cousin Rachel','89 Elm St Apt 3B, Seattle WA 98101','Sent',false,''],
  ['Cousin Tom','89 Elm St Apt 3B, Seattle WA 98101','Sent',false,'Same building as Rachel'],
  ['BFF Tina','730 Oak Ave, Portland OR 97202','Sent',true,'She sent one back! 🎄'],
  ['College friend Jamie','15 Park Blvd, Eugene OR 97401','Sent',false,''],
  ['Old neighbor Mr. & Mrs. Chen','6 Chestnut Ct, Beaverton OR 97005','Sent',false,''],
  ['Work friend Dana','Ship via email — no address on file','Not Sent',false,'Send digital card'],
  ['Childhood friend Kevin','2200 Spruce Rd, Medford OR 97501','Not Sent',false,'Need to find address'],
  ['Church friend Linda','77 Cedar St, Tigard OR 97223','Sent',true,'Beautiful card from her'],
  ['Mom\'s best friend Barb','334 Walnut Ave, Lake Oswego OR 97034','Not Sent',false,'Mom will hand-deliver'],
  ['Sister Lily','119 Fir Circle, Hillsboro OR 97124','Sent',true,''],
  ['Brother Mike','550 Pine St, Gresham OR 97030','Sent',false,''],
  ['Neighbor Sue','Next door — hand delivered','Sent',true,'She loved the cookie recipe card'],
  ['Boss Karen','Office — hand delivered','Sent',false,''],
  ['Dentist Dr. Kim','1400 Medical Plaza, Portland OR 97204','Not Sent',false,'Skip this year'],
  ['PTA co-chair Beth','912 Hemlock Dr, Portland OR 97229','Not Sent',false,'Send with Emma\'s teacher card'],
  ['Teacher Ms. Parks','Hand deliver with gift','Not Sent',false,''],
  ['Aunt Sue & Uncle Bill','88 Redwood Ave, Corvallis OR 97330','Sent',false,''],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'💌 Holiday Card & Mailing List'!A1", values: [headers] },
    { range: "'💌 Holiday Card & Mailing List'!A2", values: sampleData },
  ], 'card-values');

  const reqs = [];
  const numRows = sampleData.length;

  // Header row
  reqs.push({
    repeatCell: {
      range: gridRange(CDL, 0, 1, 0, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.trueRed),
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
        range: gridRange(CDL, r, r + 1, 0, 5),
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

  // Col A — bold name
  reqs.push({
    repeatCell: {
      range: gridRange(CDL, 1, numRows + 1, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat(textFormat)',
    },
  });

  // Col D — center checkboxes
  reqs.push({
    repeatCell: {
      range: gridRange(CDL, 1, numRows + 1, 3, 4),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment)',
    },
  });

  // Col C — center status
  reqs.push({
    repeatCell: {
      range: gridRange(CDL, 1, numRows + 1, 2, 3),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(horizontalAlignment)',
    },
  });

  // Column widths
  [160, 260, 100, 110, 180].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CDL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: CDL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'card-format');
  console.log('Holiday Card & Mailing List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
