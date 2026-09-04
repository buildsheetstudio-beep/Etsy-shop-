'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CDL = sheetMap['💌 Holiday Card & Mailing List'];

// A Recipient/Family Name  B Address  C Card Status  D Postage Cost ($)  E Notes
// Card Status options: Not Sent | Sent | Received (from Reference D2:D4)
// Partial overlap with GRT names; some card-only contacts

const headers = ['Recipient / Family Name','Mailing Address','Card Status','Postage ($)','Notes'];

// 20 rows — overlaps some GRT names, adds card-only contacts
const rows = [
  ['Grandma & Grandpa Rose', '4512 Maple Ln, Portland OR 97201',           'Received',  0.68, 'Their card arrived Dec 6! Beautiful photo card'],
  ['Sister Lily',            '119 Fir Circle, Hillsboro OR 97124',           'Received',  0.68, 'She sent one back with cute stickers 🎄'],
  ['BFF Tina',               '730 Oak Ave, Portland OR 97202',               'Received',  0.68, 'She always sends the nicest cards'],
  ['Neighbor Sue',           'Next door — hand delivered',                   'Received',  0,    'Included a recipe card with it'],
  ['Church friend Linda',    '77 Cedar St, Tigard OR 97223',                 'Received',  0.68, 'Beautiful watercolor card from her'],
  ['Aunt Carol & Uncle Dave','221 Birchwood Dr, Salem OR 97301',             'Sent',      0.68, 'Sent Dec 5 — tracking shows delivered'],
  ['Brother Mike',           '550 Pine St, Gresham OR 97030',                'Sent',      0.68, ''],
  ['Cousin Rachel & Tom',    '89 Elm St Apt 3B, Seattle WA 98101',           'Sent',      0.68, 'Shared a card since same address'],
  ['College friend Jamie',   '15 Park Blvd, Eugene OR 97401',                'Sent',      0.68, ''],
  ['Old neighbor Mr & Mrs Chen','6 Chestnut Ct, Beaverton OR 97005',         'Sent',      0.68, 'Every year without fail'],
  ['Church pastor & family', '200 Grace Ave, Portland OR 97210',             'Sent',      0.68, ''],
  ['Mom\'s best friend Barb','334 Walnut Ave, Lake Oswego OR 97034',         'Sent',      0,    'Mom hand-delivering ours and hers together'],
  ['Childhood friend Kevin', '2200 Spruce Rd, Medford OR 97501',             'Sent',      0.68, 'Reconnected at reunion this year'],
  ['Work friend Dana',       'Office — hand delivered',                      'Sent',      0,    'Gave it at the holiday party'],
  ['Boss Karen',             'Office — hand delivered',                      'Sent',      0,    ''],
  ['College roommate Sophie','18 Harbor View Dr, Portland OR 97201',         'Not Sent',  0,    'Need to find new address — she moved'],
  ['Dentist Dr Kim',         '1400 Medical Plaza, Portland OR 97204',        'Not Sent',  0,    'Skipping this year'],
  ['PTA co-chair Beth',      '912 Hemlock Dr, Portland OR 97229',            'Not Sent',  0,    'Will hand-deliver at the PTA meeting Dec 18'],
  ['Great-aunt Mae',         '501 Birch St, Roseburg OR 97471',              'Not Sent',  0.68, 'Need to mail by Dec 18 for on-time delivery'],
  ['Pen pal in London (Amy)','14 Grosvenor Sq, London W1K 6JP, UK',         'Not Sent',  1.50, 'International — use airmail stamp'],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'💌 Holiday Card & Mailing List'!A1", values: [headers] },
    { range: "'💌 Holiday Card & Mailing List'!A2", values: rows },
  ], 'card-values');

  const reqs = [];
  const n = rows.length;

  // Header
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

  for (let r = 1; r <= n; r++) {
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

  reqs.push({ repeatCell: { range: gridRange(CDL, 1, n + 1, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat(textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(CDL, 1, n + 1, 2, 3), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(CDL, 1, n + 1, 3, 4), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  [190, 270, 100, 90, 210].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CDL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CDL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'card-format');
  console.log('Holiday Card & Mailing List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
