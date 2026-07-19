'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const LR = sheetMap['📞 Living Relatives Directory'];

// Columns: A=Name, B=Relationship to User, C=Phone, D=Email, E=Linked FM ID, F=Last Contact Date, G=Notes
const HEADERS = ['Name','Relationship to User','Phone','Email','Linked Family Member ID','Last Contact Date','Notes'];

const ROWS = [
  ['Carol Jean Nakamura',    'Parent',      '+1-617-555-0192', 'carol.nakamura@example.com',    20, '2026-06-15', 'Calls every Sunday. Best source for Nakamura family stories.'],
  ['Christopher Sullivan',   'Sibling',     '+1-617-555-0284', 'chris.sullivan@example.com',    33, '2026-07-01', 'Has boxes of old photos in his attic — ask to scan them.'],
  ['Jennifer Torres',        'Sibling',     '+1-305-555-0143', 'jen.torres@example.com',        34, '2026-04-10', 'Lives in Miami. Fluent Spanish — contact Ramirez line.'],
  ['Patricia Ann Sullivan',  'Parent',      '+1-617-555-0371', 'patty.sullivan@example.com',    21, '2026-07-10', 'Has Dad\'s military discharge papers and letters.'],
  ['Susan Marie Sullivan',   'Aunt/Uncle',  '+1-602-555-0517', 'sue.ramirez@example.com',       26, '2025-10-20', 'Last contacted October 2025. Has Danny\'s Vietnam letters.'],
  ['James Kevin Sullivan',   'Aunt/Uncle',  '+1-617-555-0628', 'jimmy.sullivan@example.com',    23, '2026-03-14', 'Hardware store owner. Good source for Sullivan Boston history.'],
  ['Barbara Ann Kelly',      'Aunt/Uncle',  '+1-617-555-0739', 'barb.kelly@example.com',        24, '2025-09-05', 'Hasn\'t replied to last email. Try calling.'],
  ['Kevin Robert Greene',    'Cousin',      '+1-401-555-0840', 'kev.greene@example.com',        35, '2025-12-25', 'Christmas card exchange. Has photos of Robert\'s construction projects.'],
  ['Amy Lynn Greene',        'Cousin',      '+1-401-555-0951', 'amy.greene@example.com',        36, '2026-02-18', 'Graphic designer — helped design this spreadsheet template!'],
  ['Thomas James Sullivan',  'Cousin',      '+1-617-555-0162', 'tommyjr.sullivan@example.com',  37, '2026-01-03', 'New Year\'s catch-up call. Reminds everyone of grandfather Tommy.'],
  ['Emily Grace Sullivan',   'Cousin',      '+1-617-555-0273', 'emily.sullivan@example.com',    38, '2026-06-30', 'Therapist. Very interested in family psychology/trauma history.'],
  ['Hiroshi Nakamura',       'Grandparent', '+1-213-555-0384', 'hiroshi.nakamura@example.com',  17, '2026-05-20', 'Grandfather. Sharp memory. Should be interviewed on video before health declines.'],
  ['Linda Sato',             'Grandparent', '+1-213-555-0495', 'linda.sato@example.com',        18, '2026-05-20', 'Grandmother. Has Yuki\'s Bible with pressed chrysanthemum.'],
  ['Alan James Nakamura',    'Cousin',      '+1-213-555-0506', 'alan.nakamura@example.com',     30, '2025-11-14', 'Uncle technically (mom\'s brother). Has business records from 1970s-80s.'],
  ['Mia Yuki Nakamura',      'Cousin',      '+1-213-555-0617', 'mia.nakamura@example.com',      40, '2026-07-05', 'Second cousin once removed, practically. Great at digital archiving.'],
];

(async () => {
  const data = [{ range: "'📞 Living Relatives Directory'!A1:G1", values: [HEADERS] }];
  data.push({ range: "'📞 Living Relatives Directory'!A2:G16", values: ROWS });
  await valuesBatchUpdate(id, data, 'lr-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(LR, 0, 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 15; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(LR, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  // Linked ID col — centered
  reqs.push({ repeatCell: { range: gridRange(LR, 1, 16, 4, 5), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
  // Notes col — wrap
  reqs.push({ repeatCell: { range: gridRange(LR, 1, 16, 6, 7), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,170],[1,120],[2,130],[3,200],[4,120],[5,110],[6,300]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: LR, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: LR, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: LR, dimension: 'ROWS', startIndex: 1, endIndex: 16 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'lr-format');
  console.log('Living Relatives Directory complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
