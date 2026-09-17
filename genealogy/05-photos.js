'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PH = sheetMap['📸 Photos & Notes'];
const FM_NAME = "'👨‍👩‍👧‍👦 Family Members'";

// Columns: A=ID, B=Name (formula), C=Photo URL (hidden helper), D=Photo Link (HYPERLINK formula), E=Personality/Memory Note
// Display: A, B, D, E (C is the raw URL col, kept narrow)

const HEADERS = ['ID','Name','Photo URL (paste here)','Photo Link','Personality / Memory Note'];

// name formula per row
const nameFormula = r => `=IFERROR(INDEX(${FM_NAME}!B:B,MATCH(A${r},${FM_NAME}!A:A,0)),"")`;
const linkFormula = r => `=IF(C${r}="","",HYPERLINK(C${r},"📸 View Photo"))`;

const ROWS = [
  // [ID, url, note]
  [1,  'https://example.com/photos/patrick-sullivan.jpg', 'Pat had a booming laugh you could hear from the next street over. Always smelled of pipe tobacco and river water.'],
  [2,  'https://example.com/photos/brigid-flanagan.jpg',  'Biddy kept a small statue of St. Brigid on the kitchen windowsill her whole life. Her soda bread recipe is still used every Christmas.'],
  [3,  'https://example.com/photos/jozef-kowalski.jpg',   'Józek could build anything from wood. His dovetail joints were so precise they needed no glue — the cabinet still stands in Aunt Eleanor\'s hallway.'],
  [5,  'https://example.com/photos/taro-nakamura.jpg',    'Taro grew the most extraordinary chrysanthemums. Grandmother Yuki kept a pressed one in her Bible for the rest of her life.'],
  [6,  'https://example.com/photos/yuki-tanaka.jpg',      'Yuki crossed the Pacific with almost nothing and built a life with an extraordinary quiet stubbornness. She learned English from library books.'],
  [7,  'https://example.com/photos/wilhelm-fischer.jpg',  'Willi wrote letters to his students even after they graduated. Dozens of them wrote back after his death, saying he changed the course of their lives.'],
  [9,  'https://example.com/photos/thomas-sullivan.jpg',  'Tommy never talked about the war. But every year on August 15 he went quiet and sat alone on the back porch for exactly one hour.'],
  [10, 'https://example.com/photos/eleanor-kowalski.jpg', 'Ellie\'s pierogies were legendary — she always made extra so there\'d be enough for the neighbors. In her last years she forgot names but never forgot the recipe.'],
  [11, 'https://example.com/photos/kenji-nakamura.jpg',   'Ken designed the electrical system for the house he bought after internment as a deliberate act of ownership. He was very proud of it.'],
  [12, 'https://example.com/photos/ruth-fischer.jpg',     'Ruthie played her violin at every family event until she was 80. She said music was the only language with no accent.'],
  [19, 'https://example.com/photos/michael-sullivan.jpg', 'Dad coached Little League for 20 years and never once cared if the team won. He just wanted everyone to feel proud of themselves afterward.'],
  [20, 'https://example.com/photos/carol-nakamura.jpg',   'Mom keeps every birthday card anyone has ever sent her in shoeboxes organized by year. There are forty-three boxes.'],
  [25, 'https://example.com/photos/daniel-sullivan.jpg',  'Danny\'s photo on the mantle shows him laughing at something off-camera. Nobody knows what it was, which somehow feels right.'],
  [31, 'https://example.com/photos/sarah-sullivan.jpg',   'Sarah (that\'s you!) — the family researcher. You inherited Patrick\'s laugh and Ellie\'s stubbornness. Both have served you well.'],
  [32, 'https://example.com/photos/david-park.jpg',       'Dave can fix anything mechanical and refuses to admit when something is genuinely beyond repair. He once spent six hours on a clock that cost four dollars.'],
];

(async () => {
  const data = [{ range: "'📸 Photos & Notes'!A1:E1", values: [HEADERS] }];

  ROWS.forEach(([personId, url, note], i) => {
    const r = i + 2;
    data.push({ range: `'📸 Photos & Notes'!A${r}`, values: [[personId]] });
    data.push({ range: `'📸 Photos & Notes'!B${r}`, values: [[nameFormula(r)]] });
    data.push({ range: `'📸 Photos & Notes'!C${r}`, values: [[url]] });
    data.push({ range: `'📸 Photos & Notes'!D${r}`, values: [[linkFormula(r)]] });
    data.push({ range: `'📸 Photos & Notes'!E${r}`, values: [[note]] });
  });

  await valuesBatchUpdate(id, data, 'ph-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(PH, 0, 1, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 16; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(PH, r, r+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, verticalAlignment: 'TOP' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ID col
  reqs.push({ repeatCell: { range: gridRange(PH, 1, 17, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  // Formula cols B & D — pale sepia
  reqs.push({ repeatCell: { range: gridRange(PH, 1, 17, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(PH, 1, 17, 3, 4), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  // URL col C — input
  reqs.push({ repeatCell: { range: gridRange(PH, 1, 17, 2, 3), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  // Note col E — wrap
  reqs.push({ repeatCell: { range: gridRange(PH, 1, 17, 4, 5), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  // Column widths: ID=40, Name=160, URL=200, Link=100, Note=380
  [[0,40],[1,160],[2,200],[3,100],[4,380]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: PH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row height header + data rows
  reqs.push({ updateDimensionProperties: { range: { sheetId: PH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PH, dimension: 'ROWS', startIndex: 1, endIndex: 17 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'ph-format');
  console.log('Photos & Notes complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
