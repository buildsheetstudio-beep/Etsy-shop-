'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SC = sheetMap['📜 Source & Citation Tracker'];

// Columns: A=Fact/Claim, B=Person ID, C=Source Type, D=Source Details, E=Confidence Level, F=Date Found, G=Notes
const HEADERS = ['Fact / Claim','Person ID','Source Type','Source Details','Confidence Level','Date Found','Notes'];

const ROWS = [
  // Patrick Sullivan (1)
  ['Birth Date',       1,  'Birth Certificate',    'Cork County Council birth register, Vol.4, p.22, 1878',   'Confirmed',  '2022-03-10', 'Original scanned from Cork Archives; legible.'],
  ['Emigration Year',  1,  'Immigration Record',   'Ellis Island manifest, SS Campania, Oct 14 1901, line 44','Confirmed',  '2022-03-12', 'Name spelled "Sulivan" — one L. Confirmed via age match.'],
  ['Death Date',       1,  'Death Certificate',    'Massachusetts Death Index, 1952, Cert. #47211',           'Confirmed',  '2022-03-15', 'Cause listed as cardiac arrest.'],
  // Brigid Flanagan (2)
  ['Birth Date',       2,  'Census Record',        '1901 Ireland Census, County Kerry, Townland Tralee',      'Probable',   '2022-04-01', 'Age listed as 19 in 1901 — implies ~1882. No birth cert found.'],
  ['Arrival USA',      2,  'Immigration Record',   'Ellis Island manifest, SS Mauretania, June 3 1904',       'Confirmed',  '2022-04-02', 'Listed as Brigid Flanagan, age 22, from Kerry.'],
  // Józef Kowalski (3)
  ['Birth Date',       3,  'Birth Certificate',    'Kraków parish register, St. Mary Basilica, Sept 1875',    'Confirmed',  '2023-01-08', 'Obtained from Polish State Archives via mail request.'],
  ['Death',            3,  'Interview/Oral History','Oral account from Eleanor Kowalski, recorded 1987',      'Probable',   '2023-01-10', 'Eleanor said he died "during the occupation." No document.'],
  // Taro Nakamura (5)
  ['Birth Date',       5,  'Census Record',        '1920 Japan National Census, Hiroshima Prefecture',        'Probable',   '2022-09-05', 'Record shows Taro Nakamura, age 35. Inferred birth year 1885.'],
  ['Death',            5,  'Newspaper',            'Chugoku Shimbun obituary supplement, Aug 1945',           'Uncertain',  '2022-09-07', 'Family believes he died 6 Aug 1945 in Hiroshima. No cert.'],
  // Yuki Tanaka (6)
  ['Immigration',      6,  'Immigration Record',   'Hawaii immigration record, SS Chiyo Maru, June 1920',     'Confirmed',  '2022-09-10', 'Lists Yuki Tanaka, 29, and infant Kenji Nakamura.'],
  // Thomas Sullivan (9)
  ['Military Service', 9,  'Military Record',      'WWII Selective Service Registration Card, Reg. #123-A',   'Confirmed',  '2022-06-14', 'Confirms rank, unit (Pacific). Available NARA online.'],
  ['Birth Date',       9,  'Birth Certificate',    'Massachusetts vital records, Cert. #1908-BOS-7741',        'Confirmed',  '2022-06-14', 'Suffolk County Registry of Deeds.'],
  // Eleanor Kowalski (10)
  ['Birth Date',       10, 'Birth Certificate',    'Illinois vital records, Cook County, April 1912',         'Confirmed',  '2022-07-01', 'Full original on file.'],
  ['Marriage',         10, 'Marriage Record',      'Suffolk County Marriage Index, 1930, Cert. #7202',        'Confirmed',  '2022-07-03', 'Lists Eleanor Kowalski and Thomas Sullivan, Oct 5 1930.'],
  // Kenji Nakamura (11)
  ['Internment',       11, 'Military Record',      'Manzanar War Relocation Authority records, Reg. #11K-249','Confirmed',  '2022-10-20', 'Full family card. Confirms internment 1942-1944.'],
  ['Birth Date',       11, 'Birth Certificate',    'Hawaii DOH birth record, Cert. #1915-HNL-5523',           'Confirmed',  '2022-10-20', 'Copy obtained from Hawaii Vital Records office.'],
  // Ruth Fischer (12)
  ['Immigration',      12, 'Immigration Record',   'NY passenger manifest, SS Manhattan, March 1938',         'Confirmed',  '2023-02-14', 'Ruth Miriam Fischer, age 18. Arrived Port of NY.'],
  // Daniel Sullivan (25)
  ['KIA',              25, 'Military Record',      'Vietnam Casualty Index, NARA, June 20 1972, ref. VN-7741','Confirmed',  '2022-11-01', 'Confirmed KIA, Quang Nam Province. Listed on Vietnam Wall.'],
  // Michael Sullivan (19)
  ['Birth Date',       19, 'Birth Certificate',    'Massachusetts vital records, Suffolk County, Aug 1940',    'Confirmed',  '2023-03-05', ''],
  ['Death Date',       19, 'Death Certificate',    'Massachusetts Death Index 2018, Cert. #2018-18-7791',     'Confirmed',  '2023-03-05', 'Cause: prostate carcinoma, stage IV.'],
  // Sarah Sullivan (31)
  ['Birth Date',       31, 'Birth Certificate',    'Massachusetts vital records, Suffolk County, May 1970',    'Confirmed',  '2021-01-01', 'Researcher\'s own birth certificate.'],
  // Hiroshi Nakamura (17)
  ['Birth Date',       17, 'Birth Certificate',    'Manzanar camp birth record, May 20 1940',                 'Confirmed',  '2022-10-21', 'Confirms birth during internment.'],
  // Wilhelm Fischer (7)
  ['Birth Date',       7,  'Census Record',        'Bavaria State Census 1900, Munich region, entry #4471',   'Probable',   '2023-04-10', 'Age listed as 28 in 1900 — implies ~1872. Corroborates family story.'],
  // Józef marriage (3+4)
  ['Marriage',         3,  'Marriage Record',      'Kraków civil registry, marriage register 1899, entry 77', 'Confirmed',  '2023-01-12', 'Lists Józef Kowalski and Maria Wisniewska, October 1899.'],
  // Generic uncertain
  ['Birth Year',       8,  'Interview/Oral History','Family memory — Hilde\'s daughter recalled "born around 1877"','Uncertain','2023-05-01','No document yet. Searching Stuttgart civil registry.'],
];

(async () => {
  const data = [{ range: "'📜 Source & Citation Tracker'!A1:G1", values: [HEADERS] }];
  data.push({ range: "'📜 Source & Citation Tracker'!A2:G26", values: ROWS });
  await valuesBatchUpdate(id, data, 'sc-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(SC, 0, 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 25; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(SC, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  // Person ID col — centered bold
  reqs.push({ repeatCell: { range: gridRange(SC, 1, 26, 1, 2), cell: { userEnteredFormat: { textFormat: { bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  // Notes col — wrap
  reqs.push({ repeatCell: { range: gridRange(SC, 1, 26, 6, 7), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });
  // Source details — wrap
  reqs.push({ repeatCell: { range: gridRange(SC, 1, 26, 3, 4), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,130],[1,70],[2,130],[3,260],[4,100],[5,100],[6,240]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SC, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: SC, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SC, dimension: 'ROWS', startIndex: 1, endIndex: 26 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'sc-format');
  console.log('Source & Citation Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
