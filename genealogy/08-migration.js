'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const MIG = sheetMap['🚢 Immigration & Migration Tracker'];

// Columns: A=Person ID, B=Migration Type, C=Origin Location, D=Destination Location,
//          E=Date, F=Ship/Vessel Name, G=Port of Entry, H=Notes
const HEADERS = ['Person ID','Migration Type','Origin Location','Destination Location','Date','Ship / Vessel Name','Port of Entry','Notes'];

const ROWS = [
  [1, 'Immigration',       'County Cork, Ireland',   'Boston, MA, USA',          '1901-10-14','SS Campania',       'Boston, MA',      'Name spelled "Sulivan" on manifest. Arrived age 23.'],
  [2, 'Immigration',       'County Kerry, Ireland',  'New York, NY, USA',        '1904-06-03','SS Mauretania',     'New York (Ellis Island)','Listed as domestic servant, destination: Boston.'],
  [3, 'Emigration',        'Kraków, Poland',         'Warsaw, Poland',           '1939-09-05','','',                'Moved after German invasion of Poland.'],
  [4, 'Immigration',       'Łódź, Poland',           'Chicago, IL, USA',         '1921-03-22','SS Kroonland',      'New York (Ellis Island)','Traveled alone with two children after the war.'],
  [6, 'Immigration',       'Osaka, Japan',           'Honolulu, HI, USA',        '1920-04-08','SS Chiyo Maru',     'Honolulu, HI',    'Picture bride arrangement. Brought infant Kenji.'],
  [8, 'Immigration',       'Stuttgart, Germany',     'New York, NY, USA',        '1938-07-21','SS Manhattan',      'New York',        'Fled National Socialist Germany. Widowed 3 years.'],
  [11,'Internal Migration','Honolulu, HI, USA',      'Los Angeles, CA, USA',     '1944-11-10','','',                'Post-internment resettlement. Interned at Manzanar 1942-44.'],
  [17,'Internal Migration','Manzanar, CA, USA',      'Los Angeles, CA, USA',     '1944-11-10','','',                'Born at Manzanar; moved with family after internment ended.'],
  [20,'Internal Migration','Los Angeles, CA, USA',   'Boston, MA, USA',          '1965-08-01','','',                'Moved for nursing program at Boston City Hospital.'],
  [26,'Internal Migration','Boston, MA, USA',        'Tucson, AZ, USA',          '1985-05-15','','',                'Relocated for husband\'s work.'],
  [31,'Internal Migration','Boston, MA, USA',        'Cambridge, MA, USA',       '1992-09-01','','',                'Moved for graduate studies at Harvard.'],
  [32,'Immigration',       'Portland, OR, USA',      'Cambridge, MA, USA',       '1993-06-01','','',                'Moved after meeting Sarah at academic conference.'],
];

(async () => {
  const data = [{ range: "'🚢 Immigration & Migration Tracker'!A1:H1", values: [HEADERS] }];
  data.push({ range: "'🚢 Immigration & Migration Tracker'!A2:H13", values: ROWS });
  await valuesBatchUpdate(id, data, 'mig-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(MIG, 0, 1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 12; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(MIG, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  // ID col
  reqs.push({ repeatCell: { range: gridRange(MIG, 1, 13, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  // Notes col — wrap
  reqs.push({ repeatCell: { range: gridRange(MIG, 1, 13, 7, 8), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,70],[1,120],[2,160],[3,160],[4,100],[5,130],[6,150],[7,260]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: MIG, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: MIG, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: MIG, dimension: 'ROWS', startIndex: 1, endIndex: 13 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'mig-format');
  console.log('Immigration & Migration Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
