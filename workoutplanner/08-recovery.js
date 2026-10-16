'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REC = sheetMap['😴 Rest & Recovery'];

// Columns: A=Date, B=Sleep(hrs), C=Sleep Quality, D=Soreness Level,
//          E=Sore Areas, F=Rest Day(checkbox), G=Notes
// Row 1: title, Row 2: headers, Rows 3-32: 30 days (Jul 1-30)

const recoveryData = [
  ['2026-07-01',7.5,'Good','Mild','Chest',false,''],
  ['2026-07-02',8.0,'Excellent','Moderate','Back, Lower Back',false,'Great workout day'],
  ['2026-07-03',7.0,'Good','Moderate','Legs, Glutes',false,'Heavy squats'],
  ['2026-07-04',8.5,'Excellent','Mild','Legs',false,'Active recovery run'],
  ['2026-07-05',9.0,'Excellent','None','',true,'Rest day'],
  ['2026-07-06',7.5,'Good','Mild','Shoulders',false,''],
  ['2026-07-07',6.5,'Fair','None','',false,'Late night — slept poorly'],
  ['2026-07-08',7.5,'Good','Mild','Chest, Shoulders',false,''],
  ['2026-07-09',8.0,'Good','Moderate','Back',false,''],
  ['2026-07-10',7.0,'Good','Moderate','Legs',false,'Heavy leg press'],
  ['2026-07-11',7.5,'Good','Mild','Core',false,''],
  ['2026-07-12',8.0,'Excellent','None','',true,'Rest day'],
  ['2026-07-13',7.5,'Good','Mild','Glutes',false,'Hip thrust PR day'],
  ['2026-07-14',8.0,'Good','None','',false,''],
  ['2026-07-15',6.0,'Fair','Mild','Arms',false,'Work stress, poor sleep'],
  ['2026-07-16',7.5,'Good','Moderate','Legs',false,'RDL day'],
  ['2026-07-17',8.5,'Excellent','Severe','Chest, Shoulders',false,'Bench PR — very sore'],
  ['2026-07-18',7.0,'Good','Moderate','Chest',false,'Day after PR'],
  ['2026-07-19',8.5,'Excellent','Mild','',true,'Rest day — cardio only'],
  ['2026-07-20',7.5,'Good','None','',false,''],
  ['2026-07-21',8.0,'Good','Mild','Back',false,''],
  ['2026-07-22',7.0,'Good','Moderate','Legs',false,'Squat PR!'],
  ['2026-07-23',7.5,'Good','Mild','Core',false,''],
  ['2026-07-24',8.0,'Excellent','Moderate','Full Body',false,'Circuit day'],
  ['2026-07-25',8.5,'Excellent','Mild','',false,'Stretch & foam roll'],
  ['2026-07-26',9.0,'Excellent','None','',true,'Rest day'],
  ['2026-07-27',7.5,'Good','None','',false,''],
  ['2026-07-28',7.0,'Good','None','',false,''],
  ['2026-07-29',7.5,'Good','None','',false,''],
  ['2026-07-30',8.0,'Good','None','',false,''],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'😴 Rest & Recovery'!A1", values: [['😴']] },
    { range: "'😴 Rest & Recovery'!B1", values: [['Rest & Recovery Log']] },
    { range: "'😴 Rest & Recovery'!A2", values: [['Date','Sleep (hrs)','Sleep Quality','Soreness Level','Sore Areas','Rest Day','Notes']] },
    { range: "'😴 Rest & Recovery'!A3", values: recoveryData },
  ], 'rec-data');

  const reqs = [];

  reqs.push({ mergeCells: { range: gridRange(REC, 0, 1, 1, 7), mergeType: 'MERGE_ALL' } });

  reqs.push({
    repeatCell: {
      range: gridRange(REC, 0, 1, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.purple),
          textFormat: { foregroundColor: hex(C.white), fontSize: 16, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(REC, 0, 1, 1, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.purple),
          textFormat: { foregroundColor: hex(C.white), fontSize: 13, bold: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(REC, 1, 2, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.darkGray),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 2; r < 32; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(REC, r, r + 1, 0, 7),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 0 ? C.offWhite : C.altRow),
            textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Date format
  reqs.push({
    repeatCell: {
      range: gridRange(REC, 2, 32, 0, 1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Center numeric/dropdown cols B-F
  reqs.push({
    repeatCell: {
      range: gridRange(REC, 2, 32, 1, 6),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(horizontalAlignment)',
    },
  });

  // Sleep hours: 1 decimal
  reqs.push({
    repeatCell: {
      range: gridRange(REC, 2, 32, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } },
      fields: 'userEnteredFormat(numberFormat)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: REC, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: REC, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: REC, dimension: 'ROWS', startIndex: 2, endIndex: 32 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Column widths
  [[0,100],[1,85],[2,110],[3,110],[4,145],[5,75],[6,180]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REC, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'rec-format');
  console.log('Rest & Recovery complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
