'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BMT = sheetMap['📏 Body Measurements'];

// Columns: A=Date, B=Weight(lbs), C=Body Fat%, D=Chest(in), E=Waist(in),
//          F=Hips(in), G=Biceps(in), H=Thighs(in), I=Photo Link, J=Notes
// Row 1: title, Row 2: headers, Rows 3-10: 8 bi-weekly check-ins

const checkIns = [
  // [date, weight, bodyFat%, chest, waist, hips, biceps, thighs, photoLink, notes]
  ['2026-03-22', 195.0, 22.5, 42.0, 36.5, 41.0, 14.5, 24.0, '', 'Starting measurements'],
  ['2026-04-05', 193.5, 22.0, 42.2, 36.0, 40.8, 14.7, 23.8, '', ''],
  ['2026-04-19', 191.8, 21.5, 42.5, 35.5, 40.5, 14.9, 23.5, '', 'Down 3.2lbs'],
  ['2026-05-03', 190.2, 21.0, 42.8, 35.0, 40.2, 15.1, 23.2, '', ''],
  ['2026-05-17', 188.5, 20.5, 43.0, 34.5, 40.0, 15.3, 23.0, '', 'Feeling stronger'],
  ['2026-05-31', 187.0, 20.0, 43.2, 34.0, 39.8, 15.5, 22.8, '', ''],
  ['2026-06-14', 185.5, 19.5, 43.5, 33.5, 39.5, 15.7, 22.5, '', 'Best readings yet!'],
  ['2026-06-28', 184.0, 19.0, 43.8, 33.0, 39.2, 15.9, 22.2, '', '-11lbs / -3.5% BF'],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📏 Body Measurements'!A1", values: [['📏']] },
    { range: "'📏 Body Measurements'!B1", values: [['Body Measurements & Progress Log']] },
    { range: "'📏 Body Measurements'!A2", values: [['Date','Weight (lbs)','Body Fat %','Chest (in)','Waist (in)','Hips (in)','Biceps (in)','Thighs (in)','Photo Link','Notes']] },
    { range: "'📏 Body Measurements'!A3", values: checkIns },
  ], 'bmt-data');

  const reqs = [];

  reqs.push({ mergeCells: { range: gridRange(BMT, 0, 1, 1, 10), mergeType: 'MERGE_ALL' } });

  reqs.push({
    repeatCell: {
      range: gridRange(BMT, 0, 1, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.coral),
          textFormat: { foregroundColor: hex(C.white), fontSize: 16, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(BMT, 0, 1, 1, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.coral),
          textFormat: { foregroundColor: hex(C.white), fontSize: 13, bold: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(BMT, 1, 2, 0, 10),
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

  for (let r = 2; r < 10; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(BMT, r, r + 1, 0, 10),
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
      range: gridRange(BMT, 2, 10, 0, 1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Center numeric cols B-H
  reqs.push({
    repeatCell: {
      range: gridRange(BMT, 2, 10, 1, 8),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(horizontalAlignment)',
    },
  });

  // % format for Body Fat col C
  reqs.push({
    repeatCell: {
      range: gridRange(BMT, 2, 10, 2, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0"%"' } } },
      fields: 'userEnteredFormat(numberFormat)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: BMT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BMT, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BMT, dimension: 'ROWS', startIndex: 2, endIndex: 10 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths
  [[0,100],[1,90],[2,80],[3,80],[4,80],[5,80],[6,85],[7,90],[8,140],[9,180]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: BMT, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'bmt-format');
  console.log('Body Measurements complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
