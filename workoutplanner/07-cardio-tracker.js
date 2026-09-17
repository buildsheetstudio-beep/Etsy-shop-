'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CAR = sheetMap['🏃 Cardio Tracker'];

// Columns: A=Date, B=Cardio Type, C=Distance(mi), D=Duration(min),
//          E=Pace(min/mi, formula=D/C), F=Calories, G=Avg HR(bpm), H=Notes
// Row 1: title, Row 2: headers, Rows 3-12: 10 entries

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'🏃 Cardio Tracker'!A1", values: [['🏃']] },
    { range: "'🏃 Cardio Tracker'!B1", values: [['Cardio Tracker']] },
    { range: "'🏃 Cardio Tracker'!A2", values: [['Date','Cardio Type','Distance (mi)','Duration (min)','Pace (min/mi)','Calories','Avg HR (bpm)','Notes']] },
  ], 'car-headers');

  const rawData = [
    // [date, type, distance, duration, calories, heartRate, notes]
    ['2026-07-04','Running',   3.1, 28,  280, 155, '5K morning run'],
    ['2026-07-08','Cycling',  15.0, 52,  420, 142, 'Trail ride — flat'],
    ['2026-07-10','Rowing',    5.0, 25,  300, 160, 'Steady state erg'],
    ['2026-07-12','Jump Rope', 0.5, 20,  240, 172, '3min rounds × 5'],
    ['2026-07-15','Running',   6.2, 58,  530, 158, '10K long run'],
    ['2026-07-18','Cycling',  20.0, 68,  560, 145, 'Road ride'],
    ['2026-07-20','Elliptical',4.0, 35,  320, 148, 'Steady pace'],
    ['2026-07-22','Swimming',  1.0, 40,  380, 140, '40 laps'],
    ['2026-07-25','Running',   3.1, 27,  270, 153, '5K — 1min faster!'],
    ['2026-07-27','Jump Rope', 0.5, 15,  180, 168, 'Warm-up only'],
  ];

  const dataRows = rawData.map((row, i) => {
    const r = i + 3;
    return [
      row[0],                            // A: Date
      row[1],                            // B: Cardio Type
      row[2],                            // C: Distance
      row[3],                            // D: Duration
      `=IFERROR(D${r}/C${r},0)`,         // E: Pace = Duration ÷ Distance
      row[4],                            // F: Calories
      row[5],                            // G: Avg HR
      row[6],                            // H: Notes
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'🏃 Cardio Tracker'!A3", values: dataRows },
  ], 'car-data');

  const reqs = [];

  reqs.push({ mergeCells: { range: gridRange(CAR, 0, 1, 1, 8), mergeType: 'MERGE_ALL' } });

  reqs.push({
    repeatCell: {
      range: gridRange(CAR, 0, 1, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.teal),
          textFormat: { foregroundColor: hex(C.white), fontSize: 16, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(CAR, 0, 1, 1, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.teal),
          textFormat: { foregroundColor: hex(C.white), fontSize: 13, bold: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(CAR, 1, 2, 0, 8),
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

  for (let r = 2; r < 12; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(CAR, r, r + 1, 0, 8),
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

  // Pace column (E) = formula, formulaBg
  reqs.push({
    repeatCell: {
      range: gridRange(CAR, 2, 12, 4, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.formulaBg),
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'NUMBER', pattern: '0.00' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,numberFormat)',
    },
  });

  // Date format
  reqs.push({
    repeatCell: {
      range: gridRange(CAR, 2, 12, 0, 1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Center numeric cols C-G
  reqs.push({
    repeatCell: {
      range: gridRange(CAR, 2, 12, 2, 7),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(horizontalAlignment)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAR, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAR, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAR, dimension: 'ROWS', startIndex: 2, endIndex: 12 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Column widths
  [[0,100],[1,110],[2,95],[3,95],[4,95],[5,80],[6,90],[7,180]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CAR, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'car-format');
  console.log('Cardio Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
