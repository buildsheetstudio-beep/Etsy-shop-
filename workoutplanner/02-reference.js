'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

const muscleGroups  = ['Chest','Back','Shoulders','Arms','Core','Legs','Glutes','Full Body','Cardio','Mobility'];
const workoutTypes  = ['Strength','Hypertrophy','Endurance','HIIT','Cardio','Stretching','Yoga','Circuit','Powerlifting','CrossFit'];
const programTmpls  = ['Push Day','Pull Day','Legs Day','Upper Body','Lower Body','Full Body','Bro Split','Cardio Day','Active Recovery','Rest Day'];
const cardioTypes   = ['Running','Cycling','Swimming','Rowing','Elliptical','Jump Rope'];
const sleepQuality  = ['Excellent','Good','Fair','Poor'];
const sorenessLvls  = ['None','Mild','Moderate','Severe'];

(async () => {
  // ── Write headers ────────────────────────────────────────────────────────────
  const headerRow = [['Muscle Group','Workout Type','Program Template','Cardio Type','Sleep Quality','Soreness Level']];

  // ── Build data columns ───────────────────────────────────────────────────────
  const maxRows = Math.max(muscleGroups.length, workoutTypes.length, programTmpls.length, cardioTypes.length, sleepQuality.length, sorenessLvls.length);
  const dataRows = [];
  for (let i = 0; i < maxRows; i++) {
    dataRows.push([
      muscleGroups[i]  || '',
      workoutTypes[i]  || '',
      programTmpls[i]  || '',
      cardioTypes[i]   || '',
      sleepQuality[i]  || '',
      sorenessLvls[i]  || '',
    ]);
  }

  await valuesBatchUpdate(id, [
    { range: "'📋 Reference Data'!A1", values: headerRow },
    { range: "'📋 Reference Data'!A2", values: dataRows },
  ], 'reference-data');

  // ── Format header row ────────────────────────────────────────────────────────
  const reqs = [
    {
      repeatCell: {
        range: gridRange(REF, 0, 1, 0, 6),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.steelGray),
            textFormat: { foregroundColor: hex(C.white), bold: true },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    },
    {
      updateDimensionProperties: {
        range: { sheetId: REF, dimension: 'COLUMNS', startIndex: 0, endIndex: 6 },
        properties: { pixelSize: 140 },
        fields: 'pixelSize',
      },
    },
  ];
  await batchUpdate(id, reqs, 'reference-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
