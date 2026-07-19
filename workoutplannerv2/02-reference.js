'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📋 Reference Data'!A1", values: [
      ['Muscle Group / Focus'],
      ['Chest'], ['Back'], ['Shoulders'], ['Arms'], ['Legs'], ['Glutes'], ['Core'], ['Cardio'], ['Rest'],
    ]},
    { range: "'📋 Reference Data'!B1", values: [
      ['Workout Type'],
      ['Strength'], ['Cardio'], ['HIIT'], ['Circuit/Superset'], ['Mobility/Stretching'], ['Rest'],
    ]},
    { range: "'📋 Reference Data'!C1", values: [
      ['Cardio Activity'],
      ['Running'], ['Cycling'], ['Swimming'], ['Rowing'], ['Walking'], ['Other'],
    ]},
    { range: "'📋 Reference Data'!D1", values: [
      ['PR Type'],
      ['1-Rep Max'], ['Max Reps'], ['Best Time'], ['Longest Distance'],
    ]},
    { range: "'📋 Reference Data'!E1", values: [
      ['Heart Rate Zone'],
      ['Zone 1: Warm-up (50-60% Max HR)'],
      ['Zone 2: Fat Burn (60-70%)'],
      ['Zone 3: Cardio (70-80%)'],
      ['Zone 4: Hard (80-90%)'],
      ['Zone 5: Max Effort (90-100%)'],
    ]},
  ], 'ref-values');

  const reqs = [];
  // Header row
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  // Data rows
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 12, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark), fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  [[0,140],[1,140],[2,120],[3,130],[4,240]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
