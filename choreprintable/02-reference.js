'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['Reference Data'];
const R = "'Reference Data'";

(async () => {
  const data = [];

  // A. Family Member Slots A1:C7
  data.push({ range: `${R}!A1:C1`, values: [['Family Member','Default Name','Assigned Color']] });
  data.push({ range: `${R}!A2:C7`, values: [
    ['Member 1','Mom',     '#F4C7AB'],
    ['Member 2','Dad',     '#E8B7C5'],
    ['Member 3','Child 1', '#BFD7EA'],
    ['Member 4','Child 2', '#D7C6E8'],
    ['Member 5','Child 3', '#C7E5D1'],
    ['Member 6','Child 4', '#F3E4A6'],
  ]});

  // B. Chore Library E1:F31
  data.push({ range: `${R}!E1:F1`, values: [['Category','Chore']] });
  data.push({ range: `${R}!E2:F31`, values: [
    ['Bedroom',   'Make bed'],
    ['Bedroom',   'Put clothes away'],
    ['Bedroom',   'Organize bedroom'],
    ['Bedroom',   'Change bed sheets'],
    ['Kitchen',   'Load dishwasher'],
    ['Kitchen',   'Unload dishwasher'],
    ['Kitchen',   'Wipe kitchen counters'],
    ['Kitchen',   'Sweep kitchen floor'],
    ['Kitchen',   'Help prepare dinner'],
    ['Kitchen',   'Pack lunch'],
    ['Bathroom',  'Clean bathroom sink'],
    ['Bathroom',  'Wipe bathroom mirror'],
    ['Bathroom',  'Replace towels'],
    ['Bathroom',  'Clean toilet'],
    ['Laundry',   'Sort laundry'],
    ['Laundry',   'Fold laundry'],
    ['Laundry',   'Put laundry away'],
    ['Pets',      'Feed pet'],
    ['Pets',      'Refill water bowl'],
    ['Pets',      'Walk dog'],
    ['Pets',      'Clean litter box'],
    ['School',    'Complete homework'],
    ['School',    'Pack school bag'],
    ['School',    'Read for 20 minutes'],
    ['Household', 'Take out trash'],
    ['Household', 'Water plants'],
    ['Household', 'Dust furniture'],
    ['Household', 'Vacuum'],
    ['Outdoor',   'Bring in mail'],
    ['Outdoor',   'Help with yard work'],
  ]});

  // C. Status/Settings H1:H3, J1:J7
  data.push({ range: `${R}!H1:H3`, values: [['Yes / No'],['Yes'],['No']] });
  data.push({ range: `${R}!J1:J7`, values: [['Family Size'],['1'],['2'],['3'],['4'],['5'],['6']] });

  // D. Chart Labels L1:L8
  data.push({ range: `${R}!L1:L8`, values: [
    ['Summary Labels'],['Total Chores'],['Completed'],['Remaining'],
    ['Completion %'],['Week Start'],['Family Members'],['Notes']
  ]});

  await valuesBatchUpdate(id, data, 'ref-values');

  // Format Reference Data
  const reqs = [];
  const bold10 = { textFormat: { bold: true, fontSize: 10 }, backgroundColor: hex(C.headerGray) };
  reqs.push({ repeatCell: { range: gridRange(REF, 0, 1, 0, 13), cell: { userEnteredFormat: bold10 }, fields: 'userEnteredFormat' } });
  // Color swatches for C2:C7
  for (let i = 0; i < 6; i++) {
    reqs.push({ repeatCell: {
      range: gridRange(REF, i+1, i+2, 2, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.member[i]) } },
      fields: 'userEnteredFormat.backgroundColor'
    }});
  }
  // Col widths
  [[0,90],[1,90],[2,100],[4,90],[5,160],[7,60],[9,70],[11,110]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });
  // Freeze row 1
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, gridProperties: { frozenRowCount: 1, frozenColumnCount: 0 } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
  }});

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
