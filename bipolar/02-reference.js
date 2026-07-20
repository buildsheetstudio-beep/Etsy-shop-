'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📋 Reference Data'!A1", values: [
      ['Episode Type'],
      ['Manic'], ['Hypomanic'], ['Depressive'], ['Stable/Euthymic'],
    ]},
    { range: "'📋 Reference Data'!B1", values: [
      ['Severity Scale'],
      ['Minimal (0-20%)'], ['Mild (20-40%)'], ['Moderate (40-60%)'], ['High (60-80%)'], ['Severe (80-100%)'],
    ]},
    { range: "'📋 Reference Data'!C1", values: [
      ['Medication Frequency'],
      ['Once Daily'], ['Twice Daily'], ['As Needed'], ['Weekly'], ['Other'],
    ]},
    { range: "'📋 Reference Data'!D1", values: [
      ['Sleep Quality'],
      ['Poor'], ['Fair'], ['Good'], ['Great'],
    ]},
    { range: "'📋 Reference Data'!E1", values: [
      ['Trigger Category'],
      ['Work/School'], ['Relationship'], ['Financial'], ['Health'],
      ['Sleep Disruption'], ['Substance Use'], ['Other'],
    ]},
    { range: "'📋 Reference Data'!F1", values: [
      ['Impact Severity'],
      ['Low'], ['Medium'], ['High'],
    ]},
    { range: "'📋 Reference Data'!G1", values: [
      ['Provider Type'],
      ['Psychiatrist'], ['Therapist'], ['Primary Care'], ['Other'],
    ]},
    { range: "'📋 Reference Data'!H1", values: [
      ['Manic Symptoms (reference)'],
      ['Elevated or irritable mood'],
      ['Decreased need for sleep'],
      ['Racing thoughts'],
      ['Rapid or pressured speech'],
      ['Increased goal-directed activity'],
      ['Impulsive or risky decisions'],
      ['Grandiosity or inflated self-esteem'],
      ['Distractibility'],
    ]},
    { range: "'📋 Reference Data'!I1", values: [
      ['Depressive Symptoms (reference)'],
      ['Persistent low mood'],
      ['Loss of interest or pleasure'],
      ['Fatigue or low energy'],
      ['Difficulty concentrating'],
      ['Changes in appetite or sleep'],
      ['Feelings of worthlessness'],
      ['Social withdrawal'],
      ['Slowed thinking or movement'],
    ]},
  ], 'ref-values');

  const reqs = [];
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 9),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 12, 0, 9),
      cell: { userEnteredFormat: { backgroundColor: hex(C.fog), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  [[0,120],[1,140],[2,130],[3,100],[4,130],[5,90],[6,110],[7,220],[8,220]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
