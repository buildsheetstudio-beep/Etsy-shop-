'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DNA = sheetMap['🧬 DNA & Ethnicity Estimate'];

// Columns: A=Tested Person, B=Region/Ethnicity, C=Percentage, D=Source, E=Notes
const HEADERS = ['Tested Person','Region / Ethnicity','Percentage','Source','Notes'];

// Sarah (user) = Irish-American + Polish + German + Japanese (via marriage/family research, not actual DNA)
// 10 rows summing to 100% for Sarah; all tested as "Sarah Sullivan"
// Regions pulled from Reference!C: Western Europe, Eastern Europe, British Isles, Scandinavia,
//   Southern Europe, Sub-Saharan Africa, North Africa/Middle East, East Asia, South Asia,
//   Southeast Asia, Native American/Indigenous, Pacific Islander, Other
const ROWS = [
  ['Sarah Sullivan','British Isles',       0.32,'AncestryDNA',   'Likely from Sullivan/Flanagan/Murphy Irish lines.'],
  ['Sarah Sullivan','Western Europe',       0.22,'AncestryDNA',   'Probably German — Fischer family line via mother.'],
  ['Sarah Sullivan','Eastern Europe',       0.18,'AncestryDNA',   'Consistent with Kowalski Polish ancestry.'],
  ['Sarah Sullivan','Scandinavia',          0.06,'AncestryDNA',   'Unexpected result — possibly Norman ancestry.'],
  ['Sarah Sullivan','East Asia',            0.12,'23andMe',       'Japanese ancestry via Nakamura maternal line. Higher than expected; may reflect grandmother Yuki\'s contribution.'],
  ['Sarah Sullivan','Southern Europe',      0.04,'23andMe',       'Italy/Greece trace. Possibly via Ramirez family or earlier migrations.'],
  ['Sarah Sullivan','Native American/Indigenous',0.02,'23andMe', 'Trace amount — may be within margin of error; further research needed.'],
  ['Sarah Sullivan','British Isles',        0.00,'Family Lore/Estimate','Duplicate suppressed — see row 1.'],
  ['Sarah Sullivan','Other',                0.02,'Family Lore/Estimate','Unassigned/unresolved in both test results.'],
  ['Sarah Sullivan','Sub-Saharan Africa',   0.02,'MyHeritage',   'Trace result. Awaiting additional test to confirm.'],
];

(async () => {
  const data = [{ range: "'🧬 DNA & Ethnicity Estimate'!A1:E1", values: [HEADERS] }];
  data.push({ range: "'🧬 DNA & Ethnicity Estimate'!A2:E11", values: ROWS });
  await valuesBatchUpdate(id, data, 'dna-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(DNA, 0, 1, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 10; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(DNA, r, r+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  // Percentage col C — centered
  reqs.push({ repeatCell: { range: gridRange(DNA, 1, 11, 2, 3), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat(horizontalAlignment,numberFormat)' } });
  // Notes col — wrap
  reqs.push({ repeatCell: { range: gridRange(DNA, 1, 11, 4, 5), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,140],[1,180],[2,90],[3,160],[4,320]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DNA, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: DNA, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DNA, dimension: 'ROWS', startIndex: 1, endIndex: 11 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'dna-format');
  console.log('DNA & Ethnicity Estimate complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
