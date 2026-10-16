'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📋 Reference Data'!A1", values: [
      ['Source Type'],
      ['Birth Certificate'], ['Death Certificate'], ['Census Record'], ['Marriage Record'],
      ['Immigration Record'], ['Interview/Oral History'], ['Photo'], ['Newspaper'],
      ['Military Record'], ['Other'],
    ]},
    { range: "'📋 Reference Data'!B1", values: [
      ['Confidence Level'],
      ['Confirmed'], ['Probable'], ['Uncertain'],
    ]},
    { range: "'📋 Reference Data'!C1", values: [
      ['Region/Ethnicity'],
      ['Western Europe'], ['Eastern Europe'], ['British Isles'], ['Scandinavia'],
      ['Southern Europe'], ['Sub-Saharan Africa'], ['North Africa/Middle East'],
      ['East Asia'], ['South Asia'], ['Southeast Asia'],
      ['Native American/Indigenous'], ['Pacific Islander'], ['Other'],
    ]},
    { range: "'📋 Reference Data'!D1", values: [
      ['Migration Type'],
      ['Immigration'], ['Emigration'], ['Internal Migration'],
    ]},
    { range: "'📋 Reference Data'!E1", values: [
      ['Relationship to User'],
      ['Parent'], ['Sibling'], ['Grandparent'], ['Aunt/Uncle'], ['Cousin'], ['Child'], ['Other'],
    ]},
  ], 'ref-values');

  const reqs = [];
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 14, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.parchment), textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  [[0,160],[1,120],[2,180],[3,140],[4,160]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
