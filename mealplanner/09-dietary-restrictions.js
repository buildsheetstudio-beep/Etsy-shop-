'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DRT = sheetMap['🚫 Dietary Restrictions'];

// A=FamilyMember B=Restriction C=Severity/Notes D=SafeSubstitutions
const rows = [
  ['Alex','Gluten-Free','Celiac disease — strict avoidance required','Rice flour, GF pasta, Tamari, Corn tortillas'],
  ['Sam','Dairy-Free','Lactose intolerant — moderate sensitivity','Almond milk, Coconut yogurt, Oat milk, Vegan cheese'],
  ['Jamie','Nut-Free','Tree nut & peanut allergy — anaphylactic (EpiPen)','Sunflower butter, Seed oils, Coconut milk'],
  ['Alex','Egg-Free','Egg allergy — mild contact reaction','Flax eggs (1T flax + 3T water), Chia eggs, Applesauce'],
  ['Sam','Vegetarian','Ethical choice — no meat or fish','Tofu, Tempeh, Lentils, Jackfruit, Black beans'],
  ['Jamie','Soy-Free','Soy sensitivity — moderate digestive issues','Coconut aminos, Sunflower seed butter, Chickpea pasta'],
];

const headers = ['Family Member','Restriction','Severity / Notes','Safe Substitutions'];

(async () => {
  const data = [
    { range: "'🚫 Dietary Restrictions'!A1", values: [['🚫 Dietary Restrictions']] },
    { range: "'🚫 Dietary Restrictions'!A2", values: [headers] },
    { range: "'🚫 Dietary Restrictions'!A3", values: rows },
  ];
  await valuesBatchUpdate(id, data, 'drt-values');

  const reqs = [];

  // Merge A1:D1
  reqs.push({ mergeCells: { range: gridRange(DRT, 0, 1, 0, 4), mergeType: 'MERGE_ALL' } });

  // Title
  reqs.push({ repeatCell: {
    range: gridRange(DRT, 0, 1, 0, 4),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBerry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header
  reqs.push({ repeatCell: {
    range: gridRange(DRT, 1, 2, 0, 4),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Alt rows
  for (let r = 0; r < 6; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DRT, 2 + r, 3 + r, 0, 4),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Center col A and B
  [0, 1].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(DRT, 2, 8, c, c + 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));
  // Wrap text for cols C and D
  [2, 3].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(DRT, 2, 8, c, c + 1),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat.wrapStrategy',
  }}));

  const widths = [120, 110, 250, 250];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: DRT, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DRT, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});
  // Data rows taller for wrapped text
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DRT, dimension: 'ROWS', startIndex: 2, endIndex: 8 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'drt-format');
  console.log('Dietary Restrictions complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
