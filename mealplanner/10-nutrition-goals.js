'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const NGO = sheetMap['🎯 Nutrition Goals'];

// A=FamilyMember B=WeekOf C=CalorieGoal D=ActualCalories E=ProteinGoal(g) F=ActualProtein(g) G=Variance%
// Current week rows (index 3,4,5) use formulas; historical (index 0,1,2) hardcoded
const rows = [
  // Historical week (2026-07-06)
  ['Alex','2026-07-06',2000,1892,150,138],
  ['Sam','2026-07-06',1800,1740,120,118],
  ['Jamie','2026-07-06',2200,2310,160,172],
  // Current week (2026-07-13) - formulas for actual values
  ['Alex','2026-07-13',2000,null,150,null],
  ['Sam','2026-07-13',1800,null,120,null],
  ['Jamie','2026-07-13',2200,null,160,null],
];

const headers = ['Family Member','Week Of','Calorie Goal','Actual Calories','Protein Goal (g)','Actual Protein (g)','Variance %'];

(async () => {
  const dataRows = rows.map((r, i) => {
    const row = i + 3;
    const isCurrent = i >= 3;
    const actualCal = isCurrent
      ? `=IFERROR(SUMIFS('📅 Weekly Meal Plan'!$E:$E,'📅 Weekly Meal Plan'!$C:$C,A${row}),0)`
      : r[3];
    const actualProt = isCurrent
      ? `=IFERROR(SUMPRODUCT(('📅 Weekly Meal Plan'!$C$3:$C$23=A${row})*IFERROR(INDEX('📖 Recipe Book'!$J:$J,MATCH('📅 Weekly Meal Plan'!$D$3:$D$23,'📖 Recipe Book'!$B:$B,0)),0)),0)`
      : r[5];
    return [
      r[0], r[1], r[2], actualCal, r[4], actualProt,
      `=IFERROR((D${row}-C${row})/C${row},0)`,
    ];
  });

  const data = [
    { range: "'🎯 Nutrition Goals'!A1", values: [['🎯 Nutrition Goals']] },
    { range: "'🎯 Nutrition Goals'!A2", values: [headers] },
    { range: "'🎯 Nutrition Goals'!A3", values: dataRows },
  ];
  await valuesBatchUpdate(id, data, 'ngo-values');

  const reqs = [];

  // Merge A1:G1
  reqs.push({ mergeCells: { range: gridRange(NGO, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });

  // Title
  reqs.push({ repeatCell: {
    range: gridRange(NGO, 0, 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBasil), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header
  reqs.push({ repeatCell: {
    range: gridRange(NGO, 1, 2, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Alt rows
  for (let r = 0; r < 6; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(NGO, 2 + r, 3 + r, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Formula cols D, F, G: formulaBg
  [3, 5, 6].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(NGO, 2, 8, c, c + 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
  }}));
  // Percent format col G
  reqs.push({ repeatCell: {
    range: gridRange(NGO, 2, 8, 6, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Date format col B
  reqs.push({ repeatCell: {
    range: gridRange(NGO, 2, 8, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }});
  // Center numeric cols
  [2, 3, 4, 5].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(NGO, 2, 8, c, c + 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  const widths = [120, 110, 100, 120, 110, 130, 90];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: NGO, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: NGO, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'ngo-format');
  console.log('Nutrition Goals complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
