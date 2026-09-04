'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Tags & Categories'];
const S   = "'Recipe Tags & Categories'";
const RB  = "'Recipe Book'";
const REF = "'Reference Data'";

const NC = 16;

(async () => {
  const fmt = [];
  const vals = [];

  // Title
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.peach), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['🏷️ Recipe Tags & Categories']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Browse, filter, and explore your recipe library — all data flows live from Recipe Book']] });

  // Controls
  const controls = [
    ['Meal Type',        'All'],
    ['Recipe Category',  'All'],
    ['Dietary Tag',      'All'],
    ['Allergen Tag',     'All'],
    ['Favorite Only?',   false],
    ['Kid Friendly?',    false],
    ['Meal Prep Friendly?', false],
    ['Max Prep Time (min)', ''],
  ];
  controls.forEach(([label, value], i) => {
    const r = i + 2;
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 3), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 3, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    if (typeof value === 'boolean') {
      fmt.push({ setDataValidation: { range: gridRange(SID, r, r+1, 3, 4), rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } });
      vals.push({ range: `${S}!D${r+1}`, values: [[value]] });
    } else {
      vals.push({ range: `${S}!D${r+1}`, values: [[value]] });
    }
  });

  // Summary Cards (row 10-11, 0-indexed)
  const cardDefs = [
    ['Active Recipes',       `=IFERROR(COUNTIF(${RB}!$N$6:$N$1005,"Active")+COUNTIF(${RB}!$N$6:$N$1005,"Favorite"),0)`],
    ['Favorite Recipes',     `=IFERROR(COUNTIF(${RB}!$M$6:$M$1005,TRUE),0)`],
    ['Quick Meals (≤30 min)',`=IFERROR(COUNTIF(${RB}!$I$6:$I$1005,"<=30")-COUNTIF(${RB}!$B$6:$B$1005,""),0)`],
    ['Kid-Friendly',         `=IFERROR(COUNTIF(${RB}!$Q$6:$Q$1005,TRUE),0)`],
    ['Meal-Prep Friendly',   `=IFERROR(COUNTIF(${RB}!$P$6:$P$1005,TRUE),0)`],
    ['Freezer-Friendly',     `=IFERROR(COUNTIF(${RB}!$O$6:$O$1005,TRUE),0)`],
    ['Vegetarian Recipes',   `=IFERROR(COUNTIFS(${RB}!$R$6:$R$1005,"Vegetarian"),0)+IFERROR(COUNTIFS(${RB}!$R$6:$R$1005,"Vegan"),0)`],
    ['Avg Cost per Serving', `=IFERROR(AVERAGEIF(${RB}!$L$6:$L$1005,">"&0),0)`],
  ];

  cardDefs.forEach(([label, formula], i) => {
    const col = (i % 4) * 4;
    const row = 10 + Math.floor(i / 4) * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.peach), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+2, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Currency for avg cost card
  fmt.push({ repeatCell: { range: gridRange(SID, 12, 13, 12, 16), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Separator
  fmt.push({ mergeCells: { range: gridRange(SID, 14, 15, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 14, 15, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.peach) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Recipe Browser Table header (row 15, 0-indexed)
  const TBL_HDR = 15;
  const TBL_D0  = 16;
  const browserCols = ['Recipe ID','Recipe Name','Category','Meal Type','Base\nServings',
    'Total Time\n(min)','Difficulty','Cost/\nServing','Dietary\nTag 1','Dietary\nTag 2',
    'Dietary\nTag 3','Allergen\nTag 1','Favorite?','Kid\nFriendly?','Meal Prep\nFriendly?','Status'];

  fmt.push({
    repeatCell: {
      range: gridRange(SID, TBL_HDR, TBL_HDR+1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A${TBL_HDR+1}`, values: [browserCols] });

  // Browser data — live formulas pulling from Recipe Book rows 6:55 (50 recipes)
  const browserRows = [];
  for (let i = 0; i < 50; i++) {
    const r1 = TBL_D0 + i;
    const srcRow = i + 6;
    browserRows.push([
      `=IFERROR(${RB}!A${srcRow},"")`,   // A Recipe ID
      `=IFERROR(${RB}!B${srcRow},"")`,   // B Name
      `=IFERROR(${RB}!C${srcRow},"")`,   // C Category
      `=IFERROR(${RB}!D${srcRow},"")`,   // D Meal Type
      `=IFERROR(${RB}!F${srcRow},"")`,   // E Base servings
      `=IFERROR(${RB}!I${srcRow},"")`,   // F Total time
      `=IFERROR(${RB}!J${srcRow},"")`,   // G Difficulty
      `=IFERROR(${RB}!L${srcRow},"")`,   // H Cost/serving
      `=IFERROR(${RB}!R${srcRow},"")`,   // I Dietary 1
      `=IFERROR(${RB}!S${srcRow},"")`,   // J Dietary 2
      `=IFERROR(${RB}!T${srcRow},"")`,   // K Dietary 3
      `=IFERROR(${RB}!U${srcRow},"")`,   // L Allergen 1
      `=IFERROR(${RB}!M${srcRow},FALSE)`,// M Favorite
      `=IFERROR(${RB}!Q${srcRow},FALSE)`,// N Kid Friendly
      `=IFERROR(${RB}!P${srcRow},FALSE)`,// O Meal Prep
      `=IFERROR(${RB}!N${srcRow},"")`,   // P Status
    ]);

    const bgColor = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, TBL_D0+i, TBL_D0+i+1, 0, NC),
        cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    // Formula cells: all of these
    for (let c = 0; c < NC; c++) {
      fmt.push({ repeatCell: { range: gridRange(SID, TBL_D0+i, TBL_D0+i+1, c, c+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }

  vals.push({ range: `${S}!A${TBL_D0+1}:P${TBL_D0+50}`, values: browserRows });

  // Currency for H (col 7)
  fmt.push({ repeatCell: { range: gridRange(SID, TBL_D0, TBL_D0+50, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths
  const colWidths = [80,170,110,100,60,70,90,70,110,110,110,110,70,70,80,100];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: TBL_HDR, endIndex: TBL_HDR+1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Freeze rows 1:6
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'tags-fmt');
  await valuesBatchUpdate(id, vals, 'tags-vals');
  console.log('✓ Recipe Tags & Categories complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
