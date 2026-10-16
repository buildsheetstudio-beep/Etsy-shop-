'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Household Setup'];
const S   = "'Household Setup'";
const REF = "'Reference Data'";

// Named ranges for Reference Data columns
const refCol = (col) => `${REF}!${col}2:${col}50`;
const NC = 14; // number of columns used

(async () => {
  const fmt = [];
  const vals = [];

  // ─── Title banner ─────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.peach),
          textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['🏠 Household Setup']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.butter),
          textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Configure your household, members, preferences, and meal planning defaults']] });

  // Freeze rows 1:3
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  // ─── Section A: Household Settings (rows 4-17, 0-indexed 3-16) ────────────
  const sectionHdr = (row, label, color) => {
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, NC), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, NC),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(color),
            textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) },
            horizontalAlignment: 'LEFT',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
    vals.push({ range: `${S}!A${row+1}`, values: [[`  ${label}`]] });
  };

  sectionHdr(2, 'HOUSEHOLD SETTINGS', C.text);

  // Label/Value rows
  const settingRows = [
    ['Household Name',                'The Rivera-Nguyen Family',     false],
    ['Primary Planner',               'Sofia Rivera',                  false],
    ['Primary Currency',              'USD ($)',                        false],
    ['Default Household Serving Size','4',                             false],
    ['Weekly Grocery Budget',         '$200.00',                       false],
    ['Preferred Week Start Day',      'Monday',                        false],
    ['Grocery Shopping Day',          'Saturday',                      false],
    ['Default Store',                 'Green Valley Market',           false],
    ['Include Snacks in Weekly Plan?',true,                            true ],
    ['Include Desserts in Weekly Plan?',true,                          true ],
    ['Use Pantry Inventory in Grocery Calculations?',true,             true ],
    ['Track Meal Costs?',             true,                            true ],
    ['Notes',                         'Allergic to tree nuts — always check labels.', false],
  ];

  settingRows.forEach(([label, value, isCheckbox], i) => {
    const r = i + 3; // 0-indexed, row 4 = index 3
    // Label cell
    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 0, 5),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.altRow),
            textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
            horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    // Value cell
    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 5, 10),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.input),
            textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    if (isCheckbox) {
      fmt.push({
        setDataValidation: {
          range: gridRange(SID, r, r+1, 5, 6),
          rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
        },
      });
      vals.push({ range: `${S}!F${r+1}`, values: [[value]] });
    } else {
      vals.push({ range: `${S}!F${r+1}`, values: [[value]] });
    }
  });

  // ─── Serving Controls section (rows 17-24, 0-indexed 16-23) ──────────────
  const SC_ROW = 16;
  sectionHdr(SC_ROW, 'SERVING CONTROLS', '#547B82');

  const servingControls = [
    ['Default Household Servings',  '4',  false],
    ['Minimum Recipe Servings',     '1',  false],
    ['Maximum Recipe Servings',     '12', false],
    ['Default Leftover Servings',   '2',  false],
    ['Include Guests?',             false, true],
    ['Guest Serving Override',      '0',  false],
  ];
  servingControls.forEach(([label, value, isCheckbox], i) => {
    const r = SC_ROW + 1 + i;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 5, 10),
        cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    if (isCheckbox) {
      fmt.push({ setDataValidation: { range: gridRange(SID, r, r+1, 5, 6), rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } });
      vals.push({ range: `${S}!F${r+1}`, values: [[value]] });
    } else {
      vals.push({ range: `${S}!F${r+1}`, values: [[value]] });
    }
  });

  // ─── Dashboard Controls (rows 23-30, 0-indexed 22-29) ────────────────────
  const DC_ROW = 23;
  sectionHdr(DC_ROW, 'DASHBOARD CONTROLS', '#547B82');

  const dashControls = [
    ['Selected Week Start Date',          '=DATE(2026,8,17)', false],
    ['Selected Household Member / All',   'All',              false],
    ['Selected Meal Type / All',          'All',              false],
    ['Selected Recipe Category / All',    'All',              false],
    ['Grocery Category / All',            'All',              false],
    ['Budget View',                       'Weekly',           false],
  ];
  dashControls.forEach(([label, value, isCheckbox], i) => {
    const r = DC_ROW + 1 + i;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 5, 10),
        cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' }, userEnteredValue: { stringValue: '' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    vals.push({ range: `${S}!F${r+1}`, values: [[value]] });
  });

  // Dropdown for dashboard member
  fmt.push({
    setDataValidation: {
      range: gridRange(SID, DC_ROW+2, DC_ROW+3, 5, 6),
      rule: { condition: { type: 'ONE_OF_LIST', values: [
        { userEnteredValue: 'All' },
        { userEnteredValue: 'Sofia Rivera' },
        { userEnteredValue: 'Marco Rivera' },
        { userEnteredValue: 'Lily Rivera' },
        { userEnteredValue: 'Jake Rivera' },
        { userEnteredValue: 'Grandma Rosa' },
        { userEnteredValue: 'Uncle Ben' },
      ]}, showCustomUi: true, strict: false },
    },
  });

  // ─── Summary Cards (row 31-36, 0-indexed 30-35) ──────────────────────────
  const CARD_ROW = 30;
  sectionHdr(CARD_ROW, 'HOUSEHOLD SUMMARY', C.peach.replace('#','') === 'F3C9B6' ? C.text : C.text);

  const cards = [
    ['Active Household Members',    `=IFERROR(COUNTIF('Household Setup'!L16:L27,TRUE),0)`],
    ['Default Household Servings',  `=IFERROR('Household Setup'!F20,0)`],
    ['Weekly Grocery Budget',       `=IFERROR('Household Setup'!F8,0)`],
    ['Recipes Available',           `=IFERROR(COUNTA('Recipe Book'!B6:B1005),0)`],
    ['Pantry Items',                `=IFERROR(COUNTA('Pantry Inventory'!B6:B1505),0)`],
    ['Dietary Restrictions Tracked',`=IFERROR(COUNTA('Household Setup'!J16:J27),0)`],
    ['Allergens Tracked',           `=IFERROR(COUNTA('Household Setup'!K16:K27),0)`],
    ['Favorite Recipes',            `=IFERROR(COUNTIF('Recipe Book'!M6:M1005,TRUE),0)`],
  ];

  cards.forEach(([label, formula], i) => {
    const col = (i % 4) * 3;
    const row = CARD_ROW + 1 + Math.floor(i / 4) * 3;
    // Card merge
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+3), mergeType: 'MERGE_ALL' } });
    // Card label
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    // Card value
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+2, col, col+3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // ─── Household Member Table (rows 36-49, 0-indexed 35-48) ────────────────
  const MEM_TITLE_ROW = 38;
  const MEM_HDR_ROW = 39;
  const MEM_D0 = 40;
  const NM = 12;

  sectionHdr(MEM_TITLE_ROW - 1, 'HOUSEHOLD MEMBERS', '#547B82');

  const memCols = ['Member ID','Member Name','Household Role','Default Serving\nMultiplier',
    'Dietary\nRestriction 1','Dietary\nRestriction 2','Dietary\nRestriction 3',
    'Allergen 1','Allergen 2','Allergen 3','Active?','Notes'];

  // Column header row
  fmt.push({
    repeatCell: {
      range: gridRange(SID, MEM_HDR_ROW, MEM_HDR_ROW+1, 0, memCols.length),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.text),
          textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A${MEM_HDR_ROW+1}`, values: [memCols] });

  // Data rows alternating
  for (let r = 0; r < NM; r++) {
    const bgColor = r % 2 === 0 ? C.panel : C.altRow;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, MEM_D0+r, MEM_D0+r+1, 0, memCols.length),
        cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    // Formula cells: ID (col A)
    fmt.push({
      repeatCell: {
        range: gridRange(SID, MEM_D0+r, MEM_D0+r+1, 0, 1),
        cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
    // Active checkbox (col K = index 10)
    fmt.push({
      setDataValidation: {
        range: gridRange(SID, MEM_D0+r, MEM_D0+r+1, 10, 11),
        rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
      },
    });
    // Dropdown: role (col C=2), dietary (D,E,F = 3,4,5,  but offset from 0), allergens (G,H,I)
    [[2, 'G'], [4, 'I'], [5, 'I'], [6, 'I'], [7, 'H'], [8, 'H'], [9, 'H']].forEach(([colIdx, refColLetter]) => {
      const refMap = {2: 'G', 4: 'I', 5: 'I', 6: 'I', 7: 'J', 8: 'J', 9: 'J'};
      // use ONE_OF_RANGE for dropdowns
    });
    // Dropdowns with ONE_OF_RANGE
    const dropMap = [
      [2, 'G'],  // Household Role
      [4, 'I'],  // Dietary 1
      [5, 'I'],  // Dietary 2
      [6, 'I'],  // Dietary 3
      [7, 'J'],  // Allergen 1
      [8, 'J'],  // Allergen 2
      [9, 'J'],  // Allergen 3
    ];
    dropMap.forEach(([colIdx, refLetter]) => {
      fmt.push({
        setDataValidation: {
          range: gridRange(SID, MEM_D0+r, MEM_D0+r+1, colIdx, colIdx+1),
          rule: {
            condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$${refLetter}$2:$${refLetter}$50` }] },
            showCustomUi: true, strict: false,
          },
        },
      });
    });
  }

  // Member ID formula
  vals.push({
    range: `${S}!A${MEM_D0+1}:A${MEM_D0+NM}`,
    values: Array.from({length: NM}, (_, i) => [`=IF(B${MEM_D0+1+i}="","","MEM-"&TEXT(ROW()-${MEM_D0},"00"))`]),
  });

  // Sample members
  const members = [
    ['Sofia Rivera',   'Self',           '1.0',  'No Restriction', '', '', 'Tree Nuts', '', 'None', true,  'Primary planner'],
    ['Marco Rivera',   'Partner',        '1.0',  'No Restriction', '', '', 'Tree Nuts', '', 'None', true,  ''],
    ['Lily Rivera',    'Child',          '0.6',  'No Restriction', '', '', 'None',      '', 'None', true,  'Age 7'],
    ['Jake Rivera',    'Child',          '0.75', 'No Restriction', '', '', 'None',      '', 'None', true,  'Age 11'],
    ['Grandma Rosa',   'Family Member',  '0.9',  'Low Sodium',     '', '', 'None',      '', 'None', true,  'Prefers soft foods'],
    ['Uncle Ben',      'Guest',          '1.0',  'Vegan',          '', '', 'None',      '', 'None', true,  'Visits monthly'],
  ];
  members.forEach((row, i) => {
    vals.push({ range: `${S}!B${MEM_D0+1+i}:L${MEM_D0+1+i}`, values: [row] });
  });

  // ─── Disclaimer (row 53-56) ───────────────────────────────────────────────
  const DISC_ROW = MEM_D0 + NM + 1;
  fmt.push({ mergeCells: { range: gridRange(SID, DISC_ROW, DISC_ROW+4, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, DISC_ROW, DISC_ROW+4, 0, NC),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.attention),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white), bold: false, italic: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'TOP', wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A${DISC_ROW+1}`, values: [['⚠ IMPORTANT DISCLAIMER: Dietary and allergen tags in this workbook are organizational aids only. They are not a substitute for ingredient-label review, medical advice, allergy management, or professional dietary guidance. Always verify ingredient labels, cross-contamination risks, preparation methods, and individual dietary needs before serving food.']] });

  // ─── Column widths ────────────────────────────────────────────────────────
  const colWidths = [90, 140, 120, 80, 110, 110, 110, 110, 110, 110, 60, 180];
  colWidths.forEach((w, i) => {
    fmt.push({
      updateDimensionProperties: {
        range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
        properties: { pixelSize: w }, fields: 'pixelSize',
      },
    });
  });

  // Row heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: MEM_HDR_ROW, endIndex: MEM_HDR_ROW+1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  await batchUpdate(id, fmt, 'household-fmt');
  await valuesBatchUpdate(id, vals, 'household-vals');
  console.log('✓ Household Setup complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
