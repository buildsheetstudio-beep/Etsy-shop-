'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TP = sheetMap['Table Planner'];
const S  = "'Table Planner'";

const HEADERS = [
  'Table #', 'Table Name', 'Shape', 'Capacity', 'Zone / Location',
  'Assigned (formula)', 'Open Seats (formula)', '% Full (formula)',
  'Reserved For', 'Accessibility', 'AV / Tech Needs', 'Status', 'Meal Notes', 'Notes',
];

// 14 tables for ~120 guests (mix of 8-10 per table)
const tables = [
  [1,  'Head Table',     'Rectangle', 10, 'Head Table',       'Wedding Party',  true,  false, 'Confirmed', 'All entrées available', ''],
  [2,  'Oak',            'Round',      8, 'Family Front',     'Osei Family',    false, false, 'Confirmed', '', ''],
  [3,  'Magnolia',       'Round',      8, 'Family Front',     'Thornton Family',false, false, 'Confirmed', '', ''],
  [4,  'Willow',         'Round',      8, 'Center',           'General',        false, false, 'Confirmed', '', ''],
  [5,  'Cedar',          'Round',      8, 'Center',           'General',        false, false, 'Confirmed', '', ''],
  [6,  'Birch',          'Round',      8, 'Center',           'General',        false, false, 'Confirmed', '', ''],
  [7,  'Cypress',        'Round',      8, 'Window',           'General',        false, true,  'Confirmed', '', 'Garden view — romantic corner'],
  [8,  'Rosewood',       'Round',      8, 'Window',           'General',        false, false, 'Confirmed', '', ''],
  [9,  'Hazel',          'Round',      8, 'Garden',           'General',        false, false, 'Confirmed', '', ''],
  [10, 'Jasmine',        'Round',      8, 'Garden',           'General',        false, false, 'Confirmed', '', ''],
  [11, 'Lavender',       'Round',      8, 'Near Dance Floor', 'General',        false, false, 'Confirmed', '', ''],
  [12, 'Sage',           'Round',      8, 'Near Dance Floor', 'General',        false, false, 'Confirmed', '', ''],
  [13, 'Ivory',          'Round',      6, 'Corner',           'Children\'s',    false, false, 'Confirmed', 'Kids meals only', 'Kids table — near parents'],
  [14, 'Pearl',          'Round',      8, 'Corner',           'General',        false, true,  'Confirmed', '', 'Accessible seating — wide aisle'],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['TABLE PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(TP, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(TP, 0, 1, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: TP, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Summary row 2 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [[
    'Total Tables:','',
    `=IFERROR(COUNTA(A5:A100),"")`, '',
    'Total Seats:','',
    `=IFERROR(SUM(D5:D100),"")`, '',
    'Seats Filled:','',
    `=IFERROR(SUM(F5:F100),"")`, '',
    'Seats Available:','',
    `=IFERROR(SUM(G5:G100),"")`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(TP, 1, 2, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Headers row 4 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A4`, values: [HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(TP, 3, 4, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: TP, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ── Table Data ─────────────────────────────────────────────────────────────
  tables.forEach((t, i) => {
    const row = 5 + i;
    const [num, name, shape, cap, zone, reservedFor, avNeeds, a11y, status, mealNotes, notes] = t;
    data.push({ range: `${S}!A${row}`, values: [[num]] });
    data.push({ range: `${S}!B${row}`, values: [[name]] });
    data.push({ range: `${S}!C${row}`, values: [[shape]] });
    data.push({ range: `${S}!D${row}`, values: [[cap]] });
    data.push({ range: `${S}!E${row}`, values: [[zone]] });
    // F: Assigned count — COUNTIF against Seating Chart col H
    data.push({ range: `${S}!F${row}`, values: [[`=IFERROR(SUMPRODUCT(('Seating Chart'!H$6:H$600=A${row})*('Seating Chart'!G$6:G$600="Yes")),0)`]] });
    // G: Open seats
    data.push({ range: `${S}!G${row}`, values: [[`=IFERROR(D${row}-F${row},0)`]] });
    // H: % Full
    data.push({ range: `${S}!H${row}`, values: [[`=IFERROR(F${row}/D${row},0)`]] });
    data.push({ range: `${S}!I${row}`, values: [[reservedFor]] });
    data.push({ range: `${S}!J${row}`, values: [[a11y]] });    // Accessibility checkbox
    data.push({ range: `${S}!K${row}`, values: [[avNeeds]] }); // AV needs checkbox
    data.push({ range: `${S}!L${row}`, values: [[status]] });
    data.push({ range: `${S}!M${row}`, values: [[mealNotes]] });
    data.push({ range: `${S}!N${row}`, values: [[notes]] });
  });

  // Checkboxes: J (Accessibility), K (AV Needs)
  [9, 10].forEach(col => {
    reqs.push({ setDataValidation: {
      range: gridRange(TP, 4, 4 + tables.length, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Status dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(TP, 4, 100, 11, 12),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Confirmed', 'Draft', 'Needs Review', 'TBD',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Shape dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(TP, 4, 100, 2, 3),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Round', 'Rectangle', 'Square', 'Oval', 'Banquet',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Zone dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(TP, 4, 100, 4, 5),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Head Table', 'Family Front', 'Center', 'Window', 'Garden', 'Balcony', 'Corner', 'Near Dance Floor', 'Near Bar', 'Other',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Format col F/G as number, H as percent
  reqs.push({ repeatCell: { range: gridRange(TP, 4, 4 + tables.length, 5, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
    numberFormat: { type: 'NUMBER', pattern: '#,##0' },
  }}, fields: 'userEnteredFormat(backgroundColor,numberFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(TP, 4, 4 + tables.length, 7, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
    numberFormat: { type: 'PERCENT', pattern: '0%' },
  }}, fields: 'userEnteredFormat(backgroundColor,numberFormat)' }});

  // Alternating row colors
  for (let i = 0; i < tables.length; i++) {
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(TP, 4 + i, 5 + i, 0, 14), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, foregroundColor: hex(C.mainText) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // CF: % Full >= 100% → red; >= 80% → amber
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(TP, 4, 4 + tables.length, 7, 8)],
      booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] },
        format: { backgroundColor: hex(C.attention), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
    }, index: 0,
  }});
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(TP, 4, 4 + tables.length, 7, 8)],
      booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.8' }] },
        format: { backgroundColor: hex(C.warning) } },
    }, index: 0,
  }});

  // Column widths
  const widths = [60, 90, 80, 65, 120, 90, 80, 70, 120, 80, 80, 90, 130, 160];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: TP, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze row 4
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TP, gridProperties: { frozenRowCount: 4 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TP, tabColor: hex(C.warning) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'tables-data');
  await batchUpdate(id, reqs, 'tables-format');

  console.log('✓ Table Planner built with', tables.length, 'tables');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
