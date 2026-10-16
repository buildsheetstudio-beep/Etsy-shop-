'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SC = sheetMap['Seating Chart'];
const S  = "'Seating Chart'";

const HEADERS = [
  '#', 'First Name', 'Last Name', 'Full Name', 'Household ID',
  'RSVP Status (auto)', 'Attending (auto)', 'Table #', 'Seat #',
  'Table Name (auto)', 'Zone (auto)', 'Meal Preference (auto)',
  'Dietary Restrictions (auto)', 'Accessibility (auto)',
  'Plus One Guest', 'Household Conflict (auto)', 'Over Capacity (auto)',
  'Special Arrangement', 'Seating Notes', 'Confirmed',
];

// 64 guests assigned to 14 tables
// Format: [first, last, householdId, tableNum, seatNum, plusOne, arrangement, notes, confirmed]
const seats = [
  // Table 1 — Head Table (Wedding Party, 10 seats)
  ['Amara',    'Osei',       'H00', 1,  1, false, 'Bride',        '', true],
  ['Marcus',   'Thornton',   'H00', 1,  2, false, 'Groom',        '', true],
  ['Jordan',   'Osei',       'H14', 1,  3, false, 'Maid of Honor','', true],
  ['Noel',     'Parker',     'H18', 1,  4, false, 'Best Man',     '', true],
  ['Daniela',  'Rivera',     'H15', 1,  5, false, 'Bridesmaid',   '', true],
  ['Carlos',   'Rivera',     'H15', 1,  6, true,  'Bridesmaid+1', '', true],
  ['Sophie',   'Chen',       'H16', 1,  7, false, 'Bridesmaid',   '', true],
  ['Tristan',  'Davis',      'H17', 1,  8, false, 'Groomsman',    '', true],
  ['Aria',     'Davis',      'H17', 1,  9, true,  'Groomsman+1',  '', true],
  ['Lena',     'Park',       'H14', 1, 10, true,  'MoH+1',        '', true],
  // Table 2 — Oak (Osei Family, 8 seats)
  ['Kwame',    'Osei',       'H01', 2,  1, false, '', '', true],
  ['Abena',    'Osei',       'H01', 2,  2, false, '', '', true],
  ['Kofi',     'Osei',       'H02', 2,  3, false, '', '', true],
  ['Priya',    'Osei',       'H02', 2,  4, false, '', 'Nut allergy — check entrée', true],
  ['Akosua',   'Mensah',     'H03', 2,  5, false, '', '', true],
  ['Yaw',      'Mensah',     'H03', 2,  6, false, '', '', true],
  ['Kwame Jr', 'Mensah',     'H03', 2,  7, false, 'Child', 'Kids meal', true],
  ['Serwa',    'Acheampong', 'H06', 2,  8, false, '', '', true],
  // Table 3 — Magnolia (Thornton Family, 8 seats)
  ['Robert',   'Thornton',   'H07', 3,  1, false, '', '', true],
  ['Patricia', 'Thornton',   'H07', 3,  2, false, '', '', true],
  ['Jason',    'Thornton',   'H08', 3,  3, false, '', '', true],
  ['Mei',      'Thornton',   'H08', 3,  4, false, '', 'Dairy-free', true],
  ['Oliver',   'Thornton',   'H08', 3,  5, false, 'Child', 'Kids meal', true],
  ['Charles',  'Thornton',   'H10', 3,  6, false, '', '', true],
  ['Helen',    'Thornton',   'H10', 3,  7, false, '', '', true],
  ['Michael',  'Barnes',     'H11', 3,  8, false, '', '', true],
  // Table 4 — Willow (Mixed, 8 seats)
  ['Susan',    'Barnes',     'H11', 4,  1, false, '', '', true],
  ['Kira',     'Thornton',   'H09', 4,  2, false, '', '', true],
  ['Dean',     'Walsh',      'H09', 4,  3, false, '', '', true],
  ['Sandra',   'Ellis',      'H13', 4,  4, false, '', '', true],
  ['Ben',      'Ellis',      'H13', 4,  5, false, '', '', true],
  ['Ama',      'Boateng',    'H04', 4,  6, false, '', 'Wheelchair — end seat', true],
  ['Nana',     'Boateng',    'H04', 4,  7, false, '', '', true],
  ['Yaa',      'Boateng',    'H04', 4,  8, false, 'Child', 'Kids meal', true],
  // Table 5 — Cedar (Bride Friends, 8 seats)
  ['David',    'Kim',        'H19', 5,  1, false, '', '', true],
  ['Julia',    'Kim',        'H19', 5,  2, false, '', 'Nut allergy', true],
  ['Ella',     'Kim',        'H19', 5,  3, false, 'Child', 'Kids meal', true],
  ['Rohan',    'Patel',      'H20', 5,  4, false, '', '', true],
  ['Nisha',    'Patel',      'H20', 5,  5, false, '', 'Dairy-free', true],
  ['Elena',    'Torres',     'H22', 5,  6, false, '', '', true],
  ['Emeka',    'Abara',      'H32', 5,  7, false, '', '', true],
  ['Chioma',   'Abara',      'H32', 5,  8, false, '', '', true],
  // Table 6 — Birch (Groom Friends, 8 seats)
  ['Thanh',    'Nguyen',     'H21', 6,  1, false, '', '', true],
  ['Linh',     'Nguyen',     'H21', 6,  2, false, '', '', true],
  ['Sam',      'Morgan',     'H23', 6,  3, false, '', '', true],
  ['Alex',     'Morgan',     'H23', 6,  4, false, '', '', true],
  ['Marco',    'Reyes',      'H24', 6,  5, false, '', '', true],
  ['Lucia',    'Reyes',      'H24', 6,  6, false, '', 'Gluten-free', true],
  ['Nathan',   'Shaw',       'H36', 6,  7, false, '', '', true],
  ['Chloe',    'Shaw',       'H36', 6,  8, false, '', '', true],
  // Table 7 — Cypress (Window, 8 seats)
  ['Grace',    'Foster',     'H26', 7,  1, false, '', '', true],
  ['Tyler',    'Foster',     'H26', 7,  2, false, '', '', true],
  ['Tunde',    'Adeyemi',    'H35', 7,  3, false, '', '', true],
  ['Bimpe',    'Adeyemi',    'H35', 7,  4, false, '', 'Nut allergy', true],
  ['Wei',      'Huang',      'H34', 7,  5, false, '', 'Vegan + Gluten-free', true],
  ['Priscilla','Okafor',     'H27', 7,  6, false, '', '', true],
  ['Brandon',  'Hayes',      'H28', 7,  7, false, '', '', true],
  ['Marcus',   'Bell',       'H38', 7,  8, false, '', '', true],
  // Table 8 — Rosewood (8 seats)
  ['Fatima',   'Al-Hassan',  'H30', 8,  1, false, '', 'Halal meal', true],
  ['Kevin',    "O'Brien",    'H31', 8,  2, false, '', '', true],
  ['Sarah',    "O'Brien",    'H31', 8,  3, false, '', '', true],
  ['Esi',      'Owusu',      'H37', 8,  4, false, '', '', false],
  ['Kwame',    'Osei',       'H01', 8,  5, false, '', 'NOTE: Duplicate name', true],
  // Tables 9-12: leave for planner to assign remaining/pending guests
  // Table 13 — Ivory (Children's table)
  ['Kwame Jr', 'Mensah',     'H03', 13, 1, false, 'Child', 'Kids meal', true],
  // Table 14 — Pearl (Accessible)
  ['Ama',      'Boateng',    'H04', 14, 1, false, 'Wheelchair', '', true],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['SEATING CHART']] });
  reqs.push({ mergeCells: { range: gridRange(SC, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SC, 0, 1, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SC, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Summary cards row 2-3 ─────────────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [[
    'Guests Seated:','',
    `=IFERROR(COUNTA(H6:H600),"")`, '',
    'Tables Used:','',
    `=IFERROR(SUMPRODUCT(1/COUNTIF(H6:H600,H6:H600)),"")`, '',
    'Conflicts:','',
    `=IFERROR(SUMPRODUCT((P6:P600="YES")*1),"")`, '',
    'Unconfirmed:','',
    `=IFERROR(SUMPRODUCT((T6:T600=FALSE)*1),"")`, '',
    'Over Capacity Tables:','',
    `=IFERROR(SUMPRODUCT((Q6:Q600="YES")*1),"")`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(SC, 1, 2, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Headers row 5 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A5`, values: [HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(SC, 4, 5, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SC, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ── Seating data rows 6+ ──────────────────────────────────────────────────
  seats.forEach((s, i) => {
    const row = 6 + i;
    const [first, last, hid, tableNum, seatNum, plusOne, arrangement, notes, confirmed] = s;
    data.push({ range: `${S}!A${row}`, values: [[i + 1]] });
    data.push({ range: `${S}!B${row}`, values: [[first]] });
    data.push({ range: `${S}!C${row}`, values: [[last]] });
    data.push({ range: `${S}!D${row}`, values: [[`=IFERROR(B${row}&" "&C${row},"")`]] });
    data.push({ range: `${S}!E${row}`, values: [[hid]] });
    // F: RSVP Status — lookup from Guest List by Full Name
    data.push({ range: `${S}!F${row}`, values: [[`=IFERROR(INDEX('Guest List & RSVP'!R$6:R$600,MATCH(D${row},'Guest List & RSVP'!D$6:D$600,0)),"Not Found")`]] });
    // G: Attending — "Yes" if RSVP=Attending
    data.push({ range: `${S}!G${row}`, values: [[`=IFERROR(IF(F${row}="Attending","Yes",IF(F${row}="Not Found","Check","No")),"No")`]] });
    data.push({ range: `${S}!H${row}`, values: [[tableNum]] });
    data.push({ range: `${S}!I${row}`, values: [[seatNum]] });
    // J: Table Name — lookup from Table Planner
    data.push({ range: `${S}!J${row}`, values: [[`=IFERROR(INDEX('Table Planner'!B$5:B$100,MATCH(H${row},'Table Planner'!A$5:A$100,0)),"")`]] });
    // K: Zone
    data.push({ range: `${S}!K${row}`, values: [[`=IFERROR(INDEX('Table Planner'!E$5:E$100,MATCH(H${row},'Table Planner'!A$5:A$100,0)),"")`]] });
    // L: Meal Preference — lookup
    data.push({ range: `${S}!L${row}`, values: [[`=IFERROR(INDEX('Guest List & RSVP'!J$6:J$600,MATCH(D${row},'Guest List & RSVP'!D$6:D$600,0)),"")`]] });
    // M: Dietary
    data.push({ range: `${S}!M${row}`, values: [[`=IFERROR(INDEX('Guest List & RSVP'!K$6:K$600,MATCH(D${row},'Guest List & RSVP'!D$6:D$600,0)),"")`]] });
    // N: Accessibility
    data.push({ range: `${S}!N${row}`, values: [[`=IFERROR(INDEX('Guest List & RSVP'!M$6:M$600,MATCH(D${row},'Guest List & RSVP'!D$6:D$600,0)),"")`]] });
    data.push({ range: `${S}!O${row}`, values: [[plusOne]] });
    // P: Household Conflict — same household at different tables
    data.push({ range: `${S}!P${row}`, values: [[`=IFERROR(IF(AND(E${row}<>"",E${row}<>"H00",COUNTIFS(E$6:E$600,E${row},H$6:H$600,"<>"&H${row})>0),"YES",""),"")` ]] });
    // Q: Over Capacity — table assigned count > capacity
    data.push({ range: `${S}!Q${row}`, values: [[`=IFERROR(IF(H${row}<>"",IF(IFERROR(INDEX('Table Planner'!G$5:G$100,MATCH(H${row},'Table Planner'!A$5:A$100,0)),1)<0,"YES",""),""),"")` ]] });
    data.push({ range: `${S}!R${row}`, values: [[arrangement]] });
    data.push({ range: `${S}!S${row}`, values: [[notes]] });
    data.push({ range: `${S}!T${row}`, values: [[confirmed]] });
  });

  // Checkboxes: O (Plus One), T (Confirmed)
  [14, 19].forEach(col => {
    reqs.push({ setDataValidation: {
      range: gridRange(SC, 5, 5 + seats.length, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Formula cells (auto columns) — blue-gray background
  const formulaCols = [3, 5, 6, 9, 10, 11, 12, 13, 15, 16]; // D,F,G,J,K,L,M,N,P,Q
  formulaCols.forEach(col => {
    reqs.push({ repeatCell: { range: gridRange(SC, 5, 5 + seats.length, col, col + 1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Alternating row colors
  for (let i = 0; i < seats.length; i++) {
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SC, 5 + i, 6 + i, 0, 20), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, foregroundColor: hex(C.mainText) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // CF: Conflict flag (col P) → red
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SC, 5, 200, 15, 16)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'YES' }] },
        format: { backgroundColor: hex(C.attention), textFormat: { bold: true, foregroundColor: hex(C.white) } },
      },
    },
    index: 0,
  }});

  // CF: Over Capacity (col Q) → red
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SC, 5, 200, 16, 17)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'YES' }] },
        format: { backgroundColor: hex(C.attention), textFormat: { bold: true, foregroundColor: hex(C.white) } },
      },
    },
    index: 0,
  }});

  // CF: Attending (col G) — Yes = green, No = red
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(SC, 5, 200, 6, 7)],
      booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Yes' }] },
        format: { backgroundColor: hex(C.success), textFormat: { foregroundColor: hex(C.white) } } },
    }, index: 0,
  }});
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(SC, 5, 200, 6, 7)],
      booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'No' }] },
        format: { backgroundColor: hex(C.attention), textFormat: { foregroundColor: hex(C.white) } } },
    }, index: 0,
  }});

  // Column widths
  const widths = [35, 75, 80, 130, 55, 90, 65, 60, 55, 90, 100, 90, 110, 100, 60, 90, 90, 130, 140, 70];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SC, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze row 5
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SC, gridProperties: { frozenRowCount: 5 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SC, tabColor: hex(C.primary) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'seating-data');
  await batchUpdate(id, reqs, 'seating-format');

  console.log('✓ Seating Chart built with', seats.length, 'assignments');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
