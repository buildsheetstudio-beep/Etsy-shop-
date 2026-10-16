'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WS = sheetMap['Wedding Setup'];
const S  = "'Wedding Setup'";

(async () => {
  const data = [];
  const reqs = [];

  // ── Title Row ─────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['ULTIMATE WEDDING PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(WS, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(WS, 0, 1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: WS, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // ── Subtitle ──────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [['Wedding Details & Configuration — Update your information here']] });
  reqs.push({ mergeCells: { range: gridRange(WS, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(WS, 1, 2, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 10 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // ── Section: Couple & Date ────────────────────────────────────────────────
  const sectionHeader = (row, label, col = 8) => {
    data.push({ range: `${S}!A${row}`, values: [[label]] });
    reqs.push({ mergeCells: { range: gridRange(WS, row - 1, row, 0, col), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(WS, row - 1, row, 0, col), cell: { userEnteredFormat: {
      backgroundColor: hex(C.accent),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
      horizontalAlignment: 'LEFT',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
  };

  const field = (row, label, value, isFormula = false) => {
    data.push({ range: `${S}!A${row}`, values: [[label]] });
    data.push({ range: `${S}!B${row}`, values: [[value]] });
    // Label style
    reqs.push({ repeatCell: { range: gridRange(WS, row - 1, row, 0, 1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { bold: true, foregroundColor: hex(C.mainText), fontSize: 10 },
      horizontalAlignment: 'RIGHT',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
    // Value cell style
    reqs.push({ repeatCell: { range: gridRange(WS, row - 1, row, 1, 4), cell: { userEnteredFormat: {
      backgroundColor: isFormula ? hex(C.formula) : hex(C.input),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 10 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
    // Merge value cells B:D
    reqs.push({ mergeCells: { range: gridRange(WS, row - 1, row, 1, 4), mergeType: 'MERGE_ALL' } });
  };

  sectionHeader(4, '  COUPLE & WEDDING DAY');
  field(5,  'Partner 1 Name',       'Amara Osei');
  field(6,  'Partner 2 Name',       'Marcus Thornton');
  field(7,  'Wedding Date',         '6/14/2026');
  field(8,  'Ceremony Start Time',  '4:00 PM');
  field(9,  'Reception Start Time', '6:00 PM');
  field(10, 'Ceremony Venue',       'The Botanical Gardens at Westlake');
  field(11, 'Reception Venue',      'Westlake Estate Ballroom');
  field(12, 'City, State',          'Austin, TX');
  field(13, 'Wedding Style/Theme',  'Garden Romance');
  field(14, 'Wedding Colors',       'Deep Atlantic Blue, Soft Persimmon, Ivory');

  sectionHeader(16, '  GUEST COUNT & BUDGET');
  field(17, 'Target Guest Count',   120);
  field(18, 'Ceremony Capacity',    150);
  field(19, 'Reception Capacity',   130);
  field(20, 'Total Budget',         35000);
  field(21, 'Currency Symbol',      '$');
  field(22, 'Rehearsal Dinner Date','6/13/2026');
  field(23, 'Morning Brunch Date',  '6/15/2026');

  sectionHeader(25, '  PLANNING CONTACTS');
  field(26, 'Planner / Coordinator', 'Sophia Rivera');
  field(27, 'Planner Phone',          '(512) 555-0192');
  field(28, 'Planner Email',          'sophia@riveraplanningatx.com');
  field(29, 'Emergency Day Contact',  'Jordan Osei (Maid of Honor)');
  field(30, 'Emergency Phone',        '(512) 555-0341');
  field(31, 'Officiant Name',         'Rev. David Nguyen');
  field(32, 'Caterer Name',           'Harvest Table Catering');

  sectionHeader(34, '  SUMMARY CARDS');

  // Summary KPI cards row 35-36 (label) / 37-38 (value)
  // Card 1: Days Until Wedding (A35:B38)
  data.push({ range: `${S}!A35`, values: [['DAYS UNTIL']] });
  data.push({ range: `${S}!A36`, values: [['WEDDING']] });
  data.push({ range: `${S}!A37`, values: [['=IFERROR(MAX(0,INT(B7-TODAY())),"—")']] });
  data.push({ range: `${S}!A38`, values: [['=IFERROR(IF(B7-TODAY()<0,"Past","Days Away"),"")' ]] });

  // Card 2: Budget
  data.push({ range: `${S}!C35`, values: [['TOTAL']] });
  data.push({ range: `${S}!C36`, values: [['BUDGET']] });
  data.push({ range: `${S}!C37`, values: [['=IFERROR("$"&TEXT(B20,"#,##0"),"—")']] });
  data.push({ range: `${S}!C38`, values: [['Target Amount']] });

  // Card 3: Guest Count
  data.push({ range: `${S}!E35`, values: [['TARGET']] });
  data.push({ range: `${S}!E36`, values: [['GUESTS']] });
  data.push({ range: `${S}!E37`, values: [['=IFERROR(B17,"—")']] });
  data.push({ range: `${S}!E38`, values: [['Guest Slots']] });

  // Card 4: Wedding Date
  data.push({ range: `${S}!G35`, values: [['WEDDING']] });
  data.push({ range: `${S}!G36`, values: [['DATE']] });
  data.push({ range: `${S}!G37`, values: [['=IFERROR(TEXT(B7,"mmm d"),"—")']] });
  data.push({ range: `${S}!G38`, values: [['=IFERROR(TEXT(B7,"yyyy"),"—")']] });

  const cardCols = [0, 2, 4, 6];
  cardCols.forEach(c => {
    reqs.push({ mergeCells: { range: gridRange(WS, 34, 36, c, c + 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(WS, 36, 38, c, c + 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(WS, 38, 39, c, c + 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(WS, 34, 36, c, c + 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ repeatCell: { range: gridRange(WS, 36, 38, c, c + 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 18 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ repeatCell: { range: gridRange(WS, 38, 39, c, c + 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { italic: true, foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
  });

  // Card row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: WS, dimension: 'ROWS', startIndex: 34, endIndex: 38 },
    properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: WS, dimension: 'ROWS', startIndex: 36, endIndex: 37 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Disclaimer ────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A41`, values: [['⚑  HOW TO USE THIS PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(WS, 40, 41, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(WS, 40, 41, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
    horizontalAlignment: 'LEFT',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  const instructions = [
    ['1. Start here — fill in your wedding details above (yellow cells are editable).'],
    ['2. Guest List & RSVP — add all guests; assign households for invitation grouping.'],
    ['3. Invitation Tracker — track each mailing by household with delivery/follow-up status.'],
    ['4. Table Planner — define tables, capacity, and zones first.'],
    ['5. Seating Chart — assign guests to tables; conflict flags highlight problems automatically.'],
    ['6. Vendors & Venue — log every vendor with contract, deposit, and payment details.'],
    ['7. Reception Planner — build your timeline segment by segment.'],
    ['8. Food, Drinks & Cake — plan your menu with dietary coverage tracking.'],
    ['9. Budget & Payments — track all spending against your category budgets.'],
    ['10. Wedding Checklist — 70+ tasks across 10 planning phases with progress tracking.'],
    ['11. Dashboard — real-time overview of all key metrics and upcoming tasks.'],
    [''],
    ['YELLOW cells = enter your data here.  BLUE-GRAY cells = auto-calculated (do not edit).'],
  ];
  instructions.forEach((row, i) => {
    data.push({ range: `${S}!A${42 + i}`, values: [row] });
    reqs.push({ mergeCells: { range: gridRange(WS, 41 + i, 42 + i, 0, 8), mergeType: 'MERGE_ALL' } });
  });
  reqs.push({ repeatCell: { range: gridRange(WS, 41, 41 + instructions.length, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
    textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
    wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' }});
  // Last instruction line (legend) — bold
  reqs.push({ repeatCell: { range: gridRange(WS, 41 + instructions.length - 1, 41 + instructions.length, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [220, 200, 120, 120, 120, 120, 120, 120];
  colWidths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: WS, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze rows 1-2
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: WS, gridProperties: { frozenRowCount: 2 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: WS, tabColor: hex(C.secondary) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'setup-data');
  await batchUpdate(id, reqs, 'setup-format');

  console.log('✓ Wedding Setup tab built');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
