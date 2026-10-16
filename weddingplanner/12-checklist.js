'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CL = sheetMap['Wedding Checklist'];
const S  = "'Wedding Checklist'";

const HEADERS = [
  '#', 'Phase', 'Phase Name', 'Category', 'Task', 'Description',
  'Due Date', 'Days Until Due (auto)', 'Assigned To', 'Priority',
  'Status', 'Completed Date', 'Vendor Related', 'Cost Estimate', 'Notes', 'Urgency (auto)',
];

// 70+ tasks across 10 planning phases
// [phase, phaseName, category, task, desc, dueDate, assignedTo, priority, status, completedDate, vendorRelated, costEstimate, notes]
const tasks = [
  // Phase 1 — Early Planning (12+ months)
  [1,'Early Planning',     'Venue',       'Set wedding date',          'Choose ceremony + reception date',                               '10/1/2025',  'Both',      'Critical', 'Completed', '9/15/2025',  false, 0,    ''],
  [1,'Early Planning',     'Budget',      'Establish total budget',    'Agree on total budget range and contributor split',               '10/1/2025',  'Both',      'Critical', 'Completed', '9/20/2025',  false, 0,    ''],
  [1,'Early Planning',     'Venue',       'Research venues',           'Visit 5+ ceremony and reception venues, request packages',        '10/15/2025', 'Both',      'Critical', 'Completed', '10/10/2025', true,  0,    ''],
  [1,'Early Planning',     'Venue',       'Book ceremony venue',       'Sign contract and pay deposit for ceremony venue',                '11/1/2025',  'Both',      'Critical', 'Completed', '10/28/2025', true,  4500, ''],
  [1,'Early Planning',     'Venue',       'Book reception venue',      'Sign contract and pay deposit for reception venue',               '11/1/2025',  'Both',      'Critical', 'Completed', '10/28/2025', true,  8000, ''],
  [1,'Early Planning',     'Planning',    'Hire wedding coordinator',  'Research and hire full-service planner or day-of coordinator',    '11/15/2025', 'Both',      'High',     'Completed', '11/12/2025', true,  4200, ''],
  [1,'Early Planning',     'Guest List',  'Draft initial guest list',  'Compile combined guest list with household groupings',            '11/1/2025',  'Both',      'High',     'Completed', '10/25/2025', false, 0,    ''],
  // Phase 2 — Venue & Date Confirmed (10-12 months)
  [2,'Venue & Date',       'Photography', 'Book photographer',         'Review portfolios, schedule meetings, sign contract + deposit',   '12/1/2025',  'Both',      'Critical', 'Completed', '11/30/2025', true,  4800, ''],
  [2,'Venue & Date',       'Videography', 'Book videographer',         'Review reels, confirm style, sign contract',                     '12/15/2025', 'Both',      'High',     'Completed', '12/10/2025', true,  3200, ''],
  [2,'Venue & Date',       'Planning',    'Create planning timeline',  'Build 12-month planning roadmap with coordinator',               '12/1/2025',  'Amara',     'High',     'Completed', '11/28/2025', false, 0,    ''],
  [2,'Venue & Date',       'Budget',      'Open wedding savings account','Set up dedicated account for all wedding payments',            '11/20/2025', 'Marcus',    'Medium',   'Completed', '11/18/2025', false, 0,    ''],
  [2,'Venue & Date',       'Guest List',  'Finalize guest count',      'Lock in final guest list and RSVP target count',                 '1/1/2026',   'Both',      'Critical', 'Completed', '12/20/2025', false, 0,    ''],
  // Phase 3 — Key Vendors (8-10 months)
  [3,'Key Vendors',        'Catering',    'Book caterer',              'Taste test, negotiate package, sign contract',                   '2/1/2026',   'Both',      'Critical', 'Completed', '1/25/2026',  true,  18000,''],
  [3,'Key Vendors',        'Florals',     'Book florist',              'Share color palette, confirm style, sign contract',              '2/15/2026',  'Amara',     'High',     'Completed', '2/12/2026',  true,  5500, ''],
  [3,'Key Vendors',        'Music',       'Book DJ or band',           'Discuss playlist, review contract, sign deposit',               '2/20/2026',  'Both',      'High',     'Completed', '2/18/2026',  true,  2800, ''],
  [3,'Key Vendors',        'Beauty',      'Book hair & makeup artist', 'Schedule trials, confirm team size for bridal party',           '3/1/2026',   'Amara',     'High',     'Completed', '2/28/2026',  true,  1800, ''],
  [3,'Key Vendors',        'Officiant',   'Book officiant',            'Connect with Rev. David Nguyen, discuss ceremony structure',     '2/1/2026',   'Both',      'High',     'Completed', '1/20/2026',  true,  800,  ''],
  [3,'Key Vendors',        'Honeymoon',   'Plan honeymoon',            'Book flights, hotels, and activities for Italy trip',           '3/15/2026',  'Both',      'Medium',   'Completed', '3/10/2026',  true,  8500, ''],
  [3,'Key Vendors',        'Rentals',     'Book rentals & lighting',   'Linens, chairs, uplighting — confirm rental packages',          '3/20/2026',  'Marcus',    'Medium',   'Completed', '3/18/2026',  true,  4400, ''],
  // Phase 4 — Guest & Invitations (6-8 months)
  [4,'Guest & Invitations','Invitations', 'Design save-the-dates',     'Work with Paper & Petal, approve design, order',                '12/15/2025', 'Amara',     'High',     'Completed', '12/10/2025', true,  200,  ''],
  [4,'Guest & Invitations','Invitations', 'Mail save-the-dates',       'Mail to all 38 households',                                     '1/10/2026',  'Amara',     'High',     'Completed', '1/8/2026',   false, 0,    ''],
  [4,'Guest & Invitations','Invitations', 'Design wedding invitations', 'Finalize suite: invite, RSVP, details, menus, envelope',        '2/15/2026',  'Amara',     'High',     'Completed', '2/12/2026',  true,  1200, ''],
  [4,'Guest & Invitations','Invitations', 'Print & address envelopes', 'Print, assemble, and hand-address or calligraph envelopes',      '4/20/2026',  'Amara',     'High',     'Completed', '4/18/2026',  false, 0,    'Calligrapher booked'],
  [4,'Guest & Invitations','Invitations', 'Mail invitations',          'Drop at post office with correct postage',                       '5/1/2026',   'Amara',     'Critical', 'Completed', '5/1/2026',   false, 0,    ''],
  [4,'Guest & Invitations','Guest List',  'Collect dietary info',      'Follow up with guests on meal preference + restrictions',        '5/20/2026',  'Amara',     'High',     'Completed', '5/18/2026',  false, 0,    ''],
  [4,'Guest & Invitations','Planning',    'Book hotel room block',     'Arrange group rate for out-of-town guests at Hampton Inn',       '3/1/2026',   'Marcus',    'High',     'Completed', '2/28/2026',  true,  0,    ''],
  // Phase 5 — Details & Decor (4-6 months)
  [5,'Details & Decor',    'Florals',     'Floral design meeting #2',  'Finalize centerpiece designs, confirm bloom selections',         '4/1/2026',   'Amara',     'High',     'Completed', '3/28/2026',  true,  0,    ''],
  [5,'Details & Decor',    'Stationery',  'Order programs & menus',    'Design and print ceremony programs and reception menus',         '4/15/2026',  'Amara',     'Medium',   'Completed', '4/12/2026',  true,  200,  ''],
  [5,'Details & Decor',    'Cake',        'Wedding cake tasting',      'Visit Sugar & Spire for cake tasting and design consult',        '4/15/2026',  'Both',      'High',     'Completed', '4/10/2026',  true,  0,    ''],
  [5,'Details & Decor',    'Cake',        'Finalize cake design',      'Confirm flavor, design, delivery time',                         '4/20/2026',  'Amara',     'High',     'Completed', '4/18/2026',  true,  2200, ''],
  [5,'Details & Decor',    'Decor',       'Create seating chart draft','Assign guests to tables using Table Planner tab',               '5/15/2026',  'Both',      'High',     'Completed', '5/10/2026',  false, 0,    ''],
  [5,'Details & Decor',    'Music',       'DJ meeting — playlist',     'Finalize must-play and do-not-play songs, special moments',      '4/30/2026',  'Both',      'High',     'Completed', '4/28/2026',  true,  0,    ''],
  [5,'Details & Decor',    'Favors',      'Order guest favors',        'Honey jar + seed packet combo — 120 units',                     '5/1/2026',   'Amara',     'Medium',   'Completed', '4/28/2026',  true,  600,  ''],
  // Phase 6 — Attire & Beauty (3-4 months)
  [6,'Attire & Beauty',    'Attire',      'Purchase wedding dress',    'Final fitting and pickup at bridal boutique',                    '3/1/2026',   'Amara',     'Critical', 'Completed', '2/25/2026',  false, 2200, ''],
  [6,'Attire & Beauty',    'Attire',      'Bridesmaids attire',        'Select and order bridesmaids dresses',                          '3/15/2026',  'Amara',     'High',     'Completed', '3/10/2026',  false, 600,  '3 bridesmaids'],
  [6,'Attire & Beauty',    'Attire',      'Groom + groomsmen attire',  'Suit or tux fittings for groom and 2 groomsmen + best man',     '3/15/2026',  'Marcus',    'High',     'Completed', '3/10/2026',  false, 500,  ''],
  [6,'Attire & Beauty',    'Attire',      'Dress alterations',         'Schedule and complete alterations for bridal gown',             '4/30/2026',  'Amara',     'High',     'In Progress','',           false, 300,  'Final fitting May 20'],
  [6,'Attire & Beauty',    'Beauty',      'Hair & makeup trial',       'Schedule trial run with Luxe Bridal Beauty',                    '5/1/2026',   'Amara',     'High',     'Completed', '4/25/2026',  true,  0,    ''],
  [6,'Attire & Beauty',    'Jewelry',     'Select bridal jewelry',     'Finalize earrings, necklace, bracelet for ceremony',            '5/15/2026',  'Amara',     'Medium',   'Completed', '5/12/2026',  false, 400,  ''],
  [6,'Attire & Beauty',    'Rings',       'Finalize wedding bands',    'Pick up or order wedding bands, confirm engravings',            '5/1/2026',   'Both',      'Critical', 'Completed', '4/28/2026',  false, 1800, ''],
  // Phase 7 — Final Confirmations (1-3 months)
  [7,'Final Confirmations','Planning',    'Final venue walkthrough',   'Walk both venues with coordinator, finalize layout and setup',   '5/15/2026',  'Both',      'Critical', 'Completed', '5/14/2026',  true,  0,    ''],
  [7,'Final Confirmations','RSVP',        'Final RSVP deadline',       'Close RSVPs, follow up on non-responses',                       '5/25/2026',  'Amara',     'Critical', 'Completed', '5/25/2026',  false, 0,    ''],
  [7,'Final Confirmations','Catering',    'Confirm final headcount',   'Submit final attendee count + dietary breakdown to caterer',     '6/1/2026',   'Amara',     'Critical', 'In Progress','',          true,  0,    ''],
  [7,'Final Confirmations','Seating',     'Finalize seating chart',    'Complete and print final seating chart and escort cards',       '6/5/2026',   'Amara',     'Critical', 'In Progress','',          false, 0,    ''],
  [7,'Final Confirmations','Music',       'Final DJ briefing',         'Review timeline, entrance songs, special moment music',         '6/1/2026',   'Both',      'High',     'In Progress','',          true,  0,    ''],
  [7,'Final Confirmations','Vendor',      'Confirm all vendor logistics','Email or call every vendor to confirm timing and access',     '6/7/2026',   'Sophia',    'Critical', 'Not Started','',          true,  0,    ''],
  [7,'Final Confirmations','Payments',    'Finalize vendor payments',  'Clear all outstanding balances before wedding day',             '6/10/2026',  'Both',      'Critical', 'Not Started','',          true,  0,    ''],
  [7,'Final Confirmations','Stationery',  'Prepare place cards',       'Print and sort escort cards / place cards by table',           '6/7/2026',   'Amara',     'High',     'Not Started','',          false, 0,    ''],
  [7,'Final Confirmations','Ceremony',    'Submit ceremony script',    'Send final approved ceremony script to officiant',             '6/5/2026',   'Both',      'High',     'Not Started','',          true,  0,    ''],
  // Phase 8 — Last Month
  [8,'Last Month',         'Beauty',      'Final dress fitting',       'Collect gown after final alterations',                         '5/20/2026',  'Amara',     'Critical', 'Not Started','',          false, 0,    ''],
  [8,'Last Month',         'Honeymoon',   'Pack honeymoon bag',        'Confirm passports, pack carry-on and checked bags',            '6/12/2026',  'Both',      'Medium',   'Not Started','',          false, 0,    'Passports current'],
  [8,'Last Month',         'Gifts',       'Prepare gifts for vendors', 'Assemble thank-you envelopes + tips for day-of vendors',       '6/10/2026',  'Both',      'Medium',   'Not Started','',          false, 0,    ''],
  [8,'Last Month',         'Planning',    'Create wedding day timeline','Print 10 copies of run-of-show for key people',              '6/10/2026',  'Sophia',    'High',     'Not Started','',          false, 0,    ''],
  [8,'Last Month',         'Guest List',  'Prepare emergency kit',     'Aspirin, safety pins, stain pen, breath mints, backup shoes', '6/12/2026',  'Jordan',    'Medium',   'Not Started','',          false, 0,    ''],
  // Phase 9 — Final Week
  [9,'Final Week',         'Rehearsal',   'Rehearsal dinner',          'Run through ceremony at venue with wedding party + families',  '6/13/2026',  'All',       'Critical', 'Not Started','',          true,  0,    ''],
  [9,'Final Week',         'Vendor',      'Confirm vendor deliveries', 'Reconfirm florist, rental, and cake delivery windows',        '6/12/2026',  'Sophia',    'Critical', 'Not Started','',          true,  0,    ''],
  [9,'Final Week',         'Beauty',      'Bridal party nails',        'Manicures for bride + bridesmaids',                           '6/12/2026',  'Amara',     'Low',      'Not Started','',          false, 0,    ''],
  [9,'Final Week',         'Planning',    'Delegate day-of tasks',     'Brief key people on roles for ceremony + reception',          '6/12/2026',  'Sophia',    'High',     'Not Started','',          false, 0,    ''],
  [9,'Final Week',         'Honeymoon',   'Check-in for flights',      'Online check-in, download boarding passes',                   '6/14/2026',  'Marcus',    'Medium',   'Not Started','',          false, 0,    ''],
  [9,'Final Week',         'Payments',    'Prepare vendor gratuities', 'Cash tip envelopes for caterer staff, DJ, hair & makeup',     '6/12/2026',  'Marcus',    'High',     'Not Started','',          false, 0,    ''],
  // Phase 10 — Wedding Day
  [10,'Wedding Day',       'Beauty',      'Morning beauty appointments','Hair and makeup for bridal party (9am–1pm)',                 '6/14/2026',  'Amara',     'Critical', 'Not Started','',          true,  0,    ''],
  [10,'Wedding Day',       'Vendor',      'Greet vendors on arrival',  'Welcome caterer, florist, DJ, lighting crew',                 '6/14/2026',  'Sophia',    'Critical', 'Not Started','',          true,  0,    ''],
  [10,'Wedding Day',       'Ceremony',    'Ceremony (4:00 PM)',        'Walk down the aisle, exchange vows and rings',                '6/14/2026',  'Both',      'Critical', 'Not Started','',          false, 0,    ''],
  [10,'Wedding Day',       'Photos',      'Wedding portraits session', 'Couple photos in garden after ceremony',                      '6/14/2026',  'Isaiah',    'High',     'Not Started','',          true,  0,    ''],
  [10,'Wedding Day',       'Reception',   'Reception (6:00 PM)',       'Grand entrance, dinner, toasts, dancing, sendoff',            '6/14/2026',  'Sophia',    'Critical', 'Not Started','',          true,  0,    ''],
  [10,'Wedding Day',       'Payments',    'Distribute vendor payments','Pay remaining balances and tips to vendors in person',        '6/14/2026',  'Marcus',    'Critical', 'Not Started','',          true,  0,    ''],
  [10,'Wedding Day',       'Memories',    'Collect cards & gifts',     'Designate person to collect cards and gifts during reception', '6/14/2026', 'Jason',     'Medium',   'Not Started','',          false, 0,    ''],
  [10,'Wedding Day',       'Memories',    'Grand sendoff (10:45 PM)', 'Sparkler exit + final photos with LifeFrame Films',           '6/14/2026',  'Sophia',    'High',     'Not Started','',          true,  0,    ''],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['WEDDING CHECKLIST — 70+ TASKS ACROSS 10 PHASES']] });
  reqs.push({ mergeCells: { range: gridRange(CL, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CL, 0, 1, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 13 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: CL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Summary row 2 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [[
    'Total Tasks:','', `=IFERROR(COUNTA(E5:E500),"")`, '',
    'Completed:','', `=IFERROR(SUMPRODUCT((K5:K500="Completed")*1),"")`, '',
    'In Progress:','', `=IFERROR(SUMPRODUCT((K5:K500="In Progress")*1),"")`, '',
    '% Complete:','', `=IFERROR(SUMPRODUCT((K5:K500="Completed")*1)/COUNTA(E5:E500),0)`, '',
    'Overdue:','', `=IFERROR(SUMPRODUCT((K5:K500<>"Completed")*(G5:G500<>"")*((DATEVALUE(TEXT(G5:G500,"MM/DD/YYYY")))<TODAY())*1),"")`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(CL, 1, 2, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  // % complete as percent
  reqs.push({ repeatCell: { range: gridRange(CL, 1, 2, 13, 14), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // ── Headers row 4 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A4`, values: [HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(CL, 3, 4, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: CL, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ── Task data rows ────────────────────────────────────────────────────────
  let currentPhase = null;
  let adjustedRow = 5;

  tasks.forEach((task, i) => {
    const [phase, phaseName, cat, taskName, desc, dueDate, assignedTo, priority, status, completedDate, vendorRelated, costEstimate, notes] = task;

    // Phase sub-header row when phase changes
    if (phase !== currentPhase) {
      currentPhase = phase;
      data.push({ range: `${S}!A${adjustedRow}`, values: [[`PHASE ${phase} — ${phaseName.toUpperCase()}`]] });
      reqs.push({ mergeCells: { range: gridRange(CL, adjustedRow - 1, adjustedRow, 0, 16), mergeType: 'MERGE_ALL' } });
      reqs.push({ repeatCell: { range: gridRange(CL, adjustedRow - 1, adjustedRow, 0, 16), cell: { userEnteredFormat: {
        backgroundColor: hex(C.accent),
        textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
        horizontalAlignment: 'LEFT',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
      adjustedRow++;
    }

    const row = adjustedRow;
    data.push({ range: `${S}!A${row}`, values: [[i + 1]] });
    data.push({ range: `${S}!B${row}`, values: [[phase]] });
    data.push({ range: `${S}!C${row}`, values: [[phaseName]] });
    data.push({ range: `${S}!D${row}`, values: [[cat]] });
    data.push({ range: `${S}!E${row}`, values: [[taskName]] });
    data.push({ range: `${S}!F${row}`, values: [[desc]] });
    data.push({ range: `${S}!G${row}`, values: [[dueDate]] });
    // H: Days Until Due
    data.push({ range: `${S}!H${row}`, values: [[`=IFERROR(IF(G${row}="","",INT(DATEVALUE(TEXT(G${row},"MM/DD/YYYY"))-TODAY())),"")` ]] });
    data.push({ range: `${S}!I${row}`, values: [[assignedTo]] });
    data.push({ range: `${S}!J${row}`, values: [[priority]] });
    data.push({ range: `${S}!K${row}`, values: [[status]] });
    data.push({ range: `${S}!L${row}`, values: [[completedDate]] });
    data.push({ range: `${S}!M${row}`, values: [[vendorRelated]] });
    data.push({ range: `${S}!N${row}`, values: [[costEstimate || '']] });
    data.push({ range: `${S}!O${row}`, values: [[notes]] });
    // P: Urgency flag
    data.push({ range: `${S}!P${row}`, values: [[`=IFERROR(IF(K${row}="Completed","✓ Done",IF(G${row}="","",IF(INT(DATEVALUE(TEXT(G${row},"MM/DD/YYYY"))-TODAY())<0,"OVERDUE",IF(INT(DATEVALUE(TEXT(G${row},"MM/DD/YYYY"))-TODAY())<=7,"THIS WEEK",IF(INT(DATEVALUE(TEXT(G${row},"MM/DD/YYYY"))-TODAY())<=30,"THIS MONTH","On Track"))))),"")` ]] });

    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(CL, row - 1, row, 0, 16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, foregroundColor: hex(C.mainText) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

    adjustedRow++;
  });

  const lastDataRow = adjustedRow;

  // Checkboxes: M (Vendor Related)
  reqs.push({ setDataValidation: {
    range: gridRange(CL, 4, lastDataRow, 12, 13),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
  }});

  // Dropdowns
  reqs.push({ setDataValidation: {
    range: gridRange(CL, 4, lastDataRow + 100, 9, 10),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Critical', 'High', 'Medium', 'Low',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});
  reqs.push({ setDataValidation: {
    range: gridRange(CL, 4, lastDataRow + 100, 10, 11),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Not Started', 'In Progress', 'Completed', 'Blocked', 'Delegated', 'N/A',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Formula cells H (Days Until) and P (Urgency) — blue-gray
  reqs.push({ repeatCell: { range: gridRange(CL, 4, lastDataRow, 7, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
  }}, fields: 'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(CL, 4, lastDataRow, 15, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Cost estimate currency format
  reqs.push({ repeatCell: { range: gridRange(CL, 4, lastDataRow, 13, 14), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // CF: Status colors
  const statusCF = [
    { text: 'Completed',   bg: C.success   },
    { text: 'In Progress', bg: C.warning   },
    { text: 'Blocked',     bg: C.attention },
    { text: 'Not Started', bg: C.neutral   },
  ];
  statusCF.forEach(({ text, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: { ranges: [gridRange(CL, 4, lastDataRow, 10, 11)],
        booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bg), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
      }, index: 0,
    }});
  });

  // CF: Urgency flag
  const urgencyCF = [
    { text: 'OVERDUE',    bg: C.attention },
    { text: 'THIS WEEK',  bg: C.warning   },
    { text: 'THIS MONTH', bg: C.accent    },
    { text: '✓ Done',     bg: C.success   },
  ];
  urgencyCF.forEach(({ text, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: { ranges: [gridRange(CL, 4, lastDataRow, 15, 16)],
        booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bg), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
      }, index: 0,
    }});
  });

  // CF: Priority colors
  [
    { text: 'Critical', bg: C.attention },
    { text: 'High',     bg: C.secondary },
    { text: 'Medium',   bg: C.warning   },
  ].forEach(({ text, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: { ranges: [gridRange(CL, 4, lastDataRow, 9, 10)],
        booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bg), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
      }, index: 0,
    }});
  });

  // Column widths
  const widths = [35, 45, 100, 90, 180, 220, 80, 70, 100, 70, 90, 90, 70, 80, 180, 90];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CL, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze rows 1-4
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: CL, gridProperties: { frozenRowCount: 4 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: CL, tabColor: hex(C.success) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'checklist-data');
  await batchUpdate(id, reqs, 'checklist-format');

  console.log('✓ Wedding Checklist built with', tasks.length, 'tasks across 10 phases');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
