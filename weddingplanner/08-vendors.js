'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const VV = sheetMap['Vendors & Venue'];
const S  = "'Vendors & Venue'";

const HEADERS = [
  '#', 'Category', 'Company Name', 'Contact Name', 'Phone', 'Email',
  'Contract Signed', 'Contract Date', 'Service Description', 'Booking Status',
  'Event Date', 'Start Time', 'End Time', 'Location',
  'Package / Tier', 'Quoted Price', 'Deposit Amount', 'Deposit Paid',
  'Deposit Date', 'Total Cost', 'Amount Paid', 'Balance Due (auto)', 'Website', 'Notes',
];

// Venue section fields (rows 4-18 in col A/B)
const venueFields = [
  ['Ceremony Venue Name',    'The Botanical Gardens at Westlake'],
  ['Ceremony Address',       '2200 Garden View Blvd, Austin, TX 78703'],
  ['Ceremony Capacity',      '150'],
  ['Ceremony Contact',       'Maya Stern — (512) 555-0101'],
  ['Ceremony Venue Fee',     '$4,500'],
  ['Ceremony Contract Signed', 'Yes'],
  ['Reception Venue Name',   'Westlake Estate Ballroom'],
  ['Reception Address',      '2210 Garden View Blvd, Austin, TX 78703 (adjacent)'],
  ['Reception Capacity',     '130'],
  ['Reception Contact',      'Maya Stern — (512) 555-0101'],
  ['Reception Venue Fee',    '$8,000'],
  ['Reception Contract Signed', 'Yes'],
  ['Parking / Shuttle',      'Self-park + valet available; shuttle from Hampton Inn'],
  ['Venue Notes',            'Both venues on same property. Ceremony to reception walkover.'],
];

// 18 vendors
const vendors = [
  ['Caterer',                'Harvest Table Catering',     'Chef Rosa Delgado',    '(512) 555-0201', 'rosa@harvesttableATX.com',      true, '3/15/2026', 'Full-service catering — 3-course dinner + cocktail hour', 'Booked',   '6/14/2026', '5:00 PM',  '11:00 PM', 'Westlake Ballroom', 'Silver Package',         18000,  3600, true,  '3/20/2026', 18000, 9000,  'harvesttableatx.com', ''],
  ['Photographer',           'Golden Hour Studio',         'Isaiah Torres',        '(512) 555-0202', 'hello@goldenhour.photo',        true, '2/10/2026', '8-hr coverage, engagement shoot included, 2 photographers', 'Booked', '6/14/2026', '3:00 PM',  '11:00 PM', 'Full venue', '8-Hour Signature', 4800, 1200, true, '2/15/2026', 4800, 2400, 'goldenhour.photo', ''],
  ['Videographer',           'LifeFrame Films',            'Priya Kapoor',         '(512) 555-0203', 'priya@lifeframe.co',            true, '2/20/2026', 'Ceremony + reception + highlight reel (5 min)', 'Booked',    '6/14/2026', '3:30 PM',  '10:30 PM', 'Full venue', 'Standard',          3200,  800, true,  '2/25/2026',  3200, 1600, 'lifeframe.co', ''],
  ['Florist',                'Bloom & Wild Designs',       'Hazel Okonkwo',        '(512) 555-0204', 'hazel@bloomwild.com',           true, '3/1/2026',  'Ceremony arch, bridal bouquet, bridesmaids, centerpieces', 'Booked', '6/14/2026', '2:00 PM',  '4:00 PM',  'Both venues',  'Full Floral',     5500, 1100, true, '3/5/2026',  5500, 2750, 'bloomwild.com', 'Delivery 2pm'],
  ['DJ / Band',              'Vibe Events DJ',             'Marcus Lee',           '(512) 555-0205', 'marcus@vibeDJ.com',             true, '3/10/2026', 'Cocktail hour + reception + MC services',  'Booked',               '6/14/2026', '5:30 PM',  '11:00 PM', 'Ballroom',     'Premium',         2800,  700, true,  '3/15/2026',  2800, 1400, 'vibeDJ.com', ''],
  ['Hair & Makeup',          'Luxe Bridal Beauty',         'Simone Nwosu',         '(512) 555-0206', 'simone@luxebridal.com',         true, '4/1/2026',  'Bride + 3 bridesmaids hair and makeup',    'Booked',               '6/14/2026', '9:00 AM',  '1:00 PM',  'Hotel Suite',  'Bridal Party Pkg',1800,  450, true,  '4/5/2026',   1800,  900, 'luxebridal.com', 'Airbrush makeup'],
  ['Officiant',              'Rev. David Nguyen',          'David Nguyen',         '(512) 555-0207', 'rev.david@faithjourneys.org',   true, '1/15/2026', 'Custom ceremony, rehearsal coordination',  'Booked',               '6/14/2026', '4:00 PM',  '4:30 PM',  'Botanical Gardens','Custom',       800,   0, false, '',           800,   400, '',              'Provided 3 script drafts'],
  ['Cake / Bakery',          'Sugar & Spire Bakery',       'Claudette Martin',     '(512) 555-0208', 'claudette@sugarspire.com',      true, '4/15/2026', '4-tier cake (vanilla/lemon) + dessert bar', 'Booked',              '6/14/2026', '5:00 PM',  '5:30 PM',  'Ballroom',     'Tier 4 + Desserts',2200,  550, true,  '4/20/2026',  2200,  550, 'sugarspire.com', 'Delivery at 5pm'],
  ['Photo Booth',            'ClickBox ATX',               'James Ruiz',           '(512) 555-0209', 'james@clickboxatx.com',         true, '4/10/2026', '3-hr open-air booth, unlimited prints, custom backdrop', 'Booked', '6/14/2026', '7:00 PM', '10:00 PM', 'Ballroom Foyer','Standard Pkg',    900,  225, true,  '4/12/2026',   900,  450, 'clickboxatx.com', ''],
  ['Transportation',         'Luxury Limo ATX',            'Carl Dixon',           '(512) 555-0210', 'carl@luxlimo.com',              true, '3/25/2026', 'Bridal party sprinter + guest shuttle ×3', 'Booked',               '6/14/2026', '2:30 PM',  '12:00 AM', 'Multiple',     'Wedding Bundle',  1600,  400, true,  '3/28/2026',  1600,  800, 'luxlimo.com', 'Shuttle route: Hampton Inn ↔ venue'],
  ['Stationery',             'Paper & Petal Studio',       'Anastasia Cruz',       '(512) 555-0211', 'hello@paperandpetal.com',       true, '2/1/2026',  'Invitations, programs, menus, place cards, escort cards', 'Booked', '',          '',         '',         'Delivered',    'Full Suite',      1400,  350, true,  '2/5/2026',   1400, 1400, 'paperandpetal.com', 'All delivered 4/20'],
  ['Rentals',                'Event Elegance Rentals',     'Sandra Bloom',         '(512) 555-0212', 'sandra@eventelegance.com',      true, '3/20/2026', 'Linens, charger plates, chiavari chairs, lounge furniture', 'Booked', '6/14/2026', '12:00 PM', '12:00 AM', 'Both venues', 'Luxury Package', 2600,  650, true, '3/22/2026', 2600, 1300, 'eventelegance.com', ''],
  ['Lighting',               'Luminary Events',            'Derek Osei',           '(512) 555-0213', 'derek@luminaryevents.com',      true, '3/18/2026', 'Bistro lights, uplighting, dance floor spotlights', 'Booked',    '6/14/2026', '2:00 PM',  '11:30 PM', 'Both venues', 'Standard',        1800,  450, true,  '3/20/2026',  1800,  900, 'luminaryevents.com', ''],
  ['Coordinator / Planner',  'Rivera Planning ATX',        'Sophia Rivera',        '(512) 555-0192', 'sophia@riveraplanningatx.com',  true, '11/15/2025','Full-service planning + day-of coordination', 'Booked',             '6/14/2026', '12:00 PM', '11:59 PM', 'Full venue', 'Full Service',    4200, 1050, true,  '11/20/2025', 4200, 2100, 'riveraplanningatx.com', ''],
  ['Hotel Room Block',       'Hampton Inn Austin West',    'Booking Team',         '(512) 555-0214', 'groups@hamptonwest.com',        false, '',          '20-room block for out-of-town guests',     'Booked',               '6/13/2026', '3:00 PM',  '6/15/2026','5400 Westlake Dr','Group Rate',   0,     0, false, '',           0,      0,   'hamptoninnaustin.com', 'Complimentary shuttle included'],
  ['Honeymoon',              'Golden Gate Travel',         'Amaya White',          '(512) 555-0215', 'amaya@ggtravel.com',            true, '4/5/2026',  'Flights + 10-night Tuscany & Amalfi Coast itinerary', 'Booked',   '6/16/2026', '',         '',         'Italy',        'Honeymoon Pkg',  8500, 2000, true,  '4/10/2026',  8500, 4000, 'ggtravel.com', 'Depart June 16'],
  ['Favors & Gifts',         'Artisan Gifting Co.',        'Rachel Park',          '(512) 555-0216', 'rachel@artisangifting.com',     false, '',          '120× guest favors (honey jar + seed packet)', 'Booked',             '',          '',         '',         'Delivered',    'Custom',          600,    0, false, '',           600,   300, 'artisangifting.com', ''],
  ['Coordinator / Planner',  'Day-Of Emergency Contact',  'Sophia Rivera',        '(512) 555-0192', 'sophia@riveraplanningatx.com',  true, '11/15/2025','Backup contact — same vendor as above',     'Booked',              '6/14/2026', '',         '',         '',             '',                0,     0, false, '',           0,      0,  '', 'Duplicate record for day-of contact card'],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['VENDORS & VENUE']] });
  reqs.push({ mergeCells: { range: gridRange(VV, 0, 1, 0, 24), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(VV, 0, 1, 0, 24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: VV, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Venue Section header ───────────────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [['  VENUE DETAILS']] });
  reqs.push({ mergeCells: { range: gridRange(VV, 1, 2, 0, 24), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(VV, 1, 2, 0, 24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  venueFields.forEach((vf, i) => {
    const row = 3 + i;
    data.push({ range: `${S}!A${row}`, values: [[vf[0]]] });
    data.push({ range: `${S}!B${row}`, values: [[vf[1]]] });
    reqs.push({ repeatCell: { range: gridRange(VV, row - 1, row, 0, 1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.mainText) },
      horizontalAlignment: 'RIGHT',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
    reqs.push({ repeatCell: { range: gridRange(VV, row - 1, row, 1, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
    reqs.push({ mergeCells: { range: gridRange(VV, row - 1, row, 1, 8), mergeType: 'MERGE_ALL' } });
  });

  const vendorStartRow = 3 + venueFields.length + 2; // ~21

  // ── Vendor Section header ─────────────────────────────────────────────────
  data.push({ range: `${S}!A${vendorStartRow - 1}`, values: [['  VENDOR LOG']] });
  reqs.push({ mergeCells: { range: gridRange(VV, vendorStartRow - 2, vendorStartRow - 1, 0, 24), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(VV, vendorStartRow - 2, vendorStartRow - 1, 0, 24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Vendor summary row
  const sumRow = vendorStartRow;
  data.push({ range: `${S}!A${sumRow}`, values: [[
    'Total Vendors:','',
    `=IFERROR(COUNTA(B${sumRow+2}:B${sumRow+100}),"")`, '',
    'Contracts Signed:','',
    `=IFERROR(SUMPRODUCT((G${sumRow+2}:G${sumRow+100}=TRUE)*1),"")`, '',
    'Deposits Paid:','',
    `=IFERROR(SUMPRODUCT((R${sumRow+2}:R${sumRow+100}=TRUE)*1),"")`, '',
    'Total Contracted:','',
    `=IFERROR(SUM(T${sumRow+2}:T${sumRow+100}),0)`, '',
    'Total Paid:','',
    `=IFERROR(SUM(U${sumRow+2}:U${sumRow+100}),0)`, '',
    'Balance Remaining:','',
    `=IFERROR(SUM(V${sumRow+2}:V${sumRow+100}),0)`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(VV, sumRow - 1, sumRow, 0, 24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Headers
  const hdrRow = sumRow + 1;
  data.push({ range: `${S}!A${hdrRow}`, values: [HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(VV, hdrRow - 1, hdrRow, 0, 24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: VV, dimension: 'ROWS', startIndex: hdrRow - 1, endIndex: hdrRow },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // Vendor data rows
  vendors.forEach((v, i) => {
    const row = hdrRow + 1 + i;
    const [cat, company, contact, phone, email, contractSigned, contractDate, service, status, eventDate, startTime, endTime, location, pkg, quoted, deposit, depositPaid, depositDate, total, paid, website, notes] = v;
    data.push({ range: `${S}!A${row}`, values: [[i + 1]] });
    data.push({ range: `${S}!B${row}`, values: [[cat]] });
    data.push({ range: `${S}!C${row}`, values: [[company]] });
    data.push({ range: `${S}!D${row}`, values: [[contact]] });
    data.push({ range: `${S}!E${row}`, values: [[phone]] });
    data.push({ range: `${S}!F${row}`, values: [[email]] });
    data.push({ range: `${S}!G${row}`, values: [[contractSigned]] });
    data.push({ range: `${S}!H${row}`, values: [[contractDate]] });
    data.push({ range: `${S}!I${row}`, values: [[service]] });
    data.push({ range: `${S}!J${row}`, values: [[status]] });
    data.push({ range: `${S}!K${row}`, values: [[eventDate]] });
    data.push({ range: `${S}!L${row}`, values: [[startTime]] });
    data.push({ range: `${S}!M${row}`, values: [[endTime]] });
    data.push({ range: `${S}!N${row}`, values: [[location]] });
    data.push({ range: `${S}!O${row}`, values: [[pkg]] });
    data.push({ range: `${S}!P${row}`, values: [[quoted]] });
    data.push({ range: `${S}!Q${row}`, values: [[deposit]] });
    data.push({ range: `${S}!R${row}`, values: [[depositPaid]] });
    data.push({ range: `${S}!S${row}`, values: [[depositDate]] });
    data.push({ range: `${S}!T${row}`, values: [[total]] });
    data.push({ range: `${S}!U${row}`, values: [[paid]] });
    // V: Balance Due formula
    data.push({ range: `${S}!V${row}`, values: [[`=IFERROR(T${row}-U${row},0)`]] });
    data.push({ range: `${S}!W${row}`, values: [[website]] });
    data.push({ range: `${S}!X${row}`, values: [[notes]] });
  });

  // Checkboxes: G (contract), R (deposit paid)
  const vendorDataStart = hdrRow;
  const vendorDataEnd = vendorDataStart + vendors.length;
  [6, 17].forEach(col => {
    reqs.push({ setDataValidation: {
      range: gridRange(VV, vendorDataStart, vendorDataEnd, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Booking Status dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(VV, vendorDataStart, 200, 9, 10),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Booked', 'Deposit Paid', 'Inquired', 'Quoted', 'Declined', 'Backup Option', 'Cancelled',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Currency format: P, Q, T, U, V
  [15, 16, 19, 20, 21].forEach(col => {
    reqs.push({ repeatCell: { range: gridRange(VV, vendorDataStart, 200, col, col + 1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });

  // V: balance due — formula cell
  reqs.push({ repeatCell: { range: gridRange(VV, vendorDataStart, vendorDataEnd, 21, 22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Alternating rows
  for (let i = 0; i < vendors.length; i++) {
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(VV, vendorDataStart + i, vendorDataStart + i + 1, 0, 24), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, foregroundColor: hex(C.mainText) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // CF: Booking Status
  [{ text: 'Booked',       bg: C.success  },
   { text: 'Deposit Paid', bg: C.success  },
   { text: 'Cancelled',    bg: C.attention },
  ].forEach(({ text, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: { ranges: [gridRange(VV, vendorDataStart, 200, 9, 10)],
        booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bg), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
      }, index: 0,
    }});
  });

  // Column widths
  const widths = [
    35, 130, 160, 120, 100, 160, // A-F
    70, 85, 200, 90,             // G-J
    85, 70, 70, 130, 110,        // K-O
    90, 80, 70, 80,              // P-S
    90, 90, 90, 140, 160,        // T-X
  ];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: VV, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze vendor header row
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: VV, gridProperties: { frozenRowCount: hdrRow } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: VV, tabColor: hex(C.secondary) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'vendors-data');
  await batchUpdate(id, reqs, 'vendors-format');

  console.log('✓ Vendors & Venue built with', vendors.length, 'vendors + venue section');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
