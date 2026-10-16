'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL = sheetMap['Guest List & RSVP'];
const S  = "'Guest List & RSVP'";

// 39 columns A:AM
const HEADERS = [
  '#', 'First Name', 'Last Name', 'Full Name', 'Household ID', 'Household Name',
  'Relationship', 'Side', 'Group', 'Meal Preference', 'Dietary Restrictions',
  'Allergies', 'Accessibility Needs', 'Plus One Allowed', 'Plus One Name',
  'Invite Sent', 'Invite Date', 'RSVP Status', 'RSVP Date',
  'Attending Ceremony', 'Attending Reception', 'Rehearsal Dinner Invite',
  'Rehearsal Dinner RSVP', 'Brunch Invite', 'Brunch RSVP',
  'Table Assignment', 'Seat #', 'Gift Received', 'Gift Description',
  'Thank You Sent', 'Thank You Date', 'Notes', 'Phone', 'Email',
  'Street Address', 'City', 'State', 'Country', 'Postal Code',
];

// Sample data: 64 guests in 38 households
const guests = [
  // Osei Family (Bride's side)
  ['H01','Osei Family',    'Kwame',   'Osei',      "Bride's Family", 'Bride', 'Immediate Family', 'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/8/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H01','Osei Family',    'Abena',   'Osei',      "Bride's Family", 'Bride', 'Immediate Family', 'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/8/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H02','Osei Family',    'Kofi',    'Osei',      "Bride's Family", 'Bride', 'Immediate Family', 'Chicken',   '',          'No',  '',              'No',  'Yes','Priya Osei',  true, '5/1/2026','Attending','5/10/2026', 'Yes','Yes','No', 'No', 'Yes','Yes'],
  ['H02','Osei Family',    'Priya',   'Osei',      "Bride's Family", 'Bride', 'Extended Family',  'Vegetarian','Nut Allergy','No', '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/10/2026', 'Yes','Yes','No', 'No', 'Yes','Yes'],
  ['H03','Mensah Family',  'Akosua',  'Mensah',    "Bride's Family", 'Bride', 'Extended Family',  'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/15/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H03','Mensah Family',  'Yaw',     'Mensah',    "Bride's Family", 'Bride', 'Extended Family',  'Beef',       '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/15/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H04','Boateng Family', 'Ama',     'Boateng',   "Bride's Family", 'Bride', 'Extended Family',  'Chicken',    '',          'No',  'Wheelchair',   'Yes', 'Yes','',            true, '5/3/2026','Attending','5/12/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H04','Boateng Family', 'Nana',    'Boateng',   "Bride's Family", 'Bride', 'Extended Family',  'Vegetarian','',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/12/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H05','Asante Family',  'Efua',    'Asante',    "Bride's Family", 'Bride', 'Extended Family',  'Fish',       'Gluten-Free','No','',              'No',  'No', '',            true, '5/3/2026','Not Attending','5/20/2026','No','No', 'No', 'No', 'No', 'No'],
  ['H06','Acheampong',     'Serwa',   'Acheampong',"Bride's Family", 'Bride', 'Extended Family',  'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/18/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  // Thornton Family (Groom's side)
  ['H07','Thornton Family','Robert',  'Thornton',  "Groom's Family", 'Groom', 'Immediate Family', 'Beef',       '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/5/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H07','Thornton Family','Patricia','Thornton',  "Groom's Family", 'Groom', 'Immediate Family', 'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/5/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H08','Thornton Sibs',  'Jason',   'Thornton',  "Groom's Family", 'Groom', 'Immediate Family', 'Beef',       '',          'No',  '',              'No',  'Yes','Mei Thornton',true, '5/1/2026','Attending','5/7/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H08','Thornton Sibs',  'Mei',     'Thornton',  "Groom's Family", 'Groom', 'Extended Family',  'Vegetarian','Dairy-Free','No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/7/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H09','Thornton Sibs',  'Kira',    'Thornton',  "Groom's Family", 'Groom', 'Immediate Family', 'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/9/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H09','Thornton Sibs',  'Dean',    'Walsh',     "Groom's Family", 'Groom', 'Extended Family',  'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/1/2026','Attending','5/9/2026',  'Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H10','Thornton Uncle', 'Charles', 'Thornton',  "Groom's Family", 'Groom', 'Extended Family',  'Beef',       '',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/14/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H10','Thornton Uncle', 'Helen',   'Thornton',  "Groom's Family", 'Groom', 'Extended Family',  'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/14/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H11','Barnes Family',  'Michael', 'Barnes',    "Groom's Family", 'Groom', 'Extended Family',  'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/16/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H11','Barnes Family',  'Susan',   'Barnes',    "Groom's Family", 'Groom', 'Extended Family',  'Vegetarian','',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/16/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H12','Thornton Cousins','Tyler',  'Thornton',  "Groom's Family", 'Groom', 'Extended Family',  'Beef',       '',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Pending',  '',          'No', 'No', 'No', 'No', 'No', 'No'],
  ['H13','Ellis Family',   'Sandra',  'Ellis',     "Groom's Family", 'Groom', 'Extended Family',  'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/19/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H13','Ellis Family',   'Ben',     'Ellis',     "Groom's Family", 'Groom', 'Extended Family',  'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/3/2026','Attending','5/19/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  // Wedding Party
  ['H14','Jordan Osei',    'Jordan',  'Osei',      "Bride's Friend", 'Bride', 'Wedding Party',    'Fish',       '',          'No',  '',              'No',  'Yes','Lena Park',   true, '4/15/2026','Attending','4/18/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H14','Jordan Osei',    'Lena',    'Park',      "Bride's Friend", 'Both',  'Friends',          'Vegetarian','Soy-Free',  'No',  '',              'No',  'Yes','',            true, '4/15/2026','Attending','4/18/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H15','Rivera Daniela', 'Daniela', 'Rivera',    "Bride's Friend", 'Bride', 'Wedding Party',    'Chicken',    '',          'No',  '',              'No',  'Yes','Carlos Rivera',true,'4/15/2026','Attending','4/17/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H15','Rivera Daniela', 'Carlos',  'Rivera',    "Bride's Friend", 'Both',  'Friends',          'Beef',       '',          'No',  '',              'No',  'Yes','',            true, '4/15/2026','Attending','4/17/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H16','Chen Sophie',    'Sophie',  'Chen',      "Bride's Friend", 'Bride', 'Wedding Party',    'Fish',       'Shellfish Allergy','No','',         'No',  'Yes','',            true, '4/15/2026','Attending','4/20/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H17','Davis Tristan',  'Tristan', 'Davis',     "Groom's Friend", 'Groom', 'Wedding Party',    'Beef',       '',          'No',  '',              'No',  'Yes','Aria Davis',  true, '4/15/2026','Attending','4/19/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H17','Davis Tristan',  'Aria',    'Davis',     "Groom's Friend", 'Both',  'Friends',          'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '4/15/2026','Attending','4/19/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  ['H18','Parker Noel',    'Noel',    'Parker',    "Groom's Friend", 'Groom', 'Wedding Party',    'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '4/15/2026','Attending','4/22/2026','Yes','Yes','Yes','Yes','Yes','Yes'],
  // Friends (Mutual)
  ['H19','Kim Family',     'David',   'Kim',       'Mutual Friend',  'Both',  'Friends',          'Fish',       '',          'No',  '',              'No',  'Yes','Julia Kim',   true, '5/5/2026', 'Attending','5/12/2026', 'No', 'Yes','No', 'No', 'Yes','No'],
  ['H19','Kim Family',     'Julia',   'Kim',       'Mutual Friend',  'Both',  'Friends',          'Vegetarian','Nut Allergy','No','',               'No',  'Yes','',            true, '5/5/2026', 'Attending','5/12/2026', 'No', 'Yes','No', 'No', 'Yes','No'],
  ['H20','Patel Family',   'Rohan',   'Patel',     'Mutual Friend',  'Both',  'Friends',          'Vegetarian','',          'No',  '',              'No',  'Yes','Nisha Patel',  true,'5/5/2026', 'Attending','5/11/2026', 'No', 'Yes','No', 'No', 'Yes','No'],
  ['H20','Patel Family',   'Nisha',   'Patel',     'Mutual Friend',  'Both',  'Friends',          'Vegetarian','Dairy-Free','No', '',              'No',  'Yes','',            true, '5/5/2026', 'Attending','5/11/2026', 'No', 'Yes','No', 'No', 'Yes','No'],
  ['H21','Nguyen Family',  'Thanh',   'Nguyen',    'Mutual Friend',  'Both',  'Friends',          'Chicken',    '',          'No',  '',              'No',  'Yes','Linh Nguyen',  true,'5/5/2026', 'Attending','5/13/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H21','Nguyen Family',  'Linh',    'Nguyen',    'Mutual Friend',  'Both',  'Friends',          'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/5/2026', 'Attending','5/13/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H22','Torres Family',  'Elena',   'Torres',    "Bride's Friend", 'Bride', 'Friends',          'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/5/2026', 'Attending','5/14/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H23','Morgan Family',  'Sam',     'Morgan',    "Groom's Friend", 'Groom', 'Friends',          'Beef',       '',          'No',  '',              'No',  'Yes','Alex Morgan',  true,'5/5/2026', 'Attending','5/17/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H23','Morgan Family',  'Alex',    'Morgan',    "Groom's Friend", 'Both',  'Friends',          'Vegetarian','',          'No',  '',              'No',  'Yes','',            true, '5/5/2026', 'Attending','5/17/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H24','Reyes Family',   'Marco',   'Reyes',     "Bride's Friend", 'Bride', 'Friends',          'Beef',       '',          'No',  '',              'No',  'Yes','Lucia Reyes',  true,'5/5/2026', 'Attending','5/19/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H24','Reyes Family',   'Lucia',   'Reyes',     "Bride's Friend", 'Both',  'Friends',          'Fish',       'Gluten-Free','No','',              'No',  'Yes','',            true, '5/5/2026', 'Attending','5/19/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H25','Walker Family',  'James',   'Walker',    "Groom's Friend", 'Groom', 'Friends',          'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/5/2026', 'Pending',  '',          'No', 'No', 'No', 'No', 'No', 'No'],
  ['H26','Foster Family',  'Grace',   'Foster',    'Mutual Friend',  'Both',  'Friends',          'Vegan',      '',          'No',  '',              'No',  'Yes','Tyler Foster',  true,'5/5/2026', 'Attending','5/20/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H26','Foster Family',  'Tyler',   'Foster',    'Mutual Friend',  'Both',  'Friends',          'Beef',       '',          'No',  '',              'No',  'Yes','',            true, '5/5/2026', 'Attending','5/20/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  // Colleagues
  ['H27','Dr. Amara Work', 'Priscilla','Okafor',   'Colleague',      'Bride', 'Colleagues',       'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/15/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H28','Marcus Work',    'Brandon', 'Hayes',     'Colleague',      'Groom', 'Colleagues',       'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/16/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H29','Marcus Work',    'Simone',  'Laurent',   'Colleague',      'Groom', 'Colleagues',       'Vegetarian','Dairy-Free','No','',                'No',  'Yes','',            true, '5/7/2026', 'Not Attending','5/22/2026','No','No','No','No','No','No'],
  ['H30','Amara Work',     'Fatima',  'Al-Hassan', 'Colleague',      'Bride', 'Colleagues',       'Chicken',    'Halal',     'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/18/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H31','Marcus Work',    'Kevin',   'O\'Brien',  'Colleague',      'Groom', 'Colleagues',       'Beef',       '',          'No',  '',              'No',  'Yes','Sarah O\'Brien',true,'5/7/2026','Attending','5/20/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H31','Marcus Work',    'Sarah',   'O\'Brien',  'Colleague',      'Both',  'Friends',          'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/20/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  // Children
  ['H03','Mensah Family',  'Kwame Jr','Mensah',    "Bride's Family", 'Bride', 'Children',         'Kids Meal',  '',          'No',  '',              'No',  'No', '',            true, '5/1/2026', 'Attending','5/15/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H04','Boateng Family', 'Yaa',     'Boateng',   "Bride's Family", 'Bride', 'Children',         'Kids Meal',  '',          'No',  '',              'No',  'No', '',            true, '5/3/2026', 'Attending','5/12/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H08','Thornton Sibs',  'Oliver',  'Thornton',  "Groom's Family", 'Groom', 'Children',         'Kids Meal',  '',          'No',  '',              'No',  'No', '',            true, '5/1/2026', 'Attending','5/7/2026',  'No', 'Yes','No', 'No', 'No', 'No'],
  ['H19','Kim Family',     'Ella',    'Kim',       'Mutual Friend',  'Both',  'Children',         'Kids Meal',  '',          'No',  '',              'No',  'No', '',            true, '5/5/2026', 'Attending','5/12/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  // Extra guests
  ['H32','Abara Family',   'Emeka',   'Abara',     "Bride's Family", 'Bride', 'Extended Family',  'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/22/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H32','Abara Family',   'Chioma',  'Abara',     "Bride's Family", 'Bride', 'Extended Family',  'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/22/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H33','Greene Family',  'Marcus',  'Greene',    "Groom's Friend", 'Groom', 'Friends',          'Beef',       '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Pending',  '',          'No', 'No', 'No', 'No', 'No', 'No'],
  ['H34','Huang Family',   'Wei',     'Huang',     'Mutual Friend',  'Both',  'Friends',          'Vegan',      'Gluten-Free','No','',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/21/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H35','Adeyemi Family', 'Tunde',   'Adeyemi',   "Bride's Family", 'Bride', 'Extended Family',  'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/23/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H35','Adeyemi Family', 'Bimpe',   'Adeyemi',   "Bride's Family", 'Bride', 'Extended Family',  'Fish',       'Nut Allergy','No','',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/23/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H36','Shaw Family',    'Nathan',  'Shaw',      "Groom's Friend", 'Groom', 'Friends',          'Beef',       '',          'No',  '',              'No',  'Yes','Chloe Shaw',   true,'5/7/2026', 'Attending','5/24/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H36','Shaw Family',    'Chloe',   'Shaw',      "Groom's Friend", 'Both',  'Friends',          'Vegetarian','',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/24/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
  ['H37','Owusu Family',   'Esi',     'Owusu',     "Bride's Family", 'Bride', 'Extended Family',  'Chicken',    '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'No Response','',       'No', 'No', 'No', 'No', 'No', 'No'],
  ['H38','Bell Family',    'Marcus',  'Bell',      "Groom's Friend", 'Groom', 'Colleagues',       'Fish',       '',          'No',  '',              'No',  'Yes','',            true, '5/7/2026', 'Attending','5/25/2026', 'No', 'Yes','No', 'No', 'No', 'No'],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['GUEST LIST & RSVP TRACKER']] });
  reqs.push({ mergeCells: { range: gridRange(GL, 0, 1, 0, 39), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(GL, 0, 1, 0, 39), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Summary row 2 ─────────────────────────────────────────────────────────
  const sumFormulas = [
    `=IFERROR(COUNTA(B6:B600),"")`,                                                         // Total guests
    `=IFERROR(SUMPRODUCT((R6:R600="Attending")*1),"")`,                                     // Attending
    `=IFERROR(SUMPRODUCT((R6:R600="Not Attending")*1),"")`,                                 // Not attending
    `=IFERROR(SUMPRODUCT((R6:R600="Pending")*1)+(SUMPRODUCT((R6:R600="No Response")*1)),"")`,// Pending
    `=IFERROR(SUMPRODUCT((R6:R600="Attending")*1)/COUNTA(B6:B600),"")`,                     // RSVP rate
  ];
  const sumLabels = ['Total Guests:', '', 'Attending:', '', 'Not Attending:', '', 'Awaiting Response:', '', 'RSVP Response Rate:'];
  data.push({ range: `${S}!A2`, values: [sumLabels] });
  data.push({ range: `${S}!B2`, values: [[sumFormulas[0]]] });
  data.push({ range: `${S}!D2`, values: [[sumFormulas[1]]] });
  data.push({ range: `${S}!F2`, values: [[sumFormulas[2]]] });
  data.push({ range: `${S}!H2`, values: [[sumFormulas[3]]] });
  data.push({ range: `${S}!J2`, values: [[sumFormulas[4]]] });
  reqs.push({ repeatCell: { range: gridRange(GL, 1, 2, 0, 39), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  // Format J2 as percentage
  reqs.push({ repeatCell: { range: gridRange(GL, 1, 2, 9, 10), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // ── Column Headers row 5 ──────────────────────────────────────────────────
  data.push({ range: `${S}!A5`, values: [HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(GL, 4, 5, 0, 39), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER',
    wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ── Guest Data rows 6+ ───────────────────────────────────────────────────
  guests.forEach((g, i) => {
    const row = 6 + i;
    const [hid, hname, first, last, rel, side, group, meal, diet, a11y, a11yNote, plusAllowed, plusName, invSent, invDate, rsvpStatus, rsvpDate, attCer, attRec, rehInv, rehRsvp, brInv, brRsvp] = g;

    // Auto-number
    data.push({ range: `${S}!A${row}`, values: [[i + 1]] });
    // Names
    data.push({ range: `${S}!B${row}`, values: [[first]] });
    data.push({ range: `${S}!C${row}`, values: [[last]] });
    data.push({ range: `${S}!D${row}`, values: [[`=IFERROR(B${row}&" "&C${row},"")`]] });
    // Household
    data.push({ range: `${S}!E${row}`, values: [[hid]] });
    data.push({ range: `${S}!F${row}`, values: [[hname]] });
    // Categories
    data.push({ range: `${S}!G${row}`, values: [[rel]] });
    data.push({ range: `${S}!H${row}`, values: [[side]] });
    data.push({ range: `${S}!I${row}`, values: [[group]] });
    // Meal / Diet
    data.push({ range: `${S}!J${row}`, values: [[meal]] });
    data.push({ range: `${S}!K${row}`, values: [[diet]] });
    // Accessibility
    data.push({ range: `${S}!M${row}`, values: [[a11yNote]] });
    // Plus One
    data.push({ range: `${S}!N${row}`, values: [[plusAllowed]] });
    data.push({ range: `${S}!O${row}`, values: [[plusName]] });
    // Invite
    data.push({ range: `${S}!P${row}`, values: [[invSent]] });
    data.push({ range: `${S}!Q${row}`, values: [[invDate]] });
    // RSVP
    data.push({ range: `${S}!R${row}`, values: [[rsvpStatus]] });
    data.push({ range: `${S}!S${row}`, values: [[rsvpDate]] });
    // Event attendance
    data.push({ range: `${S}!T${row}`, values: [[attCer]] });
    data.push({ range: `${S}!U${row}`, values: [[attRec]] });
    data.push({ range: `${S}!V${row}`, values: [[rehInv]] });
    data.push({ range: `${S}!W${row}`, values: [[rehRsvp]] });
    data.push({ range: `${S}!X${row}`, values: [[brInv]] });
    data.push({ range: `${S}!Y${row}`, values: [[brRsvp]] });
    // Table / Seat — left blank for Seating Chart to fill
    // Gift
    data.push({ range: `${S}!AB${row}`, values: [[false]] });
    data.push({ range: `${S}!AD${row}`, values: [[false]] });
  });

  // Checkboxes: P (Invite Sent), AB (Gift), AD (Thank You)
  const cbCols = [15, 27, 29]; // 0-indexed: P=15, AB=27, AD=29
  cbCols.forEach(col => {
    reqs.push({ setDataValidation: {
      range: gridRange(GL, 5, 5 + guests.length, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Data validation dropdowns
  const dvList = (col, values) => ({
    setDataValidation: {
      range: gridRange(GL, 5, 200, col, col + 1),
      rule: { condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
    },
  });
  reqs.push(dvList(7, ['Bride', 'Groom', 'Both']));                                                              // H: Side
  reqs.push(dvList(8, ['Immediate Family', 'Extended Family', 'Wedding Party', 'Friends', 'Colleagues', 'Neighbors', 'Children', 'Other'])); // I: Group
  reqs.push(dvList(9, ['Chicken', 'Fish', 'Beef', 'Vegetarian', 'Vegan', 'Kids Meal', 'TBD']));                  // J: Meal
  reqs.push(dvList(13, ['Yes', 'No']));                                                                           // N: Plus One Allowed
  reqs.push(dvList(17, ['Attending', 'Not Attending', 'Pending', 'No Response']));                                // R: RSVP Status
  reqs.push(dvList(19, ['Yes', 'No', 'TBD']));                                                                    // T: Ceremony
  reqs.push(dvList(20, ['Yes', 'No', 'TBD']));                                                                    // U: Reception

  // Alternating row colors
  for (let i = 0; i < guests.length; i++) {
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(GL, 5 + i, 6 + i, 0, 39), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { fontSize: 9, foregroundColor: hex(C.mainText) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // Conditional format: RSVP Status color coding (col R = index 17)
  const cfRules = [
    { values: ['Attending'],     bg: C.success },
    { values: ['Not Attending'], bg: C.attention },
    { values: ['Pending'],       bg: C.warning },
    { values: ['No Response'],   bg: C.neutral },
  ];
  cfRules.forEach(({ values, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(GL, 5, 200, 17, 18)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: values[0] }] },
          format: { backgroundColor: hex(bg), textFormat: { bold: true, foregroundColor: hex(C.white) } },
        },
      },
      index: 0,
    }});
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  const widths = [
    40, 80, 80, 130, 55, 130, 120, 55, 110,  // A-I
    90, 110, 100, 120, 75, 110,               // J-O
    60, 75, 90, 75,                           // P-S
    75, 75, 70, 70, 70, 70,                   // T-Y
    80, 45, 60, 120, 60, 75, 130,             // Z-AF
    90, 140, 140, 70, 70, 70, 80,             // AG-AM
  ];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze rows 1-5, columns A-E
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: GL, gridProperties: { frozenRowCount: 5 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: GL, tabColor: hex(C.success) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'guests-data');
  await batchUpdate(id, reqs, 'guests-format');

  console.log('✓ Guest List & RSVP built with', guests.length, 'guests');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
