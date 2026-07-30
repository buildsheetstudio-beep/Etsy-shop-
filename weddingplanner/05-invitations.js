'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IT = sheetMap['Invitation Tracker'];
const S  = "'Invitation Tracker'";

const HEADERS = [
  '#', 'Household ID', 'Household Name', 'Primary Contact', 'Address Line 1', 'Address Line 2',
  'City', 'State/Province', 'Postal Code', 'Country',
  'Envelope Style', 'Guest Count', 'Invitation Printed', 'Print Date',
  'Envelope Addressed', 'Mailed Date', 'Delivery Status', 'Follow-Up Needed',
  'Follow-Up Date', 'Notes',
];

// One record per household — 38 households
const invitations = [
  ['H01','Osei Family',    'Kwame Osei',        '240 Riverside Drive',   '',   'Atlanta',    'GA','30301','USA','Single',    2, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H02','Osei Family',    'Kofi Osei',         '142 Maple Street',      '',   'Austin',     'TX','78701','USA','Single',    2, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H03','Mensah Family',  'Akosua Mensah',     '85 Oak Avenue',         '',   'Houston',    'TX','77002','USA','Single',    3, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H04','Boateng Family', 'Ama Boateng',       '310 Cedar Court',       '',   'Dallas',     'TX','75201','USA','Single',    3, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,'Wheelchair access needed at venue'],
  ['H05','Asante Family',  'Efua Asante',       '77 Birch Lane',         '',   'Nashville',  'TN','37201','USA','Single',    1, true,'4/28/2026',true,'5/1/2026', 'Returned Undelivered',true,'5/10/2026'],
  ['H06','Acheampong',     'Serwa Acheampong',  '502 Peachtree Blvd',    '',   'Atlanta',    'GA','30302','USA','Single',    1, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H07','Thornton Family','Robert Thornton',   '1822 Oak Park Drive',   '',   'San Antonio','TX','78201','USA','Outer + Inner',2,true,'4/28/2026',true,'5/1/2026','Mailed',false,''],
  ['H08','Thornton Sibs',  'Jason Thornton',    '934 Willow Way',        '',   'Austin',     'TX','78702','USA','Single',    3, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H09','Thornton Sibs',  'Kira Walsh',        '211 Elm Street',        'Apt 4B','Chicago', 'IL','60601','USA','Single',    2, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H10','Thornton Uncle', 'Charles Thornton',  '65 Heritage Road',      '',   'Austin',     'TX','78703','USA','Single',    2, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H11','Barnes Family',  'Michael Barnes',    '47 Summit View',        '',   'Fort Worth', 'TX','76101','USA','Single',    2, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H12','Thornton Cousins','Tyler Thornton',   '13 Lakeview Court',     '',   'Denver',     'CO','80201','USA','Single',    1, true,'4/28/2026',true,'5/1/2026', 'Mailed',true,'5/25/2026 — No RSVP'],
  ['H13','Ellis Family',   'Sandra Ellis',      '789 Pinecrest Avenue',  '',   'Austin',     'TX','78704','USA','Single',    2, true,'4/28/2026',true,'5/1/2026', 'Mailed',false,''],
  ['H14','Jordan Osei',    'Jordan Osei',       '504 Violet Street',     'Unit 2','Austin',  'TX','78705','USA','Outer + Inner',2,true,'4/12/2026',true,'4/15/2026','Hand Delivered',false,'Maid of Honor'],
  ['H15','Rivera Daniela', 'Daniela Rivera',    '118 Sunflower Trail',   '',   'Austin',     'TX','78706','USA','Outer + Inner',2,true,'4/12/2026',true,'4/15/2026','Hand Delivered',false,'Bridesmaid'],
  ['H16','Chen Sophie',    'Sophie Chen',       '622 Rosewood Lane',     '',   'San Marcos', 'TX','78666','USA','Outer + Inner',1,true,'4/12/2026',true,'4/15/2026','Hand Delivered',false,'Bridesmaid'],
  ['H17','Davis Tristan',  'Tristan Davis',     '300 Bluebell Drive',    '',   'Austin',     'TX','78707','USA','Outer + Inner',2,true,'4/12/2026',true,'4/15/2026','Hand Delivered',false,'Groomsman'],
  ['H18','Parker Noel',    'Noel Parker',       '158 Canyon Rim Road',   '',   'Austin',     'TX','78708','USA','Outer + Inner',1,true,'4/12/2026',true,'4/15/2026','Hand Delivered',false,'Best Man'],
  ['H19','Kim Family',     'David Kim',         '234 Harvest Moon Blvd', '',   'Round Rock',  'TX','78664','USA','Single',   3, true,'4/28/2026',true,'5/5/2026', 'Mailed',false,''],
  ['H20','Patel Family',   'Rohan Patel',       '71 Spice Market Way',   '',   'Cedar Park',  'TX','78613','USA','Single',   2, true,'4/28/2026',true,'5/5/2026', 'Mailed',false,''],
  ['H21','Nguyen Family',  'Thanh Nguyen',      '409 Lantern Court',     '',   'Pflugerville','TX','78660','USA','Single',   2, true,'4/28/2026',true,'5/5/2026', 'Mailed',false,''],
  ['H22','Torres Family',  'Elena Torres',      '55 Calico Road',        '',   'Kyle',        'TX','78640','USA','Single',   1, true,'4/28/2026',true,'5/5/2026', 'Mailed',false,''],
  ['H23','Morgan Family',  'Sam Morgan',        '820 Mockingbird Lane',  '',   'Austin',      'TX','78709','USA','Single',   2, true,'4/28/2026',true,'5/5/2026', 'Mailed',false,''],
  ['H24','Reyes Family',   'Marco Reyes',       '133 Desert Bloom Court','',   'Buda',        'TX','78610','USA','Single',   2, true,'4/28/2026',true,'5/5/2026', 'Mailed',false,''],
  ['H25','Walker Family',  'James Walker',      '44 Copperfield Drive',  '',   'Austin',      'TX','78710','USA','Single',   1, true,'4/28/2026',true,'5/5/2026', 'Mailed',true,'5/28/2026 — No RSVP'],
  ['H26','Foster Family',  'Grace Foster',      '601 Whitewater Lane',   '',   'Bastrop',     'TX','78602','USA','Single',   2, true,'4/28/2026',true,'5/5/2026', 'Mailed',false,''],
  ['H27','Dr. Amara Work', 'Priscilla Okafor',  '777 Regency Park',      '',   'Austin',      'TX','78711','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H28','Marcus Work',    'Brandon Hayes',     '1040 Sterling Oak Way', '',   'Austin',      'TX','78712','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H29','Marcus Work',    'Simone Laurent',    '58 Rue de la Paix',     '',   'Austin',      'TX','78713','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,'RSVPed No'],
  ['H30','Amara Work',     'Fatima Al-Hassan',  '321 Desert Palm',       '',   'Austin',      'TX','78714','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H31','Marcus Work',    "Kevin O'Brien",     '490 Ridgeview Terrace', '',   'Georgetown',  'TX','78626','USA','Single',   2, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H32','Abara Family',   'Emeka Abara',       '28 Kingsway Court',     '',   'Houston',     'TX','77003','USA','Single',   2, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H33','Greene Family',  'Marcus Greene',     '505 Foxglove Trail',    '',   'Austin',      'TX','78715','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',true,'5/30/2026 — No RSVP'],
  ['H34','Huang Family',   'Wei Huang',         '15 Bamboo Gardens',     '',   'Austin',      'TX','78716','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H35','Adeyemi Family', 'Tunde Adeyemi',     '702 Sunset Ridge',      '',   'Austin',      'TX','78717','USA','Single',   2, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H36','Shaw Family',    'Nathan Shaw',       '815 Creekside Path',    '',   'Austin',      'TX','78718','USA','Single',   2, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
  ['H37','Owusu Family',   'Esi Owusu',         '1100 Heritage Oak',     '',   'Austin',      'TX','78719','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',true,'6/1/2026 — No RSVP yet'],
  ['H38','Bell Family',    'Marcus Bell',       '27 Clearwater Cove',    '',   'Leander',     'TX','78641','USA','Single',   1, true,'4/28/2026',true,'5/7/2026', 'Mailed',false,''],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['INVITATION TRACKER']] });
  reqs.push({ mergeCells: { range: gridRange(IT, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(IT, 0, 1, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: IT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Summary row 2 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [['Total Households:','','',
    `=IFERROR(COUNTA(B5:B200),"")`, '',
    'Mailed:','',
    `=IFERROR(SUMPRODUCT((Q5:Q200="Mailed")*1),"")`, '',
    'Follow-Up Needed:','',
    `=IFERROR(SUMPRODUCT((R5:R200=TRUE)*1),"")`, '',
    'Printed:','',
    `=IFERROR(SUMPRODUCT((M5:M200=TRUE)*1),"")`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(IT, 1, 2, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Headers row 4 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A4`, values: [HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(IT, 3, 4, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER',
    wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: IT, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ── Data rows ─────────────────────────────────────────────────────────────
  invitations.forEach((inv, i) => {
    const row = 5 + i;
    const [hid, hname, contact, addr1, addr2, city, state, zip, country, style, guestCount, printed, printDate, addressed, mailedDate, delivStatus, followUp, followUpDate, notes] = inv;
    data.push({ range: `${S}!A${row}`, values: [[i + 1]] });
    data.push({ range: `${S}!B${row}`, values: [[hid]] });
    data.push({ range: `${S}!C${row}`, values: [[hname]] });
    data.push({ range: `${S}!D${row}`, values: [[contact]] });
    data.push({ range: `${S}!E${row}`, values: [[addr1]] });
    data.push({ range: `${S}!F${row}`, values: [[addr2]] });
    data.push({ range: `${S}!G${row}`, values: [[city]] });
    data.push({ range: `${S}!H${row}`, values: [[state]] });
    data.push({ range: `${S}!I${row}`, values: [[zip]] });
    data.push({ range: `${S}!J${row}`, values: [[country]] });
    data.push({ range: `${S}!K${row}`, values: [[style]] });
    data.push({ range: `${S}!L${row}`, values: [[guestCount]] });
    data.push({ range: `${S}!M${row}`, values: [[printed]] });
    data.push({ range: `${S}!N${row}`, values: [[printDate]] });
    data.push({ range: `${S}!O${row}`, values: [[addressed]] });
    data.push({ range: `${S}!P${row}`, values: [[mailedDate]] });
    data.push({ range: `${S}!Q${row}`, values: [[delivStatus]] });
    data.push({ range: `${S}!R${row}`, values: [[followUp]] });
    data.push({ range: `${S}!S${row}`, values: [[followUpDate]] });
    data.push({ range: `${S}!T${row}`, values: [[notes]] });
  });

  // Checkboxes: M (Printed), O (Addressed), R (Follow-Up Needed)
  [12, 14, 17].forEach(col => {
    reqs.push({ setDataValidation: {
      range: gridRange(IT, 4, 4 + invitations.length, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Delivery Status dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(IT, 4, 200, 16, 17),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Mailed', 'Hand Delivered', 'Returned Undelivered', 'Lost', 'Not Yet Sent',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Envelope style dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(IT, 4, 200, 10, 11),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Single', 'Outer + Inner',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Alternating row colors
  for (let i = 0; i < invitations.length; i++) {
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(IT, 4 + i, 5 + i, 0, 20), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, foregroundColor: hex(C.mainText) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // Conditional format: Delivery Status
  const cfRules = [
    { text: 'Mailed',                  bg: C.success },
    { text: 'Hand Delivered',          bg: C.success },
    { text: 'Returned Undelivered',    bg: C.attention },
    { text: 'Lost',                    bg: C.attention },
  ];
  cfRules.forEach(({ text, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(IT, 4, 200, 16, 17)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bg), textFormat: { bold: true, foregroundColor: hex(C.white) } },
        },
      },
      index: 0,
    }});
  });

  // Conditional format: Follow-Up Needed = TRUE → highlight
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(IT, 4, 200, 17, 18)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'TRUE' }] },
        format: { backgroundColor: hex(C.warning), textFormat: { bold: true } },
      },
    },
    index: 0,
  }});

  // Column widths
  const widths = [35, 55, 130, 140, 160, 80, 80, 55, 75, 55, 80, 55, 70, 80, 80, 80, 110, 80, 90, 160];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: IT, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze row 4
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: IT, gridProperties: { frozenRowCount: 4 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: IT, tabColor: hex(C.accent) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'invitations-data');
  await batchUpdate(id, reqs, 'invitations-format');

  console.log('✓ Invitation Tracker built with', invitations.length, 'households');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
