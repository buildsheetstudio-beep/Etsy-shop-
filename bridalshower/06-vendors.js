'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const VEN = sheetMap['💍 Venue, Vendors & Bridesmaids'];

const VENDORS = [
  { label: 'Venue Name',     vals: ['La Maison Garden Room','','',''] },
  { label: 'Category',       vals: ['Venue','','',''] },
  { label: 'Contact Name',   vals: ['Isabelle Moreau','','',''] },
  { label: 'Phone',          vals: ['+44 7890 123456','','',''] },
  { label: 'Email',          vals: ['events@lamaisongarden.co.uk','','',''] },
  { label: 'Website',        vals: ['lamaisongarden.co.uk','','',''] },
  { label: 'Quote ($)',      vals: [600,420,220,''] },
  { label: 'Deposit Paid ($)',vals: [300,0,0,''] },
  { label: 'Balance Due ($)', vals: [300,420,220,''] },
  { label: 'Contract Signed?',vals: ['Yes','No','No',''] },
  { label: 'Rating',         vals: ['5 — Excellent','4 — Very Good','',''] },
  { label: 'Notes',          vals: ['Booked & confirmed','Awaiting quote confirmation','Quote received',''] },
];

const V2_NAMES   = ['La Maison Garden Room','Garden & Co Florists','Sweet Blooms Bakery',''];
const V2_CATS    = ['Venue','Florist','Baker',''];

// Second vendor cols pre-fill
const VENDOR2 = {
  'Venue Name':     'Garden & Co Florists',
  'Category':       'Florist',
  'Contact Name':   'Priya Nair',
  'Phone':          '+44 7812 345678',
  'Email':          'hello@gardenandco.co.uk',
  'Quote ($)':      420,
};
const VENDOR3 = {
  'Venue Name':     'Sweet Blooms Bakery',
  'Category':       'Baker',
  'Contact Name':   'Hannah White',
  'Phone':          '+44 7723 456789',
  'Email':          'orders@sweetblooms.co.uk',
  'Quote ($)':      220,
};

// Bridesmaid tasks
const BM_TASKS = [
  ['Sophie Clarke',   'MOH',        'Book venue','Planning','High','Done'],
  ['Sophie Clarke',   'MOH',        'Coordinate bridesmaids\' outfits','Coordination','High','Done'],
  ['Sophie Clarke',   'MOH',        'Plan game activities','Planning','High','In Progress'],
  ['Lily Thompson',   'Bridesmaid', 'Design & print invitations','Shopping','Medium','Done'],
  ['Lily Thompson',   'Bridesmaid', 'Organise afternoon tea menu','Logistics','Medium','In Progress'],
  ['Emma Grant',      'Bridesmaid', 'Source decorations & florals','Shopping','Medium','In Progress'],
  ['Emma Grant',      'Bridesmaid', 'Set up venue on the day','Day-Of','High','Not Started'],
  ['Sophie Clarke',   'MOH',        'Arrange photographer','Planning','High','Done'],
  ['Lily Thompson',   'Bridesmaid', 'Create seating plan','Logistics','Medium','Not Started'],
  ['Emma Grant',      'Bridesmaid', 'Buy game prizes','Shopping','Low','Not Started'],
  ['Sophie Clarke',   'MOH',        'Write speech','Planning','High','In Progress'],
  ['Lily Thompson',   'Bridesmaid', 'Collect RSVPs','Communication','Medium','Done'],
  ['Emma Grant',      'Bridesmaid', 'Buy party favours','Shopping','Low','Not Started'],
  ['Sophie Clarke',   'MOH',        'Confirm final numbers with caterer','Communication','High','Not Started'],
  ['Lily Thompson',   'Bridesmaid', 'Morning-of coordination text','Day-Of','Medium','Not Started'],
  ['Emma Grant',      'Bridesmaid', 'Pack down venue after party','Day-Of','Medium','Not Started'],
  ['Sophie Clarke',   'MOH',        'Write thank-you cards for bride','Planning','Low','Not Started'],
  ['Lily Thompson',   'Bridesmaid', 'Collect & store gifts after party','Day-Of','Medium','Not Started'],
];

// Supplier contacts
const SUPPLIERS = [
  ['La Maison Garden Room','Venue','Isabelle Moreau','+44 7890 123456','events@lamaisongarden.co.uk','600','300','Deposit paid, confirmed'],
  ['Garden & Co Florists','Florist','Priya Nair','+44 7812 345678','hello@gardenandco.co.uk','420','0','Quote received'],
  ['Sweet Blooms Bakery','Baker','Hannah White','+44 7723 456789','orders@sweetblooms.co.uk','220','0','Quote received'],
  ['Events by Clara','Caterer','Clara Rossi','+44 7734 567890','clara@eventsbyClara.co.uk','650','325','Deposit paid'],
  ['SnapMoments Photo','Photographer','James Lee','+44 7845 678901','james@snapmoments.co.uk','400','400','Fully paid'],
  ['The Styling Suite','Stylist','Mia Chen','+44 7956 789012','hello@stylingsuite.co.uk','','','Consulted only'],
  ['Coco & Bloom Print','Stationery','','',  'hello@cocobloom.co.uk','100','100','Paid in full'],
  ['Harps & Strings Ent.','Entertainment','Ava Singh','+44 7867 890123','ava@harpsandstrings.co.uk','100','0','Enquiry sent'],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'💍 Venue, Vendors & Bridesmaids'";

  // Column widths
  const colW = [160,160,10,160,10,160,10,160,10,160];
  colW.forEach((w,i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: VEN, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Row 0: title
  reqs.push({ mergeCells: { range: gridRange(VEN,0,1,1,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(VEN,0,1,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: VEN, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['💍 Venue, Vendors & Bridesmaids']] });

  // Row 1: vendor comparison section header (0-indexed 1)
  reqs.push({ mergeCells: { range: gridRange(VEN,1,2,0,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(VEN,1,2,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['VENDOR COMPARISON']] });

  // Row 2: vendor column name headers (0-indexed 2)
  // Cols: A=label, B=Vendor1, C=spacer, D=Vendor2, E=spacer, F=Vendor3, G=spacer, H=Vendor4
  const vendorColHeaders = ['FIELD','VENDOR 1','','VENDOR 2','','VENDOR 3','','VENDOR 4',''];
  reqs.push({ repeatCell: {
    range: gridRange(VEN,2,3,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A3`, values: [vendorColHeaders] });

  // Rows 3-14: vendor comparison rows (0-indexed 3-14)
  const vendorRowData = [
    ['Venue Name',      'La Maison Garden Room','',  'Garden & Co Florists','',    'Sweet Blooms Bakery','',    ''],
    ['Category',        'Venue','',                  'Florist','',                 'Baker','',                  ''],
    ['Contact Name',    'Isabelle Moreau','',         'Priya Nair','',              'Hannah White','',           ''],
    ['Phone',           '+44 7890 123456','',         '+44 7812 345678','',         '+44 7723 456789','',        ''],
    ['Email',           'events@lamaisongarden.co.uk','', 'hello@gardenandco.co.uk','', 'orders@sweetblooms.co.uk','', ''],
    ['Website',         'lamaisongarden.co.uk','',   'gardenandco.co.uk','',       'sweetblooms.co.uk','',      ''],
    ['Quote ($)',        600,'',                      420,'',                       220,'',                      ''],
    ['Deposit Paid ($)', 300,'',                      0,'',                         0,'',                        ''],
    ['Balance Due ($)',  300,'',                      420,'',                       220,'',                      ''],
    ['Contract Signed?','Yes','',                    'No','',                      'No','',                     ''],
    ['Rating',          '5 — Excellent','',          '4 — Very Good','',           '','',                       ''],
    ['Notes',           'Booked & confirmed','',     'Awaiting confirmation','',   'Quote received','',         ''],
  ];
  for (let i = 0; i < vendorRowData.length; i++) {
    const ri = 3 + i;
    const bg = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(VEN, ri, ri+1, 0, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    vals.push({ range: `${TAB}!A${ri+1}`, values: [vendorRowData[i]] });
  }

  // Row 15: summary strip (0-indexed 15)
  reqs.push({ repeatCell: {
    range: gridRange(VEN,15,16,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.parchment),
      textFormat: { foregroundColor: hex(C.muted), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A16`, values: [['Total vendors: 3','','','Booked: 1','','Pending: 2','','TBC: 1','']] });

  // Divider (0-indexed 16)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: VEN, dimension: 'ROWS', startIndex: 16, endIndex: 17 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(VEN,16,17,0,10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Section 2: Bridesmaid Task Tracker (0-indexed 17)
  reqs.push({ mergeCells: { range: gridRange(VEN,17,18,0,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(VEN,17,18,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A18`, values: [['BRIDESMAID TASK TRACKER']] });

  // BM col headers (0-indexed 18)
  const bmHeaders = ['NAME','ROLE','TASK','CATEGORY','PRIORITY','STATUS','DUE DATE','DONE?','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(VEN,18,19,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A19`, values: [bmHeaders] });

  // BM tasks rows 20-37 (0-indexed 19-36)
  for (let i = 0; i < BM_TASKS.length; i++) {
    const ri  = 19 + i;
    const row = 20 + i;
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(VEN, ri, ri+1, 0, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    const [name, role, task, cat, pri, status] = BM_TASKS[i];
    vals.push({ range: `${TAB}!A${row}`, values: [[name, role, task, cat, pri, status, '', status === 'Done' ? 'TRUE' : 'FALSE', '']] });
  }

  // Divider (0-indexed 37)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: VEN, dimension: 'ROWS', startIndex: 37, endIndex: 38 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(VEN,37,38,0,10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Section 3: Supplier Contacts (0-indexed 38)
  reqs.push({ mergeCells: { range: gridRange(VEN,38,39,0,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(VEN,38,39,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A39`, values: [['SUPPLIER CONTACTS']] });

  // Supplier col headers (0-indexed 39)
  const supHeaders = ['SUPPLIER','CATEGORY','CONTACT NAME','PHONE','EMAIL','QUOTE ($)','DEPOSIT ($)','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(VEN,39,40,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A40`, values: [supHeaders] });

  // Supplier rows (0-indexed 40-47)
  for (let i = 0; i < SUPPLIERS.length; i++) {
    const ri  = 40 + i;
    const row = 41 + i;
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(VEN, ri, ri+1, 0, 8),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [SUPPLIERS[i]] });
  }

  await batchUpdate(id, reqs, 'vendors-format');
  await valuesBatchUpdate(id, vals, 'vendors-values');
  console.log('Venue, Vendors & Bridesmaids complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
