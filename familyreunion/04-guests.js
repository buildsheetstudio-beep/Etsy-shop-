'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GST = sheetMap['👨‍👩‍👧‍👦 Guest List & RSVPs'];

// name, branch, relationship, ageGroup, phone, email, city, rsvp, rsvpDate, dietary, accomNeeded, accomAssigned, contribution, notes
const GUESTS = [
  ['James Henderson','Henderson Side','Organiser','Adult (18–59)','0412 555 001','james@email.com','Melbourne, VIC','Confirmed','2025-04-10','None',true,'Lake Eildon Cabin — Master Room','Organiser',''],
  ['Sarah Henderson','Henderson Side',"Organiser's spouse",'Adult (18–59)','','','Melbourne, VIC','Confirmed','2025-04-10','None',true,'Lake Eildon Cabin — Master Room','Co-organiser',''],
  ['Liam Henderson','Henderson Side','Son','Teen (13–17)','','','Melbourne, VIC','Confirmed','2025-04-10','None',false,'Lake Eildon Cabin — Room 2','',''],
  ['Emma Henderson','Henderson Side','Daughter','Child (3–12)','','','Melbourne, VIC','Confirmed','2025-04-10','None',false,'Lake Eildon Cabin — Room 2','',''],
  ['Margaret Henderson','Henderson Side','Mother','Senior (60+)','0412 555 002','','Geelong, VIC','Confirmed','2025-04-15','Gluten Free',true,'Lake Eildon Cabin — Room 2','',''],
  ['Claire Henderson','Henderson Side','Sister','Adult (18–59)','0412 555 003','','Geelong, VIC','Confirmed','2025-04-15','None',false,'Lake Eildon Cabin — Room 2','Salads',''],
  ['Tom Henderson','Henderson Side','Brother','Adult (18–59)','0412 555 004','','Sydney, NSW','Confirmed','2025-04-20','None',true,'Overflow — Eildon Hotel Room 1','Drinks run','Flies in Aug 15'],
  ['Kate Henderson','Henderson Side','Sister-in-law','Adult (18–59)','','','Sydney, NSW','Confirmed','2025-04-20','Vegetarian',true,'Overflow — Eildon Hotel Room 1','Vegetarian dishes',''],
  ['Diana Thompson','Thompson Side','Aunt','Senior (60+)','0412 555 010','','Brisbane, QLD','Confirmed','2025-04-18','None',true,'Lake Eildon Cabin — Room 3','','Flies in Aug 15'],
  ['Robert Thompson','Thompson Side','Uncle','Senior (60+)','','','Brisbane, QLD','Confirmed','2025-04-18','None',true,'Lake Eildon Cabin — Room 3','BBQ team',''],
  ['Lucy Thompson','Thompson Side','Cousin','Adult (18–59)','0412 555 011','','Brisbane, QLD','Confirmed','2025-04-18','None',false,'Lake Eildon Cabin — Room 3','Desserts',''],
  ['Jack Thompson','Thompson Side','Cousin','Teen (13–17)','','','Brisbane, QLD','Confirmed','2025-04-18','None',false,'Local Caravan Site','',''],
  ['Mia Thompson','Thompson Side','Cousin','Teen (13–17)','','','Brisbane, QLD','Confirmed','2025-04-18','None',false,'Local Caravan Site','',''],
  ['Peter Wilson','Wilson Side','Uncle','Adult (18–59)','0412 555 020','','Melbourne, VIC','Confirmed','2025-04-22','None',true,'Lake Eildon Cabin — Room 4','BBQ team',''],
  ['Anna Wilson','Wilson Side','Aunt','Adult (18–59)','','','Melbourne, VIC','Confirmed','2025-04-22','Dairy Free',true,'Lake Eildon Cabin — Room 4','Coleslaw',''],
  ['Chloe Wilson','Wilson Side','Cousin','Child (3–12)','','','Melbourne, VIC','Confirmed','2025-04-22','None',false,'Lake Eildon Cabin — Room 4','',''],
  ['Noah Wilson','Wilson Side','Cousin','Child (3–12)','','','Melbourne, VIC','Confirmed','2025-04-22','None',false,'Lake Eildon Cabin — Room 4','',''],
  ['David Smith','Smith Side','Cousin','Adult (18–59)','0412 555 030','','Adelaide, SA','Confirmed','2025-04-25','None',true,'Overflow — Eildon Hotel Room 2','Drinks supply','Drives up Aug 15'],
  ['Rachel Smith','Smith Side',"Cousin's wife",'Adult (18–59)','','','Adelaide, SA','Confirmed','2025-04-25','Vegan',true,'Overflow — Eildon Hotel Room 2','Vegan dishes',''],
  ['Isla Smith','Smith Side','2nd cousin','Baby (0–2)','','','Adelaide, SA','Confirmed','2025-04-25','None',false,'Overflow — Eildon Hotel Room 2','',''],
  ['George Henderson','Out-of-Town','Grandfather','Senior (60+)','0412 555 040','','Perth, WA','Confirmed','2025-05-01','None',true,'Overflow — Eildon Hotel Room 2','','Flies in Aug 15'],
  ['Betty Henderson','Out-of-Town','Grandmother','Senior (60+)','','','Perth, WA','Confirmed','2025-05-01','None',true,'Overflow — Eildon Hotel Room 2','',''],
  ['Marcus Henderson','Out-of-Town','Cousin','Adult (18–59)','0412 555 041','','London, UK','No Response','','None',false,'','','Flies if confirmed'],
  ['Brian Collins','Family Friends','Family friend','Adult (18–59)','0412 555 050','','Melbourne, VIC','Confirmed','2025-05-10','None',false,'Self-arranged — AirBnB (local)','',''],
  ['Jenny Collins','Family Friends','Family friend','Adult (18–59)','','','Melbourne, VIC','Declined','2025-05-10','None',false,'','',''],
];

const BRANCHES = ['Henderson Side','Thompson Side','Wilson Side','Smith Side','Out-of-Town','Family Friends'];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'👨‍👩‍👧‍👦 Guest List & RSVPs'";
  const NCOLS = 14; // A-N

  // Column widths: A-N
  [180,150,150,110,130,200,130,110,110,130,80,160,160,200].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GST, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  // Cols O-Q for summary blocks
  [160,120,120].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GST, dimension: 'COLUMNS', startIndex: 14+i, endIndex: 15+i },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner (A standalone for col-A freeze, B:N merged)
  reqs.push({ repeatCell: {
    range: gridRange(GST,0,1,0,1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.forestGreen) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(GST,0,1,1,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GST,0,1,1,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['  👨‍👩‍👧‍👦  Guest List & RSVPs — Henderson Family Reunion 2025']] });

  // Row 1: summary header
  reqs.push({ mergeCells: { range: gridRange(GST,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GST,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['="  👥 Total: "&COUNTA(A4:A200)&"  |  Confirmed: "&COUNTIF(H4:H200,"Confirmed")&"  |  Declined: "&COUNTIF(H4:H200,"Declined")&"  |  Awaiting: "&COUNTIF(H4:H200,"No Response")+COUNTIF(H4:H200,"Invited")&"  |  % Confirmed: "&TEXT(IFERROR(COUNTIF(H4:H200,"Confirmed")/COUNTA(A4:A200),0),"0%")']] });

  // Row 2: column headers
  const hdrs = ['GUEST NAME','FAMILY BRANCH','RELATIONSHIP','AGE GROUP','PHONE','EMAIL','CITY / STATE','RSVP STATUS','RSVP DATE','DIETARY NEEDS','ACCOM NEEDED?','ACCOM ASSIGNED','CONTRIBUTION','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(GST,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [hdrs] });

  // Data rows
  reqs.push({ repeatCell: {
    range: gridRange(GST,3,3+GUESTS.length,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.white),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  for (let i = 0; i < GUESTS.length; i++) {
    if (i % 2 === 1) {
      reqs.push({ repeatCell: {
        range: gridRange(GST,3+i,4+i,0,NCOLS),
        cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) }},
        fields: 'userEnteredFormat.backgroundColor',
      }});
    }
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 3, endIndex: 3+GUESTS.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Date format col I (index 8)
  reqs.push({ repeatCell: {
    range: gridRange(GST,3,3+GUESTS.length,8,9),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }}},
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Center cols H, I, J, K
  reqs.push({ repeatCell: {
    range: gridRange(GST,3,3+GUESTS.length,7,11),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Checkbox validation col K (index 10)
  reqs.push({ setDataValidation: {
    range: gridRange(GST,3,3+GUESTS.length,10,11),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
  }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(GST,2,3+GUESTS.length,0,NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Write guest data
  for (let i = 0; i < GUESTS.length; i++) {
    const [name, branch, rel, age, phone, email, city, rsvp, rsvpDate, dietary, accomNeeded, accomAssigned, contribution, notes] = GUESTS[i];
    const row = 4 + i;
    vals.push({ range: `${TAB}!A${row}`, values: [[name, branch, rel, age, phone, email, city, rsvp, rsvpDate, dietary, accomNeeded ? 'TRUE' : 'FALSE', accomAssigned, contribution, notes]] });
  }

  // ── Family Branch Legend (cols O-P rows 1-7) ──
  const branchColors = [
    ['Henderson Side', '#FDE8D0', '#8B3A1A'],
    ['Thompson Side',  '#D0E8FD', '#1A3A8B'],
    ['Wilson Side',    '#D0FDE8', '#1A8B4A'],
    ['Smith Side',     '#FDD0E8', '#8B1A5A'],
    ['Out-of-Town',    '#E8D0FD', '#5A1A8B'],
    ['Family Friends', '#FDFFE0', '#6A6A00'],
  ];

  reqs.push({ repeatCell: {
    range: gridRange(GST,0,1,14,17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!O1`, values: [['BRANCH','COLOUR','GUESTS']] });

  branchColors.forEach(([branch, bg, fg], i) => {
    const row = 2 + i;
    reqs.push({ repeatCell: {
      range: gridRange(GST, 1+i, 2+i, 14, 15),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(GST, 1+i, 2+i, 15, 16),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 9 },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(GST, 1+i, 2+i, 16, 17),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.paleGreen),
        textFormat: { foregroundColor: hex(C.bodyText), bold: true, fontSize: 9 },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `${TAB}!O${row}`, values: [[branch, '●', `=COUNTIF(B4:B200,"${branch}")`]] });
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 1, endIndex: 8 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // ── Family Branch breakdown (cols O-P rows 9-15) ──
  reqs.push({ repeatCell: {
    range: gridRange(GST,8,9,14,17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!O9`, values: [['BRANCH','TOTAL','CONFIRMED']] });
  BRANCHES.forEach((branch, i) => {
    const row = 10 + i;
    reqs.push({ repeatCell: {
      range: gridRange(GST,9+i,10+i,14,17),
      cell: { userEnteredFormat: {
        backgroundColor: i % 2 === 0 ? hex(C.warmCream) : hex(C.white),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `${TAB}!O${row}`, values: [[branch, `=COUNTIF(B4:B200,"${branch}")`, `=COUNTIFS(B4:B200,"${branch}",H4:H200,"Confirmed")`]] });
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 8, endIndex: 16 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // ── Age group breakdown (cols O-P rows 17-22) ──
  const ageGroups = ['Baby (0–2)','Child (3–12)','Teen (13–17)','Adult (18–59)','Senior (60+)'];
  reqs.push({ repeatCell: {
    range: gridRange(GST,16,17,14,17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!O17`, values: [['AGE GROUP','TOTAL','']] });
  ageGroups.forEach((age, i) => {
    const row = 18 + i;
    reqs.push({ repeatCell: {
      range: gridRange(GST,17+i,18+i,14,17),
      cell: { userEnteredFormat: {
        backgroundColor: i % 2 === 0 ? hex(C.warmCream) : hex(C.white),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `${TAB}!O${row}`, values: [[age, `=COUNTIF(D4:D200,"${age}")`, '']] });
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 16, endIndex: 23 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // RSVP chart data block (cols P-Q rows 24-30) for donut chart
  const rsvpStatuses = ['Confirmed','Declined','Invited','No Response','Waitlisted'];
  reqs.push({ repeatCell: {
    range: gridRange(GST,23,24,15,17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!P24`, values: [['RSVP STATUS','COUNT']] });
  rsvpStatuses.forEach((status, i) => {
    vals.push({ range: `${TAB}!P${25+i}`, values: [[status, `=COUNTIF(H4:H200,"${status}")`]] });
  });

  await batchUpdate(id, reqs, 'guests-format');
  await valuesBatchUpdate(id, vals, 'guests-values');
  console.log('Guest List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
