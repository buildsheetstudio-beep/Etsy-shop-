'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const NC = sheetMap['📝 Event Notes & Contacts'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 1, endIndex: 65 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Column widths: A(30) B(Name/Type 200) C(Role/Cat 130) D(Event 180) E(Phone/Date 120) F(Email/Note 280) G(Pref Contact 100) H(Notes 200)
  const colW = [30, 200, 130, 180, 120, 280, 100, 200];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row
  reqs.push({ mergeCells: { range: gridRange(NC, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(NC, 0, 1, 0, 8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyBrown),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Key Contacts section header row 2
  reqs.push({ mergeCells: { range: gridRange(NC, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(NC, 1, 2, 0, 8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Contacts header row 3
  reqs.push({ repeatCell: {
    range: gridRange(NC, 2, 3, 0, 8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyBrown),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Contacts data rows 4-21 (indices 3-20)
  for (let r = 3; r < 21; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(NC, r, r+1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Notes section header row 22 (index 21)
  reqs.push({ mergeCells: { range: gridRange(NC, 21, 22, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(NC, 21, 22, 0, 8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Notes header row 23 (index 22)
  reqs.push({ repeatCell: {
    range: gridRange(NC, 22, 23, 0, 8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyBrown),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Notes data rows 24-65 (indices 23-64)
  for (let r = 23; r < 65; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(NC, r, r+1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Note rows — wrap text in F col (index 5)
  reqs.push({ repeatCell: {
    range: gridRange(NC, 23, 65, 5, 6),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat.wrapStrategy',
  }});
  // Taller rows for notes
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 23, endIndex: 65 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  // Date format col E in notes section
  reqs.push({ repeatCell: {
    range: gridRange(NC, 23, 65, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Center col A, C
  [0, 2].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(NC, 2, 65, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders contacts
  reqs.push({ updateBorders: {
    range: gridRange(NC, 2, 21, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});
  // Borders notes
  reqs.push({ updateBorders: {
    range: gridRange(NC, 22, 65, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'notes-format');

  const data = [];
  data.push({ range: `'📝 Event Notes & Contacts'!A1`, values: [['📝 Event Notes & Contacts']] });
  data.push({ range: `'📝 Event Notes & Contacts'!A2`, values: [['  📞 Key Contacts']] });
  data.push({ range: `'📝 Event Notes & Contacts'!A3:H3`, values: [['#','Full Name','Role','Event','Phone','Email','Pref. Contact','Notes']] });

  const contacts = [
    [1,'Sarah Johnson','Organiser','All Events','0400 001 001','sarah.j@company.com','Email','Primary event organiser'],
    [2,'Mark Davis','Co-Organiser','All Events','0400 001 002','mark.d@company.com','Phone','AV & logistics lead'],
    [3,'Lisa Chen','Co-Organiser','All Events','0400 001 003','lisa.c@company.com','Email','Guest management & marketing'],
    [4,'David Wong','Venue Contact','Tech Summit 2025','02 9000 1111','events@gcc.com.au','Email','Grand Convention Centre manager'],
    [5,'Maria Santos','Caterer','Tech Summit 2025','0411 200 001','maria@freshtable.com','Email','Fresh Table Catering Co.'],
    [6,'Sam Briggs','AV Contact','Tech Summit 2025','0411 200 003','sam@proav.com.au','Phone','ProAV Solutions — on-site day of'],
    [7,'Anne-Marie Cole','Venue Contact','Charity Gala 2025','02 9000 2222','events@riverside.com','Email','Riverside Ballroom coordinator'],
    [8,'Raj Kapoor','Caterer','Charity Gala 2025','0411 200 004','raj@prestigecatering.com','Email','Prestige Catering'],
    [9,'Sophie Green','Florist','Charity Gala 2025','0411 200 006','sophie@bloomdrape.com','Phone','Bloom & Drape Florist'],
    [10,'Nico Ferreira','Entertainment','Charity Gala 2025','0411 200 005','nico@thejazzcollective.com','Phone','The Jazz Collective'],
    [11,'Michael Osei','Venue Contact','Annual Awards Night','02 9000 3333','events@grandhotel.com','Email','The Grand Hotel events team'],
    [12,'Clara Perez','Venue Contact','5-Year Anniversary Dinner','02 9000 4444','events@harbourview.com','Email','Harbour View Restaurant'],
    [13,'Emma Walsh','Donor Contact','Charity Gala 2025','0412 000 008','e.walsh@walshphil.com','Email','Major donor — VIP treatment'],
    [14,'James Hartley','Speaker','Tech Summit 2025','0412 000 001','j.hartley@innovatetech.com','Email','Keynote speaker — confirm slides 48h prior'],
    [15,'Oliver Chen','Moderator','Tech Summit 2025','0412 000 003','o.chen@cloudbase.com','Email','Panel moderator'],
  ];

  data.push({ range: `'📝 Event Notes & Contacts'!A4:H18`, values: contacts });
  data.push({ range: `'📝 Event Notes & Contacts'!A22`, values: [['  📓 Event Notes & Action Items']] });
  data.push({ range: `'📝 Event Notes & Contacts'!A23:H23`, values: [['#','Event','Note Type','Added By','Date','Note / Detail','Follow-Up?','Status']] });

  const notes = [
    [1,'Tech Summit 2025','Action Item','Sarah Johnson','2025-07-01','Confirm final headcount with venue by 1 Sep — numbers needed for catering too','Yes','Done'],
    [2,'Tech Summit 2025','Risk','Mark Davis','2025-07-15','AV tech rehearsal MUST happen day before. If ProAV cancel last minute — contact BackupAV Solutions: 0411 999 001','Yes','In Progress'],
    [3,'Tech Summit 2025','Decision','Sarah Johnson','2025-07-20','Agreed to cap registration at 160 (10 buffer over venue capacity of 150)','No','Done'],
    [4,'Charity Gala 2025','Action Item','Lisa Chen','2025-07-25','Confirm table layout with Riverside Ballroom by 1 Sep. Send seating plan PDF to Anne-Marie','Yes','Not Started'],
    [5,'Charity Gala 2025','Risk','Mark Davis','2025-08-01','Jazz Collective contract not signed yet — follow up Nico by 15 Aug or find alternative','Yes','Not Started'],
    [6,'Charity Gala 2025','Lesson Learned','Sarah Johnson','2025-08-05','From last year: start auction item collection 3 months out — not 6 weeks','No','Done'],
    [7,'Holiday Party 2025','Idea','Lisa Chen','2025-08-10','Photo booth idea — contact SnapBooth Studios for quote. Could be great for social media content','Yes','Not Started'],
    [8,'General','Follow-Up','Sarah Johnson','2025-08-15','Review all vendor contracts by end of August for annual terms renewal','Yes','Not Started'],
    [9,'Investor Day 2025','Decision','Sarah Johnson','2025-08-20','Format confirmed: 2 hours boardroom, 1 hour Zoom Q&A, then lunch','No','Done'],
    [10,'Annual Awards Night','Action Item','Lisa Chen','2025-09-01','Collect all employee nominations from managers by 30 Sep for awards committee review','Yes','Not Started'],
  ];

  data.push({ range: `'📝 Event Notes & Contacts'!A24:H33`, values: notes });

  await valuesBatchUpdate(id, data, 'notes-values');
  console.log('Notes & Contacts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
