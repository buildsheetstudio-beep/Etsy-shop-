'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL = sheetMap['👰 Guest List & Seating'];

// Layout:
// Row 0: Title — A0 empty (allows col-A freeze), B0:M0 merged title
// Row 1: Summary stats — A1 standalone empty, then merged pairs from col B
// Row 2: Col headers
// Rows 3-22: 20 guests
// Rows 23-53: Buffer (31 empty rows)
// Row 54: Divider
// Row 55: Seating section header
// Row 56: Table name headers
// Row 57: CAPACITY row
// Row 58: FILLED formula row
// Row 59: SEATS LEFT row
// Rows 60-69: Seat assignment rows (10 seats per table)
// RSVP chart data: rows 1-5, cols O-P (0-indexed)

const NCOLS = 13; // A-M

// [name, relationship, phone, email, rsvpStatus, rsvpDate, dietary, plusOne, plusOneName, giftRcvd, tyDone, tableAssigned, notes]
const GUESTS = [
  ['Charlotte Harrison','Bride','','charlotte@gmail.com','Confirmed','01 Jul 2025','None',false,'',false,false,'Bride\'s Table','Guest of honour'],
  ['Margaret Harrison','Family (Bride\'s)','0412 555 001','','Confirmed','05 Jul 2025','None',true,'Robert Harrison',false,false,'Table 1','Charlotte\'s mum'],
  ['Claire Harrison','Family (Bride\'s)','0412 555 002','','Confirmed','08 Jul 2025','Vegetarian',false,'',false,false,'Table 1','Charlotte\'s sister'],
  ['Diana Thompson','Family (Partner\'s)','0412 555 003','','Confirmed','10 Jul 2025','None',true,'James Thompson',false,false,'Table 2','Partner\'s mum'],
  ['Sophie Clarke','MOH','0412 555 004','sophie@gmail.com','Confirmed','01 Jul 2025','None',false,'',false,false,'Bride\'s Table','Maid of Honour'],
  ['Lily Thompson','Bridesmaid','0412 555 005','','Confirmed','01 Jul 2025','Gluten Free',false,'',false,false,'Bride\'s Table','Bridesmaid'],
  ['Emma Grant','Bridesmaid','0412 555 006','','Confirmed','01 Jul 2025','None',false,'',false,false,'Bride\'s Table','Bridesmaid'],
  ['Lucy Chen','Close Friend','0412 555 007','','Confirmed','12 Jul 2025','Vegan',false,'',false,false,'Table 1',''],
  ['Mia Patel','Close Friend','0412 555 008','','Confirmed','15 Jul 2025','None',false,'',false,false,'Table 2',''],
  ['Zoe Roberts','Work Friend','0412 555 009','','Confirmed','15 Jul 2025','Dairy Free',false,'',false,false,'Table 2',''],
  ['Amy Johnson','Work Friend','0412 555 010','','Confirmed','18 Jul 2025','None',false,'',false,false,'Table 3',''],
  ['Natalie Wong','Close Friend','0412 555 011','','Confirmed','20 Jul 2025','Nut Allergy',true,'Peter Wong',false,false,'Table 3',''],
  ['Rachel Green','Mutual Friend','0412 555 012','','Declined','14 Jul 2025','None',false,'',false,false,'Not Yet Assigned','Travelling'],
  ['Olivia Brown','Close Friend','0412 555 013','','Confirmed','22 Jul 2025','None',false,'',false,false,'Table 3',''],
  ['Kate Wilson','Work Friend','0412 555 014','','Confirmed','25 Jul 2025','Vegetarian',false,'',false,false,'Table 4',''],
  ['Jessica Lee','Family (Bride\'s)','0412 555 015','','Confirmed','10 Jul 2025','None',false,'',false,false,'Table 1',''],
  ['Hannah Davis','Mutual Friend','0412 555 016','','No Response','','None',false,'',false,false,'Not Yet Assigned',''],
  ['Anna Scott','Family (Partner\'s)','0412 555 017','','Confirmed','28 Jul 2025','None',false,'',false,false,'Table 4',''],
  ['Beth Carter','Close Friend','0412 555 018','','Invited','','None',false,'',false,false,'Not Yet Assigned','RSVP pending'],
  ['Chloe Evans','Close Friend','0412 555 019','','Confirmed','28 Jul 2025','None',false,'',false,false,'Table 4',''],
];

(async () => {
  const reqs = [];

  // Column widths A-M
  [180,150,130,200,110,110,130,80,160,80,80,120,200].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GL, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  // Col N spacer + O, P chart data cols
  [30, 150, 80].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GL, dimension: 'COLUMNS', startIndex: 13 + i, endIndex: 14 + i },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  // Chart area Q-U
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'COLUMNS', startIndex: 16, endIndex: 22 },
    properties: { pixelSize: 80 }, fields: 'pixelSize',
  }});

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 3, endIndex: 54 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 54, endIndex: 55 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 55, endIndex: 57 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 57, endIndex: 70 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Title row: A0 standalone (empty), B0:M0 merged
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 1, 0, 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.dustyRose) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(GL, 0, 1, 1, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 1, 1, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 1 — A1 standalone empty, then pairs from col B
  reqs.push({ repeatCell: {
    range: gridRange(GL, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.parchment),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // A1 single cell (standalone — critical for col freeze)
  reqs.push({ mergeCells: { range: gridRange(GL, 1, 2, 0, 1), mergeType: 'MERGE_ALL' }});
  // Pairs: B1:D1, E1:G1, H1:J1, K1:M1
  [[1,4],[4,7],[7,10],[10,13]].forEach(([c1,c2]) => {
    reqs.push({ mergeCells: { range: gridRange(GL, 1, 2, c1, c2), mergeType: 'MERGE_ALL' }});
  });

  // Col headers row 2
  reqs.push({ repeatCell: {
    range: gridRange(GL, 2, 3, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows 3-22 (guests) + buffer 23-53
  for (let r = 3; r < 54; r++) {
    const bg = (r % 2 === 1) ? C.white : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(GL, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Center cols: E(4)=RSVP, F(5)=RSVP date, G(6)=dietary, H(7)=plus one, J(9)=gift, K(10)=TY, L(11)=table
  [4,5,6,7,9,10,11].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(GL, 3, 54, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  });
  // Date col F text format
  reqs.push({ repeatCell: {
    range: gridRange(GL, 3, 54, 5, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'TEXT' } }},
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Checkbox validation cols H(7), J(9), K(10)
  [7, 9, 10].forEach(c => {
    reqs.push({ setDataValidation: {
      range: gridRange(GL, 3, 54, c, c + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Borders on main table
  reqs.push({ updateBorders: {
    range: gridRange(GL, 2, 54, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Divider row 54
  reqs.push({ repeatCell: {
    range: gridRange(GL, 54, 55, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.dustyRose) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Seating section header row 55
  reqs.push({ mergeCells: { range: gridRange(GL, 55, 56, 0, 13), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GL, 55, 56, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Seating plan: table headers row 56 (cols A-F = 6 table cols)
  reqs.push({ repeatCell: {
    range: gridRange(GL, 56, 57, 0, 7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Capacity / Filled / Seats Left rows 57-59 — label col A, data cols B-G
  [57, 58, 59].forEach(r => {
    reqs.push({ repeatCell: {
      range: gridRange(GL, r, r + 1, 0, 7),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.parchment),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9, bold: true },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  });
  // Seats Left row — gold text for emphasis
  reqs.push({ repeatCell: {
    range: gridRange(GL, 59, 60, 1, 7),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.gold), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat.textFormat',
  }});

  // Seat assignment rows 60-69 — ivory bg (input cells)
  for (let r = 60; r < 70; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(GL, r, r + 1, 0, 7),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.ivory),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
  }

  // Borders on seating plan
  reqs.push({ updateBorders: {
    range: gridRange(GL, 56, 70, 0, 7),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.dustyRose), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // RSVP chart data block (cols O-P = indices 14-15, rows 1-5)
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 6, 14, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.parchment),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 1, 14, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  await batchUpdate(id, reqs, 'guestlist-format');

  // Values
  const vals = [];
  vals.push({ range: "'👰 Guest List & Seating'!B1", values: [['  👰  Guest List, RSVPs & Seating — Charlotte\'s Bridal Shower']] });

  // Summary row 2 (1-indexed = index 1)
  vals.push({ range: "'👰 Guest List & Seating'!B2", values: [['=COUNTA(A4:A54)&" invited"']] });
  vals.push({ range: "'👰 Guest List & Seating'!E2", values: [['=COUNTIF(E4:E54,"Confirmed")&" confirmed"']] });
  vals.push({ range: "'👰 Guest List & Seating'!H2", values: [['=COUNTIF(E4:E54,"Declined")&" declined · "&(COUNTIFS(E4:E54,"Invited")+COUNTIF(E4:E54,"No Response"))&" awaiting"']] });
  vals.push({ range: "'👰 Guest List & Seating'!K2", values: [['=IFERROR(SPARKLINE(COUNTIF(E4:E54,"Confirmed")/COUNTA(A4:A54),{"charttype","bar";"color1","#8B5E6A";"color2","#E8D5D8";"max",1}),"—")']] });

  // Col headers row 3 (1-indexed)
  vals.push({ range: "'👰 Guest List & Seating'!A3", values: [['GUEST NAME','RELATIONSHIP','PHONE','EMAIL','RSVP STATUS','RSVP DATE','DIETARY NEEDS','PLUS ONE?','PLUS ONE NAME','GIFT RCVD?','THANK YOU?','TABLE ASSIGNED','NOTES']] });

  // Guest data
  const guestRows = GUESTS.map(g => [
    g[0],g[1],g[2],g[3],g[4],g[5],g[6],
    g[7] ? 'TRUE' : 'FALSE',
    g[8],
    g[9] ? 'TRUE' : 'FALSE',
    g[10] ? 'TRUE' : 'FALSE',
    g[11],g[12],
  ]);
  vals.push({ range: "'👰 Guest List & Seating'!A4", values: guestRows });

  // RSVP chart data (cols O-P, rows 1-6, 1-indexed)
  vals.push({ range: "'👰 Guest List & Seating'!O1", values: [['RSVP STATUS','COUNT']] });
  [['Confirmed'],['Declined'],['Invited'],['No Response'],['Waitlisted']].forEach(([s], i) => {
    vals.push({ range: `'👰 Guest List & Seating'!O${2+i}`, values: [[s, `=COUNTIF($E$4:$E$54,"${s}")`]] });
  });

  // Seating plan section header (row 56, 1-indexed)
  vals.push({ range: "'👰 Guest List & Seating'!A56", values: [['  💺  Seating Plan']] });

  // Table name headers row 57 (1-indexed)
  vals.push({ range: "'👰 Guest List & Seating'!A57", values: [["SEAT","Bride's Table","Table 1","Table 2","Table 3","Table 4","Table 5"]] });

  // Capacity row 58
  vals.push({ range: "'👰 Guest List & Seating'!A58", values: [['CAPACITY','8','8','8','8','8','8']] });

  // Filled row 59 — COUNTIF from guest list col L
  vals.push({ range: "'👰 Guest List & Seating'!A59", values: [['FILLED']] });
  ["Bride's Table","Table 1","Table 2","Table 3","Table 4","Table 5"].forEach((t, i) => {
    const col = String.fromCharCode(66 + i); // B, C, D, E, F, G
    vals.push({ range: `'👰 Guest List & Seating'!${col}59`, values: [[`=COUNTIF($L$4:$L$54,"${t}")`]] });
  });

  // Seats Left row 60 — capacity minus filled
  vals.push({ range: "'👰 Guest List & Seating'!A60", values: [['SEATS LEFT']] });
  ['B','C','D','E','F','G'].forEach((col, i) => {
    vals.push({ range: `'👰 Guest List & Seating'!${col}60`, values: [[`=B58-B59`]] });
  });
  // Fix: use correct col references for each
  ['B','C','D','E','F','G'].forEach((col, i) => {
    vals.push({ range: `'👰 Guest List & Seating'!${col}60`, values: [[`=${col}58-${col}59`]] });
  });

  // Pre-filled seat assignments (rows 61-70, 1-indexed)
  // Bride's Table: Charlotte, Sophie, Lily, Emma
  // Table 1: Margaret, Claire, Lucy, Jessica
  // Table 2: Diana, Mia, Zoe
  // Table 3: Amy, Natalie, Olivia
  // Table 4: Kate, Anna, Chloe
  const seatData = [
    ['Seat 1','Charlotte H.','Margaret H.','Diana T.','Amy J.','Kate W.',''],
    ['Seat 2','Sophie C.','Claire H.','Mia P.','Natalie W.','Anna S.',''],
    ['Seat 3','Lily T.','Lucy C.','Zoe R.','Olivia B.','Chloe E.',''],
    ['Seat 4','Emma G.','Jessica L.','James T.','','',''],
    ['Seat 5','','Robert H.','','','',''],
    ['Seat 6','','','','Peter W.','',''],
    ['Seat 7','','','','','',''],
    ['Seat 8','','','','','',''],
    ['Seat 9','','','','','',''],
    ['Seat 10','','','','','',''],
  ];
  vals.push({ range: "'👰 Guest List & Seating'!A61", values: seatData });

  await valuesBatchUpdate(id, vals, 'guestlist-values');
  console.log('Guest List & Seating complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
