'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL = sheetMap['👥 Guest List & RSVPs'];

// Layout: Row1=Title(B1 merge, A1 empty), Row2=Summary, Row3=Headers, Rows4-23=Data
// Cols A-M data, N=spacer, O-P=relationship breakdown, Q+=chart
// Col A MUST be empty in rows 1-2 for COUNTA formula to work (counts header + 20 names = 21 - 1 = 20)

const NCOLS = 13; // A-M

// [name, relationship, phone, email, rsvp, rsvpDate, dietary, plusOne(T/F), plusOneName, giftRcvd(T/F), giftDesc, tySent(T/F), notes]
const GUESTS = [
  ['Sarah Mitchell','Mum-to-Be','','sarah@gmail.com','Confirmed','01 Jul 2025','None',false,'',true,'Personalised keepsake book',false,'Guest of honour'],
  ['Margaret Mitchell','Family (Mum\'s side)','0412 555 001','','Confirmed','05 Jul 2025','None',true,'Robert Mitchell',true,'Baby clothing bundle',true,'Sarah\'s mum'],
  ['Claire Mitchell','Family (Mum\'s side)','0412 555 002','','Confirmed','08 Jul 2025','Vegetarian',false,'',true,'Gift card ($100)',true,'Sarah\'s sister'],
  ['Diana Thompson','Family (Partner\'s side)','0412 555 003','','Confirmed','10 Jul 2025','None',true,'James Thompson',true,'Pram accessories',true,'Partner\'s mum'],
  ['Lucy Chen','Close Friend','0412 555 004','lucy@gmail.com','Confirmed','03 Jul 2025','Gluten Free',false,'',true,'Awaiting gift',false,''],
  ['Emma Clarke','Close Friend','0412 555 005','emma@gmail.com','Confirmed','01 Jul 2025','None',false,'',true,'Nursing pillow + covers',true,'Co-host'],
  ['Mia Patel','Close Friend','0412 555 006','','Confirmed','12 Jul 2025','Vegan',false,'',true,'Awaiting',false,''],
  ['Zoe Roberts','Work Friend','0412 555 007','','Confirmed','15 Jul 2025','None',false,'',true,'Baby monitor',true,''],
  ['Amy Johnson','Work Friend','0412 555 008','','Confirmed','15 Jul 2025','None',false,'',true,'Gift card ($50)',true,''],
  ['Natalie Wong','Close Friend','0412 555 009','','Confirmed','18 Jul 2025','Dairy Free',true,'Peter Wong',true,'Nursery décor set',true,''],
  ['Rachel Green','Neighbour','0412 555 010','','Confirmed','20 Jul 2025','None',false,'',true,'Awaiting',false,''],
  ['Sophie Taylor','Close Friend','0412 555 011','','Declined','14 Jul 2025','','false','',false,'',false,'Travelling that weekend'],
  ['Olivia Brown','Close Friend','0412 555 012','','Confirmed','22 Jul 2025','None',false,'',true,'Baby bath set',false,'Thank-you to send'],
  ['Hannah Davis','Partner\'s Friend','0412 555 013','','No Response','','None',false,'',false,'',false,''],
  ['Kate Wilson','Work Friend','0412 555 014','','Confirmed','25 Jul 2025','Vegetarian',false,'',true,'Baby clothing bundle',true,''],
  ['Jessica Lee','Family (Mum\'s side)','0412 555 015','','Confirmed','10 Jul 2025','None',false,'',true,'Breastfeeding bundle',true,''],
  ['Lily Thompson','Close Friend','0412 555 016','','Confirmed','01 Jul 2025','None',false,'',true,'Baby book collection',true,'Co-host'],
  ['Anna Scott','Neighbour','0412 555 017','','Invited','','None',false,'',false,'',false,'RSVP pending'],
  ['Beth Carter','Partner\'s Friend','0412 555 018','','No Response','','None',false,'',false,'',false,''],
  ['Chloe Evans','Close Friend','0412 555 019','','Confirmed','28 Jul 2025','Nut Allergy',false,'',true,'Keepsake blanket',false,''],
];

(async () => {
  const reqs = [];

  // Column widths
  [180,160,130,200,110,110,130,80,160,80,200,80,200].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GL, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  // Spacer col N + Breakdown cols O-P
  [40, 200, 100].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GL, dimension: 'COLUMNS', startIndex: 13 + i, endIndex: 14 + i },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  // Chart area Q-U
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'COLUMNS', startIndex: 16, endIndex: 21 },
    properties: { pixelSize: 120 }, fields: 'pixelSize',
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
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 3, endIndex: 23 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Title row: A1 empty (bg only), B1:M1 merged with title
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 1, 0, 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.lilac) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(GL, 0, 1, 1, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 1, 1, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 2 — A2 intentionally empty
  reqs.push({ repeatCell: {
    range: gridRange(GL, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  [[1,3],[3,5],[5,7],[7,9],[9,11],[11,13]].forEach(([c1,c2]) => {
    reqs.push({ mergeCells: { range: gridRange(GL, 1, 2, c1, c2), mergeType: 'MERGE_ALL' }});
  });

  // SPARKLINE bar for confirmation rate (row 2, col A merged to B)
  reqs.push({ mergeCells: { range: gridRange(GL, 1, 2, 0, 1), mergeType: 'MERGE_ALL' }});

  // Column headers row 3
  reqs.push({ repeatCell: {
    range: gridRange(GL, 2, 3, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  }});

  // Data rows
  for (let r = 3; r < 23; r++) {
    const bg = (r % 2 === 1) ? C.white : C.petalPink;
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

  // RSVP + dietary + plus-one cols — center
  [4, 6, 7, 9, 11].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(GL, 3, 23, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  });

  // Date col F — center
  reqs.push({ repeatCell: {
    range: gridRange(GL, 3, 23, 5, 6),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'TEXT' } }},
    fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
  }});

  // Checkbox validation H, J, L
  [7, 9, 11].forEach(c => {
    reqs.push({ setDataValidation: {
      range: gridRange(GL, 3, 23, c, c + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(GL, 2, 23, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Relationship breakdown (cols O-P, rows 3-11) and RSVP chart data (cols O-P, rows 15-20)
  reqs.push({ repeatCell: {
    range: gridRange(GL, 2, 21, 14, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Breakdown header row
  reqs.push({ repeatCell: {
    range: gridRange(GL, 2, 3, 14, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateBorders: {
    range: gridRange(GL, 2, 21, 14, 17),
    top: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'guest-format');

  // Values
  const vals = [];
  vals.push({ range: "'👥 Guest List & RSVPs'!B1", values: [['  👥  Guest List & RSVPs']] });

  // Summary row 2
  vals.push({ range: "'👥 Guest List & RSVPs'!B2", values: [['=COUNTA(A4:A200)&" invited"']] });
  vals.push({ range: "'👥 Guest List & RSVPs'!D2", values: [['=COUNTIF(E4:E200,"Confirmed")&" confirmed"']] });
  vals.push({ range: "'👥 Guest List & RSVPs'!F2", values: [['=COUNTIF(E4:E200,"Declined")&" declined"']] });
  vals.push({ range: "'👥 Guest List & RSVPs'!H2", values: [['=IFERROR(TEXT(COUNTIF(E4:E200,"Confirmed")/COUNTA(A4:A200),"0%")&" RSVPd","—")']] });
  vals.push({ range: "'👥 Guest List & RSVPs'!J2", values: [['=COUNTIFS(E4:E200,"Invited")+COUNTIF(E4:E200,"No Response")&" awaiting"']] });
  vals.push({ range: "'👥 Guest List & RSVPs'!L2", values: [['=IFERROR(SPARKLINE(COUNTIF(E4:E200,"Confirmed")/COUNTA(A4:A200),{"charttype","bar";"color1","#9B7DB5";"color2","#E5D5DF";"max",1}),"—")']] });

  // Column headers
  vals.push({ range: "'👥 Guest List & RSVPs'!A3", values: [['GUEST NAME','RELATIONSHIP','PHONE','EMAIL','RSVP STATUS','RSVP DATE','DIETARY','PLUS ONE?','PLUS ONE NAME','GIFT RCVD?','GIFT DESCRIPTION','THANK YOU?','NOTES']] });

  // Guest data
  const guestRows = GUESTS.map(g => [
    g[0],g[1],g[2],g[3],g[4],g[5],g[6],
    g[7] ? 'TRUE' : 'FALSE',
    g[8],
    g[9] ? 'TRUE' : 'FALSE',
    g[10],
    g[11] ? 'TRUE' : 'FALSE',
    g[12],
  ]);
  vals.push({ range: "'👥 Guest List & RSVPs'!A4", values: guestRows });

  // Relationship breakdown
  vals.push({ range: "'👥 Guest List & RSVPs'!O3", values: [['RELATIONSHIP','COUNT','']] });
  const RELS = ['Mum-to-Be','Family (Mum\'s side)','Family (Partner\'s side)','Close Friend','Work Friend','Neighbour','Partner\'s Friend','Child','Other'];
  RELS.forEach((rel, i) => {
    vals.push({ range: `'👥 Guest List & RSVPs'!O${4+i}`, values: [[rel, `=COUNTIF($B$4:$B$200,"${rel}")`]] });
  });

  // RSVP data for chart (rows 15-20 cols O-P)
  vals.push({ range: "'👥 Guest List & RSVPs'!O15", values: [['RSVP STATUS','COUNT']] });
  const RSVPS = ['Confirmed','Declined','Invited','No Response','Waitlisted'];
  RSVPS.forEach((s, i) => {
    vals.push({ range: `'👥 Guest List & RSVPs'!O${16+i}`, values: [[s, `=COUNTIF($E$4:$E$200,"${s}")`]] });
  });

  await valuesBatchUpdate(id, vals, 'guest-values');
  console.log('Guest List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
