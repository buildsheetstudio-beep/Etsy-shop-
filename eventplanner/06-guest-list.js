'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL = sheetMap['👥 Guest List & RSVP'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'ROWS', startIndex: 3, endIndex: 65 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths: A(30) B(First 100) C(Last 100) D(Event 180) E(Company 150) F(RSVP 100) G(Ticket 110) H(Table# 70) I(Email 200) J(Phone 110) K(Dietary 120) L(Dietary Notes 160) M(Seated 70) N(Badge 70) O(Checked In 80) P(PlusOne 70) Q(Notes 200)
  const colW = [30, 100, 100, 180, 150, 100, 110, 70, 200, 110, 120, 160, 70, 70, 80, 70, 200];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // RSVP summary cols R-U (indices 17-20)
  [130, 120, 120, 120].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: GL, dimension: 'COLUMNS', startIndex: 17+i, endIndex: 18+i }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row — cols A:Q only (leave R:V for RSVP summary header)
  reqs.push({ mergeCells: { range: gridRange(GL, 0, 1, 0, 17), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 1, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.softPurple),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 2 — guest count labels
  reqs.push({ repeatCell: {
    range: gridRange(GL, 1, 2, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 3 — values
  reqs.push({ repeatCell: {
    range: gridRange(GL, 2, 3, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleRose),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(GL, 3, 4, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.softPurple),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows alternating
  for (let r = 4; r < 65; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(GL, r, r+1, 0, 17),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // RSVP summary block header R1
  reqs.push({ mergeCells: { range: gridRange(GL, 0, 1, 17, 22), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(GL, 0, 1, 17, 22),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.antiqueGold),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // RSVP summary rows 2-17 (indices 1-16)
  for (let r = 1; r < 17; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(GL, r, r+1, 17, 22),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
  }

  // Checkbox cols M, N, O, P (indices 12,13,14,15) — center, paleRose
  [12, 13, 14, 15].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(GL, 4, 65, c, c+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose), horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    }});
  });

  // Center align columns A, F, G, H
  [0, 5, 6, 7].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(GL, 3, 65, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(GL, 3, 65, 0, 17),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'guest-format');

  const data = [];
  data.push({ range: `'👥 Guest List & RSVP'!A1`, values: [['👥 Guest List & RSVP — All Events']] });
  data.push({ range: `'👥 Guest List & RSVP'!R1`, values: [['📊 RSVP Summary by Event']] });

  // Summary labels row 2
  data.push({ range: `'👥 Guest List & RSVP'!A2:Q2`, values: [['Total','Confirmed','Declined','Pending','Waitlisted','No Show','','','','Checked In','','','','','','','']] });
  // Summary values row 3
  data.push({ range: `'👥 Guest List & RSVP'!A3:F3`, values: [[
    '=COUNTA(B5:B65)',
    '=COUNTIF(F5:F65,"Confirmed")',
    '=COUNTIF(F5:F65,"Declined")',
    '=COUNTIF(F5:F65,"Pending")',
    '=COUNTIF(F5:F65,"Waitlisted")',
    '=COUNTIF(F5:F65,"No Show")',
  ]] });
  data.push({ range: `'👥 Guest List & RSVP'!J3`, values: [['=COUNTIF(O5:O65,TRUE)']] });

  data.push({ range: `'👥 Guest List & RSVP'!A4:Q4`, values: [['#','First Name','Last Name','Event','Company','RSVP Status','Ticket Type','Table #','Email','Phone','Dietary Req','Dietary Notes','Seated','Badge Printed','Checked In','Plus One','Notes']] });

  // RSVP summary header
  data.push({ range: `'👥 Guest List & RSVP'!R2:V2`, values: [['Event','Invited','Confirmed','Declined','Pending']] });

  const guests = [
    [1,'James','Hartley','Tech Summit 2025','InnovateTech Ltd','Confirmed','VIP',1,'j.hartley@innovatetech.com','0412 000 001','None','',true,true,false,false,'Keynote speaker'],
    [2,'Priya','Sharma','Tech Summit 2025','DataFlow Systems','Confirmed','General Admission',2,'p.sharma@dataflow.com','0412 000 002','Vegetarian','No dairy either',true,false,false,false,''],
    [3,'Oliver','Chen','Tech Summit 2025','CloudBase Co.','Confirmed','Speaker',1,'o.chen@cloudbase.com','0412 000 003','None','',true,true,false,false,'Panel moderator'],
    [4,'Maya','Robinson','Tech Summit 2025','StartupHub','Confirmed','General Admission',3,'m.robinson@startuphub.com','0412 000 004','Gluten-Free','',false,false,false,true,'Bringing partner'],
    [5,'Ethan','Kowalski','Tech Summit 2025','FinTech Partners','Pending','Early Bird',3,'e.kowalski@fintechp.com','0412 000 005','None','',false,false,false,false,''],
    [6,'Sofia','Nguyen','Tech Summit 2025','GreenTech Inc','Confirmed','General Admission',4,'s.nguyen@greentech.com','0412 000 006','Vegan','',false,false,false,false,''],
    [7,'Liam','Patel','Tech Summit 2025','Digital Edge','Declined','','','l.patel@digitaledge.com','0412 000 007','','',false,false,false,false,'Travel conflict'],
    [8,'Emma','Walsh','Charity Gala 2025','Walsh Philanthropy','Confirmed','VIP',5,'e.walsh@walshphil.com','0412 000 008','None','',true,true,false,false,'Major donor'],
    [9,'Noah','Anderson','Charity Gala 2025','Anderson Group','Confirmed','Sponsor',5,'n.anderson@andersongrp.com','0412 000 009','Nut Allergy','Severe — EpiPen carried',true,false,false,false,''],
    [10,'Ava','Thompson','Charity Gala 2025','TBD','Confirmed','General Admission',6,'a.thompson@email.com','0412 000 010','Dairy-Free','',false,false,false,true,'Bringing partner'],
    [11,'William','Brown','Charity Gala 2025','Brown & Associates','Pending','','','w.brown@brownassoc.com','0412 000 011','None','',false,false,false,false,'Awaiting RSVP'],
    [12,'Isabella','Garcia','Charity Gala 2025','Garcia Foundation','Confirmed','VIP',5,'i.garcia@garciafdn.com','0412 000 012','Gluten-Free','',false,false,false,false,''],
    [13,'Mason','Lee','Annual Awards Night','Internal','Confirmed','Complimentary',7,'m.lee@company.com','0412 000 013','None','',false,false,false,false,'Employee of Year nominee'],
    [14,'Charlotte','Davis','Annual Awards Night','Internal','Confirmed','Complimentary',7,'c.davis@company.com','0412 000 014','Vegetarian','',false,false,false,false,''],
    [15,'Lucas','Martinez','Annual Awards Night','Internal','Pending','','','l.martinez@company.com','0412 000 015','None','',false,false,false,false,''],
    [16,'Amelia','Wilson','Holiday Party 2025','Internal','Confirmed','Complimentary',8,'a.wilson@company.com','0412 000 016','None','',false,false,false,true,''],
    [17,'Henry','Taylor','Holiday Party 2025','Internal','Confirmed','Complimentary',8,'h.taylor@company.com','0412 000 017','Halal','',false,false,false,false,''],
    [18,'Evelyn','White','Investor Day 2025','WhiteCap Ventures','Confirmed','General Admission',''  ,'e.white@whitecap.com','0412 000 018','None','',false,false,false,false,'Lead investor'],
    [19,'Alexander','Harris','Investor Day 2025','Horizon Capital','Confirmed','General Admission','','a.harris@horizoncap.com','0412 000 019','None','',false,false,false,false,''],
    [20,'Scarlett','Clark','5-Year Anniversary Dinner','Internal','Confirmed','Complimentary',9,'s.clark@company.com','0412 000 020','Vegan','',false,false,false,false,'Founding team member'],
  ];

  data.push({ range: `'👥 Guest List & RSVP'!A5:Q24`, values: guests });

  // RSVP summary per event (rows 3-14 = indices 2-13)
  const eventNames = ['Tech Summit 2025','Charity Gala 2025','Annual Awards Night','Holiday Party 2025','Investor Day 2025','5-Year Anniversary Dinner'];
  const summaryRows = eventNames.map((ev, i) => {
    const r = 5 + i; // data starts row 5
    return [
      ev,
      `=COUNTIF('👥 Guest List & RSVP'!D$5:D$65,"${ev}")`,
      `=COUNTIFS('👥 Guest List & RSVP'!D$5:D$65,"${ev}",'👥 Guest List & RSVP'!F$5:F$65,"Confirmed")`,
      `=COUNTIFS('👥 Guest List & RSVP'!D$5:D$65,"${ev}",'👥 Guest List & RSVP'!F$5:F$65,"Declined")`,
      `=COUNTIFS('👥 Guest List & RSVP'!D$5:D$65,"${ev}",'👥 Guest List & RSVP'!F$5:F$65,"Pending")`,
    ];
  });

  data.push({ range: `'👥 Guest List & RSVP'!R3:V8`, values: summaryRows });

  await valuesBatchUpdate(id, data, 'guest-values');
  console.log('Guest List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
