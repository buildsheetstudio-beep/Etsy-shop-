'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const EL = sheetMap['📋 Events Log'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: EL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: EL, dimension: 'ROWS', startIndex: 1, endIndex: 65 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths
  const colWidths = [30, 220, 120, 100, 100, 110, 110, 110, 90, 120, 90, 110, 90, 130, 90, 140, 140];
  colWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: EL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row — A1:Q1 merged, deep blush
  reqs.push({ mergeCells: { range: gridRange(EL, 0, 1, 0, 17), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(EL, 0, 1, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Header row (row 2) — rose bg
  reqs.push({ repeatCell: {
    range: gridRange(EL, 1, 2, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows alternating bg (champagne / ivory)
  for (let r = 2; r < 65; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(EL, r, r+1, 0, 17),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // # column center-align
  reqs.push({ repeatCell: { range: gridRange(EL, 1, 65, 0, 1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  // Date/numeric columns center
  [4,5,6,7,8,10,12,14].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(EL, 1, 65, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Date format cols E,F,G (indices 4,5,6)
  [4, 5, 6].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(EL, 2, 65, c, c+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  });

  // Currency format — Venue Cost K (col 10), Actual Spend M (col 12), Revenue N (col 13)
  [10, 12, 13].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(EL, 2, 65, c, c+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  });

  // Formula cells (Confirmed Guests I col8, Actual Spend M col12) — pale rose bg
  [8, 12].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(EL, 2, 65, c, c+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  });

  // Borders on data area
  reqs.push({ updateBorders: {
    range: gridRange(EL, 1, 65, 0, 17),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'events-log-format');

  // Values
  const data = [];
  data.push({ range: `'📋 Events Log'!A1`, values: [['📋 Events Log — All Events at a Glance']] });
  data.push({ range: `'📋 Events Log'!A2:Q2`, values: [['#','Event Name','Category','Type','Start Date','End Date','Setup Date','Strike Date','Confirmed Guests','Venue','Venue Cost','Status','Actual Spend','Revenue','Priority','Created By','Notes']] });

  const events = [
    [1,'Tech Summit 2025','Conference','In-Person','2025-09-15','2025-09-15','2025-09-14','2025-09-16','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B3,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Grand Convention Centre',8500,'Confirmed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B3,\'💰 Event Budget\'!F$19:F$200),0)',12500,'High','Sarah Johnson','Annual tech conference'],
    [2,'Charity Gala 2025','Fundraiser','In-Person','2025-10-22','2025-10-22','2025-10-22','2025-10-23','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B4,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Riverside Ballroom',12000,'Confirmed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B4,\'💰 Event Budget\'!F$19:F$200),0)',45000,'High','Sarah Johnson','Black tie fundraiser'],
    [3,'Product Launch — Nova X','Product Launch','Hybrid','2025-08-05','2025-08-05','2025-08-05','2025-08-05','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B5,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Studio 12 + Online',3200,'Completed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B5,\'💰 Event Budget\'!F$19:F$200),0)',0,'High','Mark Davis','New product reveal'],
    [4,'Team Building Day','Team Building','In-Person','2025-07-18','2025-07-18','2025-07-18','2025-07-18','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B6,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Outdoor Adventure Park',2800,'Confirmed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B6,\'💰 Event Budget\'!F$19:F$200),0)',0,'Medium','Lisa Chen','Q3 team activity'],
    [5,'Annual Awards Night','Corporate','In-Person','2025-11-28','2025-11-28','2025-11-28','2025-11-29','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B7,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','The Grand Hotel',9500,'Planning','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B7,\'💰 Event Budget\'!F$19:F$200),0)',0,'High','Sarah Johnson','Employee recognition'],
    [6,'Marketing Workshop','Workshop','In-Person','2025-08-20','2025-08-20','2025-08-20','2025-08-20','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B8,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Training Centre B',800,'Completed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B8,\'💰 Event Budget\'!F$19:F$200),0)',2400,'Medium','Mark Davis','Digital marketing training'],
    [7,'Virtual Workshop — AI Trends','Workshop','Virtual','2025-09-03','2025-09-03','','','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B9,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Online / Zoom',0,'Completed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B9,\'💰 Event Budget\'!F$19:F$200),0)',3800,'Medium','Lisa Chen','Online professional development'],
    [8,'Networking Drinks — Q4','Social','In-Person','2025-10-09','2025-10-09','2025-10-09','2025-10-09','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B10,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Rooftop Bar & Lounge',1800,'Confirmed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B10,\'💰 Event Budget\'!F$19:F$200),0)',0,'Low','Mark Davis','Quarterly networking mixer'],
    [9,'Holiday Party 2025','Holiday Party','In-Person','2025-12-12','2025-12-12','2025-12-12','2025-12-12','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B11,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','The Loft Event Space',5500,'Planning','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B11,\'💰 Event Budget\'!F$19:F$200),0)',0,'High','Sarah Johnson','End of year celebration'],
    [10,'CEO Birthday Celebration','Birthday','In-Person','2025-07-25','2025-07-25','2025-07-25','2025-07-25','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B12,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Private Dining Room',2200,'Completed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B12,\'💰 Event Budget\'!F$19:F$200),0)',0,'Medium','Lisa Chen','Surprise 50th birthday'],
    [11,'Investor Day 2025','Conference','Hybrid','2025-10-30','2025-10-30','2025-10-30','2025-10-30','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B13,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','HQ Boardroom + Online',1500,'Planning','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B13,\'💰 Event Budget\'!F$19:F$200),0)',0,'High','Sarah Johnson','Annual investor briefing'],
    [12,'5-Year Anniversary Dinner','Anniversary','In-Person','2025-11-07','2025-11-07','2025-11-07','2025-11-07','=IFERROR(COUNTIFS(\'👥 Guest List & RSVP\'!C$3:C$65,B14,\'👥 Guest List & RSVP\'!F$3:F$65,"Confirmed"),0)','Harbour View Restaurant',4800,'Confirmed','=IFERROR(SUMIF(\'💰 Event Budget\'!B$19:B$200,B14,\'💰 Event Budget\'!F$19:F$200),0)',0,'High','Mark Davis','Company 5th anniversary'],
  ];

  data.push({ range: `'📋 Events Log'!A3:Q14`, values: events });

  await valuesBatchUpdate(id, data, 'events-log-values');
  console.log('Events Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
