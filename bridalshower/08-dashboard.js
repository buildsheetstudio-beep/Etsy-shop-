'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DASH = sheetMap['🥂 Dashboard'];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🥂 Dashboard'";

  // Column widths
  const colW = [170,180,30,170,180,30,170,180];
  colW.forEach((w,i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Row 0: title banner
  reqs.push({ mergeCells: { range: gridRange(DASH,0,1,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,0,1,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 60 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['🥂 Ultimate Bridal Shower Planner']] });

  // Row 1: subtitle
  reqs.push({ mergeCells: { range: gridRange(DASH,1,2,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,2,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.gold),
      textFormat: { foregroundColor: hex(C.white), bold: false, fontSize: 10, italic: true },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['A beautifully organised celebration for your bride-to-be']] });

  // ── PARTY INFO SECTION (rows 2-9, 0-indexed) ──
  reqs.push({ mergeCells: { range: gridRange(DASH,2,3,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,2,3,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A3`, values: [['PARTY DETAILS']] });

  const partyInfo = [
    ['Bride\'s Name',  'Charlotte Harrison',   'Venue',        'La Maison Garden Room'],
    ['Event Date',     '4 October 2026',        'Start Time',   '11:00 AM'],
    ['Location',       'London, UK',            'End Time',     '5:30 PM'],
    ['Organiser',      'Sophie Clarke (MOH)',   'Dress Code',   'Garden Party Chic'],
    ['Theme',          'Soft Florals & Brunch', 'Max Guests',   '20'],
  ];
  for (let i = 0; i < partyInfo.length; i++) {
    const ri  = 3 + i;
    const row = 4 + i;
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 0, 8),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    // Label cols bold
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 0, 1),
      cell: { userEnteredFormat: {
        textFormat: { bold: true, foregroundColor: hex(C.muted), fontSize: 9 },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(textFormat,backgroundColor)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 3, 4),
      cell: { userEnteredFormat: {
        textFormat: { bold: true, foregroundColor: hex(C.muted), fontSize: 9 },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(textFormat,backgroundColor)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [partyInfo[i]] });
  }

  // Divider (0-indexed 8)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,8,9,0,8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // ── GUEST STATS section (row 9-15) ──
  reqs.push({ mergeCells: { range: gridRange(DASH,9,10,0,4), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,9,10,0,4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A10`, values: [['GUEST SUMMARY']] });

  reqs.push({ mergeCells: { range: gridRange(DASH,9,10,4,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,9,10,4,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!E10`, values: [['BUDGET SUMMARY']] });

  // Guest KPI rows (0-indexed 10-15)
  const guestKPIs = [
    ['Total Invited',    `=COUNTA('👰 Guest List & Seating'!A4:A23)`],
    ['Confirmed',        `=COUNTIF('👰 Guest List & Seating'!E4:E23,"Confirmed")`],
    ['Declined',         `=COUNTIF('👰 Guest List & Seating'!E4:E23,"Declined")`],
    ['No Response',      `=COUNTIF('👰 Guest List & Seating'!E4:E23,"No Response")`],
    ['Plus Ones',        `=COUNTIF('👰 Guest List & Seating'!H4:H23,"Yes")`],
    ['Thank Yous Sent',  `=COUNTIF('🎁 Gift Tracker & Wishlist'!H23:H60,"Thank You Sent")`],
  ];
  const budgetKPIs = [
    ['Total Budget',   `=IFERROR('💰 Budget & Expenses'!B16,"")`],
    ['Total Spent',    `=IFERROR('💰 Budget & Expenses'!C16,"")`],
    ['Remaining',      `=IFERROR('💰 Budget & Expenses'!D16,"")`],
    ['% Used',         `=IFERROR('💰 Budget & Expenses'!E16,"")`],
    ['Gifts Received', `=COUNTA('🎁 Gift Tracker & Wishlist'!A23:A60)`],
    ['Tasks Done',     `=COUNTIF('💍 Venue, Vendors & Bridesmaids'!F20:F37,"Done")`],
  ];
  for (let i = 0; i < 6; i++) {
    const ri  = 10 + i;
    const row = 11 + i;
    const bg  = i % 2 === 0 ? C.parchment : C.ivory;
    // Guest cols A-D
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 0, 4),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 0, 1),
      cell: { userEnteredFormat: {
        textFormat: { bold: true, foregroundColor: hex(C.muted), fontSize: 9 },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(textFormat,backgroundColor)',
    }});
    // Budget cols E-H
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 4, 8),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: {
        textFormat: { bold: true, foregroundColor: hex(C.muted), fontSize: 9 },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(textFormat,backgroundColor)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [[guestKPIs[i][0], guestKPIs[i][1]]] });
    vals.push({ range: `${TAB}!E${row}`, values: [[budgetKPIs[i][0], budgetKPIs[i][1]]] });
  }
  // Currency format for budget values
  reqs.push({ repeatCell: {
    range: gridRange(DASH,10,13,5,6),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }}},
    fields: 'userEnteredFormat(numberFormat)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,13,14,5,6),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' }}},
    fields: 'userEnteredFormat(numberFormat)',
  }});

  // Divider (0-indexed 16)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 16, endIndex: 17 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,16,17,0,8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // ── SPARKLINE section (rows 17-18) ──
  reqs.push({ mergeCells: { range: gridRange(DASH,17,18,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,17,18,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A18`, values: [['PROGRESS AT A GLANCE']] });

  // Sparkline rows (0-indexed 18)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,18,19,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.parchment),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A19`, values: [[
    'RSVPs Confirmed',
    `=SPARKLINE(COUNTIF('👰 Guest List & Seating'!E4:E23,"Confirmed")/COUNTA('👰 Guest List & Seating'!A4:A23),{"charttype","bar";"color1","#8B5E6A";"max",1})`,
    '',
    'Budget Used',
    `=SPARKLINE(IFERROR('💰 Budget & Expenses'!C16/'💰 Budget & Expenses'!B16,0),{"charttype","bar";"color1","#C9A96E";"max",1})`,
    '',
    'Tasks Complete',
    `=SPARKLINE(COUNTIF('💍 Venue, Vendors & Bridesmaids'!F20:F37,"Done")/COUNTA('💍 Venue, Vendors & Bridesmaids'!C20:C37),{"charttype","bar";"color1","#5A7A5A";"max",1})`,
  ]]});

  // Divider (0-indexed 19)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 19, endIndex: 20 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,19,20,0,8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // ── UPCOMING TASKS (row 20 onwards) ──
  reqs.push({ mergeCells: { range: gridRange(DASH,20,21,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,20,21,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A21`, values: [['OPEN TASKS (NOT STARTED / IN PROGRESS)']] });

  // Col headers for task list (0-indexed 21)
  const taskHdrs = ['NAME','TASK','CATEGORY','PRIORITY','STATUS'];
  reqs.push({ repeatCell: {
    range: gridRange(DASH,21,22,0,5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A22`, values: [taskHdrs] });

  // Task list using FILTER (0-indexed 22)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,22,23,0,5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.ivory),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A23`, values: [[
    `=IFERROR(FILTER('💍 Venue, Vendors & Bridesmaids'!A20:A37,'💍 Venue, Vendors & Bridesmaids'!F20:F37<>"Done"),"No open tasks")`,
    `=IFERROR(FILTER('💍 Venue, Vendors & Bridesmaids'!C20:C37,'💍 Venue, Vendors & Bridesmaids'!F20:F37<>"Done"),"")`,
    `=IFERROR(FILTER('💍 Venue, Vendors & Bridesmaids'!D20:D37,'💍 Venue, Vendors & Bridesmaids'!F20:F37<>"Done"),"")`,
    `=IFERROR(FILTER('💍 Venue, Vendors & Bridesmaids'!E20:E37,'💍 Venue, Vendors & Bridesmaids'!F20:F37<>"Done"),"")`,
    `=IFERROR(FILTER('💍 Venue, Vendors & Bridesmaids'!F20:F37,'💍 Venue, Vendors & Bridesmaids'!F20:F37<>"Done"),"")`,
  ]]});

  await batchUpdate(id, reqs, 'dashboard-format');
  await valuesBatchUpdate(id, vals, 'dashboard-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
