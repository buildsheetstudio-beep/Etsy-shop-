'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PS = sheetMap['Party Setup'];
const S = "'Party Setup'";

(async () => {
  const reqs = [];
  const data = [];

  // ── Background fill first ──────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(PS,0,100,0,8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});

  // ── Title row (row 0) full width ───────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(PS,0,1,0,8), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PS,0,1,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:18, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:PS, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:48 }, fields:'pixelSize'
  }});

  // ── Subtitle row (row 1) full width ───────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(PS,1,2,0,8), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PS,1,2,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:PS, dimension:'ROWS', startIndex:1, endIndex:2 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // ── Row 2: PARTY DETAILS (cols A-C) | PLANNING SUMMARY (cols D-H) ─────────
  // Two separate merges — no full-row merge on row 2
  reqs.push({ mergeCells: { range: gridRange(PS,2,3,0,3), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PS,2,3,0,3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  reqs.push({ mergeCells: { range: gridRange(PS,2,3,3,8), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PS,2,3,3,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:PS, dimension:'ROWS', startIndex:2, endIndex:3 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // ── Party detail fields rows 3-21 (A col=label, B col=input) ─────────────
  const fields = [
    ['Retiree Name', 'Patricia Collins'],
    ['Job Title / Role', 'Operations Director'],
    ['Organization', 'Northfield Community Services'],
    ['Years of Service', '32'],
    ['Retirement Date', 'Jun 30, 2026'],
    ['Party Date', 'Jun 27, 2026'],
    ['Start Time', '5:30 PM'],
    ['End Time', '9:00 PM'],
    ['Event Type', 'Workplace & Family Celebration'],
    ['Theme', 'Celebrating 32 Years of Service'],
    ['Venue / Location', 'Lakeside Event Hall'],
    ['Address', '4200 Lakeside Drive, Riverside CA 92501'],
    ['Host / Organizer', 'Michael Torres'],
    ['Contact Number', '(951) 555-0183'],
    ['Total Budget', '4500'],
    ['Guest Capacity', '100'],
    ['RSVP Deadline', 'Jun 10, 2026'],
    ['Surprise Party?', 'No'],
    ['Notes', 'Coordinate with Dept. Heads for speeches. Confirm A/V by June 1.'],
  ];

  fields.forEach(([label, val], i) => {
    data.push({ range:`${S}!A${4+i}`, values:[[label]] });
    data.push({ range:`${S}!B${4+i}`, values:[[val]] });
  });

  // Label formatting col A rows 3-21 (index 3-21)
  reqs.push({ repeatCell: {
    range: gridRange(PS,3,22,0,1),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.mainText), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Input formatting col B rows 3-21
  reqs.push({ repeatCell: {
    range: gridRange(PS,3,22,1,2),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
      textFormat: { fontSize:10, foregroundColor:hex(C.mainText), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Col C spacer
  reqs.push({ repeatCell: {
    range: gridRange(PS,3,22,2,3),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});

  // ── 8 summary cards: 4 rows × 2 cols (D+E label, F+G+H value) ────────────
  // Cards at row indices 3,5,7,9 (group1) then 12,14,16,18 (group2)
  // Total Budget uses B18 (row index 17 → sheet row 18 → zero-indexed field index 14 → B17+1 = B18)
  // Field 14 (0-indexed) = Total Budget → sheet row 4+14 = row 18
  // Guest Capacity field 15 → row 19
  // RSVP Deadline field 16 → row 20

  const cards = [
    { label:'Days Until Party',             formula:`=IFERROR(MAX(0,DATEVALUE("Jun 27, 2026")-TODAY()),"—")` },
    { label:'Total Budget',                  formula:`=IFERROR(B18,"—")` },
    { label:'Confirmed Guests',              formula:`=IFERROR(COUNTIF('Guest List & RSVP'!$I$6:$I$305,"Confirmed"),"—")` },
    { label:'Guest Capacity Remaining',      formula:`=IFERROR(B19-COUNTIF('Guest List & RSVP'!$I$6:$I$305,"Confirmed"),"—")` },
    { label:'Amount Spent',                  formula:`=IFERROR(SUM('Budget & Expenses'!$I$31:$I$329),0)` },
    { label:'Remaining Budget',              formula:`=IFERROR(B18-SUM('Budget & Expenses'!$I$31:$I$329),"—")` },
    { label:'Cost per Confirmed Guest',      formula:`=IFERROR(IF(COUNTIF('Guest List & RSVP'!$I$6:$I$305,"Confirmed")=0,"—",SUM('Budget & Expenses'!$I$31:$I$329)/COUNTIF('Guest List & RSVP'!$I$6:$I$305,"Confirmed")),"—")` },
    { label:'Invitations Awaiting Response', formula:`=IFERROR(COUNTIF('Guest List & RSVP'!$H$6:$H$305,"Invitation Sent"),"—")` },
  ];

  // Row indices: 3,5,7,9 | separator 10,11 | 12,14,16,18
  const cardRows = [3,5,7,9,12,14,16,18];

  cardRows.forEach((ri, i) => {
    data.push({ range:`${S}!D${ri+1}`, values:[[cards[i].label]] });
    data.push({ range:`${S}!F${ri+1}`, values:[[cards[i].formula]] });

    // Label: cols D-E merged
    reqs.push({ mergeCells: { range: gridRange(PS,ri,ri+1,3,5), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(PS,ri,ri+1,3,5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});

    // Value: cols F-H merged
    reqs.push({ mergeCells: { range: gridRange(PS,ri,ri+1,5,8), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(PS,ri,ri+1,5,8),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:14, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});

    reqs.push({ updateDimensionProperties: {
      range: { sheetId:PS, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:38 }, fields:'pixelSize'
    }});
  });

  // Separator rows between card groups (indices 10-11)
  [10,11].forEach(ri => {
    reqs.push({ repeatCell: {
      range: gridRange(PS,ri,ri+1,3,8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:PS, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:6 }, fields:'pixelSize'
    }});
  });

  // Separator: rows between individual cards (indices 4,6,8,13,15,17)
  [4,6,8,13,15,17].forEach(ri => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:PS, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:6 }, fields:'pixelSize'
    }});
  });

  // ── Column widths ──────────────────────────────────────────────────────────
  const colWidths = [200, 280, 12, 170, 10, 160, 10, 120];
  colWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:PS, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  // ── Freeze rows 1:3 ────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:PS, gridProperties: { frozenRowCount:3 } },
    fields:'gridProperties.frozenRowCount'
  }});

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:PS, tabColorStyle: { rgbColor: hex(C.primary) } },
    fields:'tabColorStyle'
  }});

  // Values first pass
  data.push({ range:`${S}!A1`, values:[['ULTIMATE RETIREMENT PARTY PLANNER']] });
  data.push({ range:`${S}!A2`, values:[['Plan every detail of a memorable retirement celebration — guests, budget, entertainment, and more']] });
  data.push({ range:`${S}!A3`, values:[['PARTY DETAILS']] });
  data.push({ range:`${S}!D3`, values:[['PLANNING SUMMARY']] });

  await batchUpdate(id, reqs);
  await valuesBatchUpdate(id, data);
  console.log('Party Setup complete');
})();
