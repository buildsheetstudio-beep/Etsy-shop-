'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const EM = sheetMap['Entertainment & Music'];
const RD = sheetMap['Reference Data'];
const S = "'Entertainment & Music'";
const R = "'Reference Data'";

// Helper for standard tab header
function tabHeader(sheetId, numCols, title, subtitle) {
  const reqs = [], data = [];
  reqs.push({ repeatCell: {
    range: gridRange(sheetId,0,300,0,numCols),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});
  reqs.push({ updateSheetProperties: {
    properties: { sheetId, gridProperties:{ frozenRowCount:5 } },
    fields:'gridProperties.frozenRowCount'
  }});
  reqs.push({ mergeCells: { range: gridRange(sheetId,0,1,0,numCols), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(sheetId,0,1,0,numCols),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:44 }, fields:'pixelSize'
  }});
  reqs.push({ mergeCells: { range: gridRange(sheetId,1,2,0,numCols), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(sheetId,1,2,0,numCols),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  [2,3].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(sheetId,ri,ri+1,0,numCols), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(sheetId,ri,ri+1,0,numCols),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:22 }, fields:'pixelSize'
    }});
  });
  return { reqs, data };
}

(async () => {
  const { reqs, data } = tabHeader(EM, 15, 'ENTERTAINMENT & MUSIC', '');

  data.push({ range:`${S}!A1`, values:[['ENTERTAINMENT & MUSIC']] });
  data.push({ range:`${S}!A2`, values:[['Coordinate all entertainment, speakers, toasts, music, and audio/visual equipment']] });
  data.push({ range:`${S}!A3`, values:[['Speaker total time calculates automatically. AV Checklist uses checkboxes. Add/remove rows as needed.']] });
  data.push({ range:`${S}!A4`, values:[['Entertainment log rows 8+. Speaker log rows 42+. Music log rows 76+. AV Checklist rows 110+.']] });

  // ── 8 Summary cards (2 rows of 4) ─────────────────────────────────────────
  const cards = [
    { label:'Total Entertainment Items', formula:`=IFERROR(COUNTA(C8:C40),"—")` },
    { label:'Confirmed Items',            formula:`=IFERROR(COUNTIF(N8:N40,TRUE),"—")` },
    { label:'Pending Items',              formula:`=IFERROR(COUNTA(C8:C40)-COUNTIF(N8:N40,TRUE),"—")` },
    { label:'Entertainment Budget',       formula:`=IFERROR(SUMIF('Budget & Expenses'!$C$11:$C$27,"Entertainment",'Budget & Expenses'!$B$11:$B$27),0)` },
    { label:'Actual Entertainment Cost',  formula:`=IFERROR(SUM(K8:K40),0)` },
    { label:'Speakers Scheduled',         formula:`=IFERROR(COUNTA(C42:C73),"—")` },
    { label:'Music Items Ready',          formula:`=IFERROR(COUNTIF(H76:H108,TRUE),"—")` },
    { label:'AV Equipment Still Needed',  formula:`=IFERROR(COUNTIF(C112:C123,FALSE),"—")` },
  ];

  const cardPos = [
    [4,0,4,4,8],[4,8,11,11,15],
    [5,0,4,4,8],[5,8,11,11,15],
    [6,0,4,4,8],[6,8,11,11,15],  // extra row
    [7,0,4,4,8],[7,8,11,11,15],
  ];
  // Actually fit 4 cards in row 4 and 4 in row 6 with separator at 5
  // Card layout: 15 cols. 4 cards. Each card: label(3)+value(4) = 7.5? Let me use label(3)+value(4-1=3) pairs
  // label(3)+value(3) = 6, × 2 = 12, + separator col = not even. Let's do: 4 cards in 15 cols
  // Each card: label(2) + value(2) = 4 cols. 4 × 4 = 16? Too many.
  // Try: label(2)+value(1) × 4 = 12 cols + 3 spacers = 15. Col layout: 0-1 label, 2 val, 3 space, 4-5 label, 6 val, 7 space, 8-9 label, 10 val, 11 space, 12-13 label, 14 val
  const cardPositions2 = [
    [4, 0,2, 2,4],   // row4: label A-B, val C-D
    [4, 4,6, 6,8],   // row4: label E-F, val G-H
    [4, 8,10,10,12], // row4: label I-J, val K-L
    [4,12,13,13,15], // row4: label M, val N-O
    [6, 0,2, 2,4],
    [6, 4,6, 6,8],
    [6, 8,10,10,12],
    [6,12,13,13,15],
  ];

  cards.forEach((card, i) => {
    const [ri, lc1, lc2, vc1, vc2] = cardPositions2[i];
    const sheetRow = ri + 1;
    const labelCell = String.fromCharCode(65+lc1) + sheetRow;
    const valueCell = String.fromCharCode(65+vc1) + sheetRow;
    data.push({ range:`${S}!${labelCell}`, values:[[card.label]] });
    data.push({ range:`${S}!${valueCell}`, values:[[card.formula]] });

    reqs.push({ mergeCells: { range: gridRange(EM,ri,ri+1,lc1,lc2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(EM,ri,ri+1,lc1,lc2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ mergeCells: { range: gridRange(EM,ri,ri+1,vc1,vc2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(EM,ri,ri+1,vc1,vc2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:13, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:EM, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:38 }, fields:'pixelSize'
    }});
  });

  // Separator row 5 (index 5)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:EM, dimension:'ROWS', startIndex:5, endIndex:6 },
    properties: { pixelSize:8 }, fields:'pixelSize'
  }});

  // ── Entertainment Planner (rows 8-41, section header row 7) ──────────────
  // Section header row 7 (index 7)
  data.push({ range:`${S}!A8`, values:[['ENTERTAINMENT PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(EM,7,8,0,15), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(EM,7,8,0,15),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:EM, dimension:'ROWS', startIndex:7, endIndex:8 },
    properties: { pixelSize:26 }, fields:'pixelSize'
  }});

  // Entertainment headers row 8 (index 8)
  const entHeaders = ['Ent. ID','Type','Name / Activity','Provider / Person','Contact','Start Time','End Time','Duration (min)','Status','Est. Cost','Actual Cost','Payment Status','Equipment Needed','Confirmed?','Notes'];
  data.push({ range:`${S}!A9`, values:[entHeaders] });
  reqs.push({ repeatCell: {
    range: gridRange(EM,8,9,0,15),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:EM, dimension:'ROWS', startIndex:8, endIndex:9 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // 17 entertainment items (rows 10-26, index 9-25)
  const ENT_ITEMS = [
    ['Slideshow','Career Highlights Slideshow','Michael Torres','(951)555-0183','5:45 PM','6:00 PM',15,'Complete',0,0,'Planned','Projector, Laptop, Screen',true,'32-year photo & achievement slideshow'],
    ['Speaker','Welcome Address','Michael Torres (Host)','(951)555-0183','5:30 PM','5:40 PM',10,'Complete',0,0,'Planned','Microphone',true,'Opening and welcome'],
    ['Speaker','Executive Director Remarks','Robert Marsh','(951)555-0110','6:00 PM','6:10 PM',10,'Complete',0,0,'Planned','Microphone',true,'Career overview remarks'],
    ['Toast','Champagne Toast','James Collins (Spouse)','(951)555-0102','6:10 PM','6:15 PM',5,'Complete',0,0,'Planned','None',true,'Personal tribute from family'],
    ['Speaker','Coworker Tribute','Sandra Okafor','(951)555-0111','6:15 PM','6:25 PM',10,'In Progress',0,0,'Planned','Microphone',false,'Tribute from Deputy Director'],
    ['Speaker','Family Tribute','Emily Collins','(951)555-0103','6:25 PM','6:35 PM',10,'Not Started',0,0,'Planned','Microphone',false,'Daughter tribute'],
    ['Award Presentation','Retirement Award Ceremony','Helen Kowalski (Board Chair)','(951)555-0125','6:35 PM','6:45 PM',10,'Not Started',0,0,'Planned','Microphone, Award',false,'Crystal award + framed certificate'],
    ['DJ','DJ Music Set','DJ Smooth','(951)555-9900','7:00 PM','9:00 PM',120,'Complete',450,450,'Paid','Speakers, Turntable',true,'Mix of era music and modern hits'],
    ['Playlist','Arrival Music','DJ Smooth','(951)555-9900','5:00 PM','5:30 PM',30,'Complete',0,0,'Planned','Speakers',true,'Smooth jazz and easy listening'],
    ['Playlist','Dinner Background Music','DJ Smooth','(951)555-9900','6:00 PM','7:00 PM',60,'Complete',0,0,'Planned','Speakers',true,'Low-key dinner ambiance'],
    ['Video Tribute','Retirement Video Montage','Rachel Kim (Comm. Mgr)','(951)555-0123','6:50 PM','7:00 PM',10,'In Progress',0,0,'Planned','Projector, Laptop, Screen',false,'Compiled by communications team'],
    ['Photo Booth','Memory Photo Booth','Party Snap Co.','(951)555-0199','5:30 PM','9:00 PM',210,  'Complete',250,250,'Deposit Paid','Backdrop, Props, Printer',true,'Prints included'],
    ['Toast','Toast by Direct Reports','David Wu','(951)555-0116','7:05 PM','7:10 PM',5,'Not Started',0,0,'Planned','Microphone',false,'Collective tribute from team'],
    ['Games','Memory & Trivia Game','Angela Torres','(951)555-0117','7:30 PM','8:00 PM',30,'Not Started',0,0,'Planned','None',false,'Fun retirement trivia for guests'],
    ['Open Mic','Open Mic Tribute','All Guests','','8:00 PM','8:30 PM',30,'Not Started',0,0,'Planned','Microphone, Timer',false,'1-minute tributes from guests'],
    ['Speaker','Closing Remarks','Michael Torres (Host)','(951)555-0183','8:45 PM','8:55 PM',10,'Not Started',0,0,'Planned','Microphone',false,'Thank-you and closing'],
    ['Live Music','String Quartet (Cocktail Hour)','Riverside String Quartet','(951)555-0200','5:00 PM','6:00 PM',60,'Complete',600,600,'Deposit Paid','Music Stands',true,'Cocktail reception music'],
  ];

  ENT_ITEMS.forEach(([type, name, provider, contact, start, end, duration, status, est, actual, payStatus, equip, confirmed, notes], i) => {
    const row = 10 + i;
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","ENT-"&TEXT(ROW()-9,"000"))`]] });
    data.push({ range:`${S}!B${row}`, values:[[type]] });
    data.push({ range:`${S}!C${row}`, values:[[name]] });
    data.push({ range:`${S}!D${row}`, values:[[provider]] });
    data.push({ range:`${S}!E${row}`, values:[[contact]] });
    data.push({ range:`${S}!F${row}`, values:[[start]] });
    data.push({ range:`${S}!G${row}`, values:[[end]] });
    data.push({ range:`${S}!H${row}`, values:[[duration]] });
    data.push({ range:`${S}!I${row}`, values:[[status]] });
    data.push({ range:`${S}!J${row}`, values:[[est]] });
    data.push({ range:`${S}!K${row}`, values:[[actual]] });
    data.push({ range:`${S}!L${row}`, values:[[payStatus]] });
    data.push({ range:`${S}!M${row}`, values:[[equip]] });
    data.push({ range:`${S}!N${row}`, values:[[confirmed]] });
    data.push({ range:`${S}!O${row}`, values:[[notes]] });
  });

  // Formula-only rows 27-41 (index 26-40)
  for (let row = 27; row <= 41; row++) {
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","ENT-"&TEXT(ROW()-9,"000"))`]] });
  }

  // Alt rows entertainment (index 9-40)
  for (let ri = 9; ri < 41; ri += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(EM,ri,ri+1,0,15),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
  }

  // Confirmed checkbox col N (index 13)
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(EM,9,41,13,14),
    rule: { condition:{ type:'BOOLEAN' }, showCustomUi:true }
  }});
  dvReqs.push({ setDataValidation: {
    range: gridRange(EM,9,41,1,2),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$K$2:$K$13` }] }, showCustomUi:true, strict:false }
  }});

  // ── Speaker & Toast Planner (rows 43-75, section header 42) ──────────────
  data.push({ range:`${S}!A42`, values:[['SPEAKER & TOAST PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(EM,41,42,0,10), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(EM,41,42,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Speech budget cell
  data.push({ range:`${S}!K42`, values:[['Speech Budget (min):']] });
  data.push({ range:`${S}!L42`, values:[[60]] });
  reqs.push({ repeatCell: {
    range: gridRange(EM,41,42,10,11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.white),fontFamily:'Arial'}, horizontalAlignment:'RIGHT' }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
  }});
  reqs.push({ repeatCell: {
    range: gridRange(EM,41,42,11,12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat:{bold:true,fontSize:9,fontFamily:'Arial'} }},
    fields:'userEnteredFormat(backgroundColor,textFormat)'
  }});

  reqs.push({ updateDimensionProperties: {
    range: { sheetId:EM, dimension:'ROWS', startIndex:41, endIndex:42 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // Speaker headers row 43 (index 42)
  const spkHeaders = ['Speaker ID','Speaker Name','Relationship / Role','Segment Type','Topic / Content','Planned Duration (min)','Actual Duration (min)','Scheduled Time','Confirmed?','Notes'];
  data.push({ range:`${S}!A43`, values:[spkHeaders] });
  reqs.push({ repeatCell: {
    range: gridRange(EM,42,43,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:EM, dimension:'ROWS', startIndex:42, endIndex:43 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // 10 speakers/toasts (rows 44-53, index 43-52)
  const SPEAKERS = [
    ['Michael Torres','Host / Organizer','Welcome','Opening welcome and event overview',10,10,'5:30 PM',true,'Opening remarks'],
    ['Robert Marsh','Executive Director','Career Tribute','32 years of service highlight',10,8,'6:00 PM',true,'Prepared speech'],
    ['James Collins','Spouse','Family Tribute','Personal tribute from family',10,null,'6:10 PM',true,'Includes champagne toast'],
    ['Sandra Okafor','Deputy Director','Coworker Tribute','Department tribute',10,null,'6:15 PM',false,'Draft in progress'],
    ['Emily Collins','Daughter','Family Tribute','Daughter personal tribute',10,null,'6:25 PM',false,'Confirm by Jun 15'],
    ['Helen Kowalski','Board Chair','Award Presentation','Crystal award and certificate presentation',10,null,'6:35 PM',false,'Coordinate with awards'],
    ['David Wu','Direct Report','Toast','Collective toast from team',5,null,'7:05 PM',false,'Coordinate group toast'],
    ['Mayor Tom Fischer','City Mayor','Career Tribute','City commendation remarks',7,null,'7:10 PM',false,'Confirm attendance first'],
    ['Barbara Eaton','Retired Colleague','Coworker Tribute','Long-time colleague tribute',5,null,'7:15 PM',false,'Surprise guest tribute'],
    ['Michael Torres','Host / Organizer','Closing Remarks','Thank-you and send-off',10,null,'8:45 PM',false,'Coordinate with DJ for music cue'],
  ];

  SPEAKERS.forEach(([name, role, segType, topic, planned, actual, time, confirmed, notes], i) => {
    const row = 44 + i;
    data.push({ range:`${S}!A${row}`, values:[[`=IF(B${row}="","","SPK-"&TEXT(ROW()-43,"000"))`]] });
    data.push({ range:`${S}!B${row}`, values:[[name]] });
    data.push({ range:`${S}!C${row}`, values:[[role]] });
    data.push({ range:`${S}!D${row}`, values:[[segType]] });
    data.push({ range:`${S}!E${row}`, values:[[topic]] });
    data.push({ range:`${S}!F${row}`, values:[[planned]] });
    if (actual) data.push({ range:`${S}!G${row}`, values:[[actual]] });
    data.push({ range:`${S}!H${row}`, values:[[time]] });
    data.push({ range:`${S}!I${row}`, values:[[confirmed]] });
    data.push({ range:`${S}!J${row}`, values:[[notes]] });
  });

  // Total planned speaking time row 55 (index 54)
  data.push({ range:`${S}!A55`, values:[['TOTAL PLANNED SPEAKING TIME (min):']] });
  data.push({ range:`${S}!F55`, values:[[`=IFERROR(SUM(F44:F73),0)`]] });
  data.push({ range:`${S}!G55`, values:[['Budget:']] });
  data.push({ range:`${S}!H55`, values:[[`=L42`]] });
  data.push({ range:`${S}!I55`, values:[[`=IFERROR(IF(F55>H55,"Over speech budget by "&(F55-H55)&" min","Within budget"),"—")`]] });
  reqs.push({ mergeCells: { range: gridRange(EM,54,55,0,5), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(EM,54,55,0,5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  // CF: total speaking time over budget
  reqs.push({ addConditionalFormatRule: { index:0, rule: {
    ranges:[gridRange(EM,54,55,8,9)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=IFERROR(F55>H55,FALSE)`}] },
      format:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } }
    }
  }}});
  reqs.push({ addConditionalFormatRule: { index:1, rule: {
    ranges:[gridRange(EM,54,55,8,9)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=IFERROR(F55<=H55,FALSE)`}] },
      format:{ backgroundColor:hex(C.success) }
    }
  }}});

  // Confirmed checkbox for speakers col I (index 8)
  dvReqs.push({ setDataValidation: {
    range: gridRange(EM,43,74,8,9),
    rule: { condition:{ type:'BOOLEAN' }, showCustomUi:true }
  }});

  // ── Music Planner (rows 57-75, section header 56) ────────────────────────
  data.push({ range:`${S}!A57`, values:[['MUSIC PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(EM,56,57,0,10), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(EM,56,57,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:EM, dimension:'ROWS', startIndex:56, endIndex:57 },
    properties: { pixelSize:26 }, fields:'pixelSize'
  }});

  const musicHeaders = ['Music ID','Category','Song / Playlist Name','Artist / Source','Requested By','Start Time','Duration (min)','Ready?','Explicit Lyrics Check','Notes'];
  data.push({ range:`${S}!A58`, values:[musicHeaders] });
  reqs.push({ repeatCell: {
    range: gridRange(EM,57,58,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});

  // 23 music items (rows 59-81)
  const MUSIC = [
    ['Arrival Music','Fly Me to the Moon','Frank Sinatra','Michael Torres','5:00 PM',4,true,'Clean','Cocktail arrival'],
    ['Arrival Music','The Way You Look Tonight','Tony Bennett','Michael Torres','5:04 PM',3,true,'Clean',''],
    ['Arrival Music','Come Fly With Me','Frank Sinatra','Michael Torres','5:07 PM',3,true,'Clean',''],
    ['Arrival Music','Smooth Jazz Playlist','DJ Smooth','DJ Smooth','5:10 PM',20,true,'Clean','Background arrival mix'],
    ['Background Music','Canon in D','Pachelbel / String Quartet','Riverside String Quartet','5:30 PM',5,true,'Clean','Cocktail hour'],
    ['Background Music','Air on G String','Bach / String Quartet','Riverside String Quartet','5:35 PM',4,true,'Clean',''],
    ['Dinner Music','All of Me','John Legend','DJ Smooth','6:00 PM',4,true,'Clean',''],
    ['Dinner Music','At Last','Etta James','DJ Smooth','6:04 PM',3,true,'Clean',''],
    ['Dinner Music','Stand By Me','Ben E. King','DJ Smooth','6:07 PM',3,true,'Clean','Patricia\'s favorite'],
    ['Dinner Music','What a Wonderful World','Louis Armstrong','DJ Smooth','6:10 PM',2,true,'Clean',''],
    ['Tribute Song','Wind Beneath My Wings','Bette Midler','DJ Smooth','6:50 PM',5,false,'Clean','Video montage accompaniment'],
    ['Tribute Song','You Raise Me Up','Josh Groban','DJ Smooth','6:55 PM',4,false,'Clean','Award ceremony music'],
    ['Dance Music','September','Earth, Wind & Fire','DJ Smooth','7:00 PM',4,false,'Clean',''],
    ['Dance Music','Don\'t Stop Believin\'','Journey','DJ Smooth','7:04 PM',4,false,'Clean',''],
    ['Dance Music','I Will Survive','Gloria Gaynor','DJ Smooth','7:08 PM',4,false,'Clean','Guest favorite'],
    ['Dance Music','Celebration','Kool & the Gang','DJ Smooth','7:12 PM',3,false,'Clean',''],
    ['Retiree Favorite','Hotel California','Eagles','DJ Smooth','7:30 PM',6,false,'Clean','Patricia\'s request'],
    ['Retiree Favorite','Livin\' on a Prayer','Bon Jovi','DJ Smooth','7:36 PM',4,false,'Clean',''],
    ['Dance Music','Sweet Caroline','Neil Diamond','DJ Smooth','8:00 PM',4,false,'Clean','Guest sing-along moment'],
    ['Dance Music','YMCA','Village People','DJ Smooth','8:04 PM',3,false,'Clean',''],
    ['Closing Song','My Way','Frank Sinatra','DJ Smooth','8:50 PM',4,false,'Clean','Closing dance'],
    ['Closing Song','Time of My Life','Bill Medley & Jennifer Warnes','DJ Smooth','8:54 PM',4,false,'Clean',''],
    ['Other','Emergency Backup Playlist','Spotify Premium','DJ Smooth','',0,true,'Clean','Loaded and ready'],
  ];

  MUSIC.forEach(([cat, song, artist, requestedBy, startTime, duration, ready, explicit, notes], i) => {
    const row = 59 + i;
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","MUS-"&TEXT(ROW()-58,"000"))`]] });
    data.push({ range:`${S}!B${row}`, values:[[cat]] });
    data.push({ range:`${S}!C${row}`, values:[[song]] });
    data.push({ range:`${S}!D${row}`, values:[[artist]] });
    data.push({ range:`${S}!E${row}`, values:[[requestedBy]] });
    if (startTime) data.push({ range:`${S}!F${row}`, values:[[startTime]] });
    data.push({ range:`${S}!G${row}`, values:[[duration]] });
    data.push({ range:`${S}!H${row}`, values:[[ready]] });
    data.push({ range:`${S}!I${row}`, values:[[explicit]] });
    data.push({ range:`${S}!J${row}`, values:[[notes]] });
  });

  // Ready checkbox col H (index 7) for music
  dvReqs.push({ setDataValidation: {
    range: gridRange(EM,58,100,7,8),
    rule: { condition:{ type:'BOOLEAN' }, showCustomUi:true }
  }});

  // ── Audio / Visual Checklist (row 84+) ───────────────────────────────────
  data.push({ range:`${S}!A84`, values:[['AUDIO / VISUAL EQUIPMENT CHECKLIST']] });
  reqs.push({ mergeCells: { range: gridRange(EM,83,84,0,10), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(EM,83,84,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  const AV_ITEMS = [
    ['Microphone','DJ Smooth',true],
    ['Speakers / PA System','Riverside AV',true],
    ['Projector','Riverside AV',true],
    ['Projection Screen','Riverside AV',true],
    ['Laptop (Presentation)','Rachel Kim',false],
    ['Video / HDMI Cable','Michael Torres',false],
    ['Extension Cord (x4)','Event Services Inc.',false],
    ['Playlist Device (DJ)','DJ Smooth',true],
    ['Backup Playlist Device','Michael Torres',false],
    ['Slideshow File (Loaded)','Michael Torres',false],
    ['Tribute Video File (Loaded)','Rachel Kim',false],
    ['Power Strip / Access','Venue Staff',false],
  ];

  // AV headers at row 84 (index 84)
  data.push({ range:`${S}!A85`, values:[['Equipment','Responsible Person','Ready?','Notes']] });
  reqs.push({ repeatCell: {
    range: gridRange(EM,84,85,0,4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  AV_ITEMS.forEach(([item, person, ready], i) => {
    const row = 86 + i;
    data.push({ range:`${S}!A${row}`, values:[[item]] });
    data.push({ range:`${S}!B${row}`, values:[[person]] });
    data.push({ range:`${S}!C${row}`, values:[[ready]] });
  });

  // Ready checkbox col C (index 2) for AV
  dvReqs.push({ setDataValidation: {
    range: gridRange(EM,85,99,2,3),
    rule: { condition:{ type:'BOOLEAN' }, showCustomUi:true }
  }});

  // AV CF: TRUE = sage, FALSE = attention
  reqs.push({ addConditionalFormatRule: { index:2, rule: {
    ranges:[gridRange(EM,85,99,2,3)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=$C86=TRUE`}] },
      format:{ backgroundColor:hex(C.success) }
    }
  }}});
  reqs.push({ addConditionalFormatRule: { index:3, rule: {
    ranges:[gridRange(EM,85,99,2,3)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=$C86=FALSE`}] },
      format:{ backgroundColor:hex(C.attention) }
    }
  }}});

  // Column widths
  const colW = [80,110,180,150,120,70,70,70,90,110,90,100,150,70,160];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:EM, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:EM, tabColorStyle:{ rgbColor:hex(C.primary) } },
    fields:'tabColorStyle'
  }});

  await batchUpdate(id, reqs);
  await batchUpdate(id, dvReqs);
  await valuesBatchUpdate(id, data);
  console.log('Entertainment & Music complete');
})();
