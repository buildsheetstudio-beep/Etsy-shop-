'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IT = sheetMap['Invitation Tracker'];
const RD = sheetMap['Reference Data'];
const S = "'Invitation Tracker'";
const R = "'Reference Data'";
const G = "'Guest List & RSVP'";

// 14 columns A-N
const HEADERS = ['Guest ID','Guest Name','Household / Org','Invitation Method','Invitation Ready?','Sent Date','RSVP Deadline','Response Status','Follow-Up Needed?','Follow-Up Date','Reminder Sent?','Days Since Sent','Contact Details','Notes'];

// 28 sample rows [guestId, method, invReady, sentDate, followUpDate, reminderSent, notes]
const ROWS = [
  ['GST-001','Printed Invitation',true,'May 15, 2026','','FALSE','Head of honor — personal delivery'],
  ['GST-002','Printed Invitation',true,'May 15, 2026','','FALSE',''],
  ['GST-003','Email',true,'May 16, 2026','','FALSE',''],
  ['GST-004','Email',true,'May 16, 2026','','FALSE',''],
  ['GST-005','Printed Invitation',true,'May 18, 2026','','FALSE','Official organization invite'],
  ['GST-006','Email',true,'May 18, 2026','','FALSE',''],
  ['GST-007','Email',true,'May 18, 2026','','FALSE',''],
  ['GST-008','Email',true,'May 18, 2026','','FALSE',''],
  ['GST-009','Email',true,'May 18, 2026','','FALSE',''],
  ['GST-010','Email',true,'May 18, 2026','','FALSE',''],
  ['GST-011','Email',true,'May 19, 2026','','FALSE',''],
  ['GST-012','Email',true,'May 19, 2026','','FALSE',''],
  ['GST-013','Email',true,'May 19, 2026','Jun 10, 2026','FALSE','Follow-up call placed May 28'],
  ['GST-014','Email',true,'May 19, 2026','','FALSE',''],
  ['GST-015','Email',true,'May 19, 2026','','FALSE',''],
  ['GST-016','Text Message',false,'','Jun 10, 2026','FALSE','Not yet sent — check email address'],
  ['GST-017','Text Message',true,'May 20, 2026','Jun 10, 2026','TRUE','Second reminder sent June 3'],
  ['GST-018','Email',true,'May 20, 2026','','FALSE',''],
  ['GST-019','Workplace Announcement',true,'May 21, 2026','','FALSE','Declined verbally'],
  ['GST-020','Printed Invitation',true,'May 22, 2026','','FALSE',''],
  ['GST-021','Printed Invitation',true,'May 22, 2026','','FALSE','Personal delivery to City Hall'],
  ['GST-022','Online Invitation',true,'May 22, 2026','Jun 10, 2026','FALSE',''],
  ['GST-023','Phone Call',true,'May 23, 2026','','FALSE',''],
  ['GST-024','Email',true,'May 23, 2026','','FALSE',''],
  ['GST-025','Email',true,'May 23, 2026','','FALSE','Accessibility check completed'],
  ['GST-026','Email',true,'May 23, 2026','Jun 10, 2026','TRUE','Reminder sent June 2'],
  ['GST-027','Email',true,'May 24, 2026','Jun 10, 2026','FALSE',''],
  ['GST-028','In Person',true,'May 24, 2026','','FALSE',''],
];

(async () => {
  const reqs = [];
  const data = [];

  // Background
  reqs.push({ repeatCell: {
    range: gridRange(IT,0,300,0,14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});

  // Freeze first
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:IT, gridProperties: { frozenRowCount:5 } },
    fields:'gridProperties.frozenRowCount'
  }});

  // Title row 0
  data.push({ range:`${S}!A1`, values:[['INVITATION TRACKER']] });
  reqs.push({ mergeCells: { range: gridRange(IT,0,1,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(IT,0,1,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:IT, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:44 }, fields:'pixelSize'
  }});

  // Subtitle row 1
  data.push({ range:`${S}!A2`, values:[['Track invitation delivery and follow-up for every guest']] });
  reqs.push({ mergeCells: { range: gridRange(IT,1,2,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(IT,1,2,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Info rows 2-3
  data.push({ range:`${S}!A3`, values:[['Guest Name, Household, Response Status, and Contact Details auto-fill from Guest List using Guest ID']] });
  data.push({ range:`${S}!A4`, values:[['Follow-Up Needed: auto-flags Yes when invitation sent, response is Maybe or No Response, and deadline has passed']] });
  [2,3].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(IT,ri,ri+1,0,14), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(IT,ri,ri+1,0,14),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:IT, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:22 }, fields:'pixelSize'
    }});
  });

  // Header row 4
  data.push({ range:`${S}!A5`, values:[HEADERS] });
  reqs.push({ repeatCell: {
    range: gridRange(IT,4,5,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:IT, dimension:'ROWS', startIndex:4, endIndex:5 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // Data rows
  ROWS.forEach(([gid, method, invReady, sentDate, followUpDate, reminderSent, notes], i) => {
    const row = 6 + i;
    // A: Guest ID
    data.push({ range:`${S}!A${row}`, values:[[gid]] });
    // B: Guest Name via VLOOKUP
    data.push({ range:`${S}!B${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,3,FALSE)),"")`]] });
    // C: Household via VLOOKUP
    data.push({ range:`${S}!C${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,2,FALSE)),"")`]] });
    // D: Method
    data.push({ range:`${S}!D${row}`, values:[[method]] });
    // E: Invitation Ready checkbox
    data.push({ range:`${S}!E${row}`, values:[[invReady]] });
    // F: Sent date
    if (sentDate) data.push({ range:`${S}!F${row}`, values:[[sentDate]] });
    // G: RSVP Deadline from Party Setup
    data.push({ range:`${S}!G${row}`, values:[[`=IFERROR(IF(A${row}="","",IFERROR('Party Setup'!B20,"")),"")` ]] });
    // H: Response Status via VLOOKUP col 9 (RSVP Status)
    data.push({ range:`${S}!H${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,9,FALSE)),"")`]] });
    // I: Follow-Up Needed
    data.push({ range:`${S}!I${row}`, values:[[`=IFERROR(IF(A${row}="","",IF(AND(E${row}=TRUE,F${row}<>"",OR(H${row}="No Response",H${row}="Maybe"),H${row}<>"Declined"),"Yes","No")),"")`]] });
    // J: Follow-Up Date
    if (followUpDate) data.push({ range:`${S}!J${row}`, values:[[followUpDate]] });
    // K: Reminder Sent checkbox
    data.push({ range:`${S}!K${row}`, values:[[reminderSent === 'TRUE']] });
    // L: Days Since Sent
    data.push({ range:`${S}!L${row}`, values:[[`=IFERROR(IF(F${row}="","",TODAY()-F${row}),"")`]] });
    // M: Contact Details
    data.push({ range:`${S}!M${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,6,FALSE)&" | "&VLOOKUP(A${row},${G}!$A$6:$R$305,7,FALSE)),"")`]] });
    // N: Notes
    data.push({ range:`${S}!N${row}`, values:[[notes]] });
  });

  // Formula-only rows for remaining
  for (let row = 34; row <= 305; row++) {
    data.push({ range:`${S}!B${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,3,FALSE)),"")`]] });
    data.push({ range:`${S}!C${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,2,FALSE)),"")`]] });
    data.push({ range:`${S}!G${row}`, values:[[`=IFERROR(IF(A${row}="","",IFERROR('Party Setup'!B20,"")),"")` ]] });
    data.push({ range:`${S}!H${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,9,FALSE)),"")`]] });
    data.push({ range:`${S}!I${row}`, values:[[`=IFERROR(IF(A${row}="","",IF(AND(E${row}=TRUE,F${row}<>"",OR(H${row}="No Response",H${row}="Maybe"),H${row}<>"Declined"),"Yes","No")),"")`]] });
    data.push({ range:`${S}!L${row}`, values:[[`=IFERROR(IF(F${row}="","",TODAY()-F${row}),"")`]] });
    data.push({ range:`${S}!M${row}`, values:[[`=IFERROR(IF(A${row}="","",VLOOKUP(A${row},${G}!$A$6:$R$305,6,FALSE)&" | "&VLOOKUP(A${row},${G}!$A$6:$R$305,7,FALSE)),"")`]] });
  }

  // Dropdowns
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(IT,5,305,3,4),
    rule: { condition: { type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$E$2:$E$8` }] }, showCustomUi:true, strict:false }
  }});
  // Reminder Sent checkbox col K (index 10)
  dvReqs.push({ setDataValidation: {
    range: gridRange(IT,5,305,10,11),
    rule: { condition: { type:'BOOLEAN' }, showCustomUi:true }
  }});
  // Invitation Ready checkbox col E (index 4)
  dvReqs.push({ setDataValidation: {
    range: gridRange(IT,5,305,4,5),
    rule: { condition: { type:'BOOLEAN' }, showCustomUi:true }
  }});

  // Alternating rows
  for (let ri = 5; ri < 305; ri += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(IT,ri,ri+1,0,14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
  }

  // Formula cols B,C,G,H,I,L,M
  [1,2,6,7,8,11,12].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(IT,5,305,ci,ci+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat:{fontSize:9,fontFamily:'Arial'} }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
  });

  // CF: Follow-Up Needed col I (index 8)
  reqs.push({ addConditionalFormatRule: { index:0, rule: {
    ranges:[gridRange(IT,5,305,8,9)],
    booleanRule: {
      condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Yes'}] },
      format:{ backgroundColor:hex(C.attention), textFormat:{foregroundColor:hex(C.white),bold:true} }
    }
  }}});
  reqs.push({ addConditionalFormatRule: { index:1, rule: {
    ranges:[gridRange(IT,5,305,8,9)],
    booleanRule: {
      condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'No'}] },
      format:{ backgroundColor:hex(C.success), textFormat:{foregroundColor:hex(C.mainText)} }
    }
  }}});

  // CF: Response Status col H (index 7)
  const rsvpCF = [
    { val:'Confirmed', bg:C.success },
    { val:'Declined',  bg:C.border },
    { val:'Maybe',     bg:C.warning },
    { val:'No Response',bg:C.mutedMauve },
  ];
  rsvpCF.forEach((r, idx) => {
    reqs.push({ addConditionalFormatRule: { index:2+idx, rule: {
      ranges:[gridRange(IT,5,305,7,8)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:r.val}] },
        format:{ backgroundColor:hex(r.bg) }
      }
    }}});
  });

  // Column widths
  const colW = [80,130,140,120,75,90,90,100,90,90,75,80,190,160];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:IT, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:IT, tabColorStyle:{ rgbColor:hex(C.primary) } },
    fields:'tabColorStyle'
  }});

  await batchUpdate(id, reqs);
  await batchUpdate(id, dvReqs);
  await valuesBatchUpdate(id, data);
  console.log('Invitation Tracker complete');
})();
