'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TY = sheetMap['Thank-You Cards'];
const S = "'Thank-You Cards'";
const G = "'Guest List & RSVP'";
const R = "'Reference Data'";

const HEADERS = ['Thank-You ID','Guest ID','Guest Name','Household / Org','Gift / Contribution?','Gift Description','Card Needed?','Message Notes','Status','Written Date','Sent Date','Delivery Method','Address / Email','Assigned To','Days Since Party','Notes'];

// 24 sample rows [guestId, cardNeeded, msgNotes, status, writtenDate, sentDate, delivery, assignedTo, notes]
// Gift data pulled via VLOOKUP from Guest List
const TY_ROWS = [
  ['GST-001','No','Retiree — no card needed','Not Needed','','','','Patricia Collins','Guest of honor'],
  ['GST-002','No','Spouse — courtesy only','Not Needed','','','','',''],
  ['GST-003','No','No gift noted','Not Needed','','','','',''],
  ['GST-004','No','No gift noted','Not Needed','','','','',''],
  ['GST-006','Yes','Office gift card pool — mention team effort specifically','To Write','','','Card','Patricia Collins',''],
  ['GST-011','Yes','Personalized crystal plaque — mention it proudly displayed','To Write','','','Card','James Collins',''],
  ['GST-020','Yes','City commendation certificate — honor and gratitude','To Write','','','Card + Email','Patricia Collins','Thank mayor for attending'],
  ['GST-023','Yes','Custom photo book — mention the memories inside','Drafted','Jun 28, 2026','','Card','Patricia Collins',''],
  ['GST-024','Yes','Wine set — mention enjoying it together','Drafted','Jun 28, 2026','','Card','Patricia Collins',''],
  ['GST-026','Yes','Sam & Joyce — photo album — mention evening together','To Write','','','Card','Patricia Collins',''],
  ['GST-029','Yes','Book collection — thank deeply, longtime colleague','To Write','','','Card','Patricia Collins',''],
  ['GST-032','Yes','Garden set — Barbara was a treasured mentor','To Write','','','Card','Patricia Collins',''],
  ['GST-035','Yes','Keepsake box — Grace and Mark — family warmth','Drafted','Jun 28, 2026','','Card','Patricia Collins',''],
  ['GST-005','Yes','Plaque and gift card pool contribution','To Write','','','Card','Michael Torres','Coordinate with HR for thank-you note'],
  ['GST-010','Yes','Gift basket from family — acknowledge each member','Ready to Send','Jun 27, 2026','','Card Hand-Delivered','',''],
  ['GST-014','Yes','Memory box / message contribution','To Write','','','Email','Patricia Collins','Group contributors'],
  ['GST-016','Yes','Gift contribution through shared pool','To Write','','','Card','James Collins',''],
  ['GST-017','Yes','Contribution via group card','To Write','','','Email','Patricia Collins',''],
  ['GST-019','Yes','Retirement wishes written in guest book','To Write','','','Card','Patricia Collins',''],
  ['GST-025','Yes','Venue coordination gift card','Sent','Jun 29, 2026','Jun 30, 2026','Email','Michael Torres',''],
  ['GST-030','Yes','Cash contribution to group gift','To Write','','','Card','Patricia Collins',''],
  ['GST-031','Yes','Heartfelt card and generous cash gift','To Write','','','Card','Patricia Collins',''],
  ['GST-033','Yes','Orchid centerpiece given as gift','Drafted','Jun 28, 2026','','Card','Patricia Collins',''],
  ['GST-036','Yes','Donation in Patricia\'s honor to her charity of choice','To Write','','','Card','Patricia Collins',''],
];

(async () => {
  const reqs = [];
  const data = [];

  // Background + freeze
  reqs.push({ repeatCell: {
    range: gridRange(TY,0,300,0,16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:TY, gridProperties:{ frozenRowCount:5 } },
    fields:'gridProperties.frozenRowCount'
  }});

  // Title
  data.push({ range:`${S}!A1`, values:[['THANK-YOU CARDS']] });
  reqs.push({ mergeCells: { range: gridRange(TY,0,1,0,16), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(TY,0,1,0,16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:TY, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:44 }, fields:'pixelSize'
  }});

  // Subtitle
  data.push({ range:`${S}!A2`, values:[['Track thank-you card status for every gift and contribution received']] });
  reqs.push({ mergeCells: { range: gridRange(TY,1,2,0,16), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(TY,1,2,0,16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Info rows 2-3
  data.push({ range:`${S}!A3`, values:[['Guest Name, Household, Gift/Contribution, and Gift Description auto-fill from Guest List using Guest ID']] });
  data.push({ range:`${S}!A4`, values:[['Status dropdown: Not Needed | To Write | Drafted | Ready to Send | Sent | Delivered']] });
  [2,3].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(TY,ri,ri+1,0,16), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(TY,ri,ri+1,0,16),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:TY, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:22 }, fields:'pixelSize'
    }});
  });

  // ── 5 summary cards (row index 4, row 5) ──────────────────────────────────
  // Cards: Cards Needed, To Write, Ready to Send, Sent, Completion %
  const cards = [
    { label:'Cards Needed',     formula:`=IFERROR(COUNTIF(I6:I205,"<>Not Needed"),"—")`, c:[0,3] },
    { label:'To Write',         formula:`=IFERROR(COUNTIF(I6:I205,"To Write"),"—")`, c:[3,6] },
    { label:'Ready to Send',    formula:`=IFERROR(COUNTIF(I6:I205,"Ready to Send"),"—")`, c:[6,9] },
    { label:'Sent / Delivered', formula:`=IFERROR(COUNTIF(I6:I205,"Sent")+COUNTIF(I6:I205,"Delivered"),"—")`, c:[9,12] },
    { label:'Completion %',     formula:`=IFERROR(IF(COUNTIF(I6:I205,"<>Not Needed")=0,"—",(COUNTIF(I6:I205,"Sent")+COUNTIF(I6:I205,"Delivered"))/COUNTIF(I6:I205,"<>Not Needed")),"—")`, c:[12,16] },
  ];

  cards.forEach(card => {
    const [c1, c2] = card.c;
    const mid = Math.floor((c1 + c2) / 2);
    const labelCell = String.fromCharCode(65 + c1) + '5';
    const valueCell = String.fromCharCode(65 + mid) + '5';
    // Split label and value within the range
    data.push({ range:`${S}!${labelCell}`, values:[[card.label]] });
    data.push({ range:`${S}!${valueCell}`, values:[[card.formula]] });

    const splitPoint = Math.ceil((c1 + c2) / 2);
    reqs.push({ mergeCells: { range: gridRange(TY,4,5,c1,splitPoint), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(TY,4,5,c1,splitPoint),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ mergeCells: { range: gridRange(TY,4,5,splitPoint,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(TY,4,5,splitPoint,c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:14, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:TY, dimension:'ROWS', startIndex:4, endIndex:5 },
    properties: { pixelSize:38 }, fields:'pixelSize'
  }});

  // ── Header row (index 5) ───────────────────────────────────────────────────
  data.push({ range:`${S}!A6`, values:[HEADERS] });
  reqs.push({ repeatCell: {
    range: gridRange(TY,5,6,0,16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:TY, dimension:'ROWS', startIndex:5, endIndex:6 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // ── Sample data rows 6-29 (index 6-29) ───────────────────────────────────
  TY_ROWS.forEach(([guestId, cardNeeded, msgNotes, status, writtenDate, sentDate, delivery, assignedTo, notes], i) => {
    const row = 7 + i;
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","TY-"&TEXT(ROW()-6,"000"))`]] });
    data.push({ range:`${S}!B${row}`, values:[[guestId]] });
    // C: Guest Name via VLOOKUP
    data.push({ range:`${S}!C${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,3,FALSE)),"")`]] });
    // D: Household
    data.push({ range:`${S}!D${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,2,FALSE)),"")`]] });
    // E: Gift / Contribution
    data.push({ range:`${S}!E${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,13,FALSE)),"")`]] });
    // F: Gift Description
    data.push({ range:`${S}!F${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,14,FALSE)),"")`]] });
    // G: Card Needed
    data.push({ range:`${S}!G${row}`, values:[[cardNeeded]] });
    // H: Message Notes
    data.push({ range:`${S}!H${row}`, values:[[msgNotes]] });
    // I: Status
    data.push({ range:`${S}!I${row}`, values:[[status]] });
    // J: Written Date
    if (writtenDate) data.push({ range:`${S}!J${row}`, values:[[writtenDate]] });
    // K: Sent Date
    if (sentDate) data.push({ range:`${S}!K${row}`, values:[[sentDate]] });
    // L: Delivery Method
    if (delivery) data.push({ range:`${S}!L${row}`, values:[[delivery]] });
    // M: Address / Email (pull from Guest List col 7 = email)
    data.push({ range:`${S}!M${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,7,FALSE)),"")`]] });
    // N: Assigned To
    data.push({ range:`${S}!N${row}`, values:[[assignedTo]] });
    // O: Days Since Party
    data.push({ range:`${S}!O${row}`, values:[[`=IFERROR(IF(B${row}="","",TODAY()-DATEVALUE("Jun 27, 2026")),"")`]] });
    // P: Notes
    data.push({ range:`${S}!P${row}`, values:[[notes]] });
  });

  // Formula rows for remaining 25-205 (index 30-204)
  for (let row = 31; row <= 205; row++) {
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","TY-"&TEXT(ROW()-6,"000"))`]] });
    data.push({ range:`${S}!C${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,3,FALSE)),"")`]] });
    data.push({ range:`${S}!D${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,2,FALSE)),"")`]] });
    data.push({ range:`${S}!E${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,13,FALSE)),"")`]] });
    data.push({ range:`${S}!F${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,14,FALSE)),"")`]] });
    data.push({ range:`${S}!M${row}`, values:[[`=IFERROR(IF(B${row}="","",VLOOKUP(B${row},${G}!$A$6:$R$305,7,FALSE)),"")`]] });
    data.push({ range:`${S}!O${row}`, values:[[`=IFERROR(IF(B${row}="","",TODAY()-DATEVALUE("Jun 27, 2026")),"")`]] });
  }

  // Alternating rows
  for (let ri = 6; ri < 205; ri += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(TY,ri,ri+1,0,16),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
  }

  // Formula col styling A,C,D,E,F,M,O
  [0,2,3,4,5,12,14].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(TY,6,205,ci,ci+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat:{fontSize:9,fontFamily:'Arial'} }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
  });

  // Status dropdown + CF
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(TY,6,205,8,9),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$O$2:$O$7` }] }, showCustomUi:true, strict:false }
  }});

  // Status CF col I (index 8)
  const statusCF = [
    { val:'To Write',       bg:C.warning,   text:C.mainText },
    { val:'Drafted',        bg:C.mutedBlue, text:C.mainText },
    { val:'Ready to Send',  bg:C.mutedBlue, text:C.mainText },
    { val:'Sent',           bg:C.success,   text:C.mainText },
    { val:'Delivered',      bg:C.success,   text:C.mainText },
    { val:'Not Needed',     bg:C.border,    text:C.secText },
  ];
  statusCF.forEach((r, idx) => {
    reqs.push({ addConditionalFormatRule: { index:idx, rule: {
      ranges:[gridRange(TY,6,205,8,9)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:r.val}] },
        format:{ backgroundColor:hex(r.bg), textFormat:{foregroundColor:hex(r.text)} }
      }
    }}});
  });

  // Full-row gray for Not Needed
  reqs.push({ addConditionalFormatRule: { index:6, rule: {
    ranges:[gridRange(TY,6,205,0,16)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=$I7="Not Needed"`}] },
      format:{ backgroundColor:hex('#F0EDEB'), textFormat:{foregroundColor:hex(C.secText)} }
    }
  }}});

  // Sent/Delivered full-row sage tint
  reqs.push({ addConditionalFormatRule: { index:7, rule: {
    ranges:[gridRange(TY,6,205,0,16)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=OR($I7="Sent",$I7="Delivered")`}] },
      format:{ backgroundColor:hex('#EEF4EF') }
    }
  }}});

  // Column widths
  const colW = [80,80,130,140,75,160,75,200,100,90,80,110,180,120,90,160];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:TY, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:TY, tabColorStyle:{ rgbColor:hex(C.primary) } },
    fields:'tabColorStyle'
  }});

  await batchUpdate(id, reqs);
  await batchUpdate(id, dvReqs);
  await valuesBatchUpdate(id, data);
  console.log('Thank-You Cards complete');
})();
