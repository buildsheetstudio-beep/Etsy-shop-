'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S = "'Dashboard'";

const G  = "'Guest List & RSVP'";
const IT = "'Invitation Tracker'";
const BE = "'Budget & Expenses'";
const EM = "'Entertainment & Music'";
const TY = "'Thank-You Cards'";
const PP = "'Prep & Party Day'";
const PS = "'Party Setup'";

(async () => {
  const reqs = [];
  const data = [];

  // Background + freeze
  reqs.push({ repeatCell: {
    range: gridRange(DB,0,100,0,12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:DB, gridProperties:{ frozenRowCount:5 } },
    fields:'gridProperties.frozenRowCount'
  }});

  // ── Title row 0 ────────────────────────────────────────────────────────────
  data.push({ range:`${S}!A1`, values:[['RETIREMENT PARTY DASHBOARD']] });
  reqs.push({ mergeCells: { range: gridRange(DB,0,1,0,12), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DB,0,1,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:20, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:DB, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:52 }, fields:'pixelSize'
  }});

  // ── Subtitle row 1 ─────────────────────────────────────────────────────────
  data.push({ range:`${S}!A2`, values:[['Guests • Budget • Invitations • Entertainment • Thank-You Cards • Preparation']] });
  reqs.push({ mergeCells: { range: gridRange(DB,1,2,0,12), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DB,1,2,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:DB, dimension:'ROWS', startIndex:1, endIndex:2 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // Info rows 2-3
  data.push({ range:`${S}!A3`, values:[['Celebrating Patricia Collins — 32 Years of Service — Northfield Community Services']] });
  data.push({ range:`${S}!A4`, values:[['Party Date: June 27, 2026  |  Venue: Lakeside Event Hall  |  Host: Michael Torres  |  All values update automatically from your planning tabs']] });
  [2,3].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(DB,ri,ri+1,0,12), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DB,ri,ri+1,0,12),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:DB, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:22 }, fields:'pixelSize'
    }});
  });

  // ── 8 Primary KPI cards (rows 4-5 indexes) ───────────────────────────────
  // 4 cards in row 4, 4 cards in row 5. Each card: label(1col=3) + value(1col=3) = 6 col. 2×4=8... hmm 12 cols.
  // Let me do: each card uses 3 cols (label+value merged row pair), 4 cards = 12 cols
  // Row 4: card label row; row 5: value row? No — let me do label above, value below in same col group.
  // Simpler: 4 cards in one row each card spanning 3 cols. Label row = row index 4, value row = row index 5.

  const primaryKPIs = [
    { label:'Days Until Party',    formula:`=IFERROR(MAX(0,DATEVALUE("Jun 27, 2026")-TODAY()),"—")` },
    { label:'Confirmed Guests',    formula:`=IFERROR(COUNTIF(${G}!$I$6:$I$305,"Confirmed"),"—")` },
    { label:'Pending RSVPs',       formula:`=IFERROR(COUNTIF(${G}!$I$6:$I$305,"No Response")+COUNTIF(${G}!$I$6:$I$305,"Maybe"),"—")` },
    { label:'Total Budget',        formula:`=IFERROR(${PS}!B18,"—")` },
    { label:'Amount Spent',        formula:`=IFERROR(SUM(${BE}!$I$31:$I$329),0)` },
    { label:'Remaining Budget',    formula:`=IFERROR(${PS}!B18-SUM(${BE}!$I$31:$I$329),"—")` },
    { label:'Cost per Guest',      formula:`=IFERROR(IF(COUNTIF(${G}!$I$6:$I$305,"Confirmed")=0,"—",SUM(${BE}!$I$31:$I$329)/COUNTIF(${G}!$I$6:$I$305,"Confirmed")),"—")` },
    { label:'Planning Completion', formula:`=IFERROR(
0.20*IFERROR(COUNTIF(${G}!$H$6:$H$305,"Invitation Sent")/MAX(1,COUNTA(${G}!$C$6:$C$305)),0)+
0.20*IFERROR((COUNTIF(${G}!$I$6:$I$305,"Confirmed")+COUNTIF(${G}!$I$6:$I$305,"Declined"))/MAX(1,COUNTA(${G}!$C$6:$C$305)),0)+
0.15*IFERROR(IF(SUM(${BE}!$B$12:$B$27)>0,1,0),0)+
0.20*IFERROR(COUNTIF(${EM}!$N$10:$N$40,TRUE)/MAX(1,COUNTA(${EM}!$C$10:$C$40)),0)+
0.15*IFERROR(COUNTIF(${PP}!$J$8:$J$207,TRUE)/MAX(1,COUNTA(${PP}!$C$8:$C$207)),0)+
0.10*IFERROR((COUNTIF(${TY}!$I$7:$I$205,"Sent")+COUNTIF(${TY}!$I$7:$I$205,"Delivered")+COUNTIF(${TY}!$I$7:$I$205,"Not Needed"))/MAX(1,COUNTA(${TY}!$C$7:$C$205)),0)
,0)` },
  ];

  // Layout: 4 cards first row (indices 4+5 label+value), 4 cards second row (indices 7+8)
  // Each card: 3 cols wide. 4 cards × 3 = 12 cols
  primaryKPIs.forEach((kpi, i) => {
    const group = i < 4 ? 0 : 1;
    const pos = i % 4;
    const labelRI = 4 + group * 3; // row 4 for group 0, row 7 for group 1
    const valueRI = labelRI + 1;
    const c1 = pos * 3;
    const c2 = c1 + 3;

    const labelCell = String.fromCharCode(65+c1) + (labelRI+1);
    const valueCell = String.fromCharCode(65+c1) + (valueRI+1);
    data.push({ range:`${S}!${labelCell}`, values:[[kpi.label]] });
    data.push({ range:`${S}!${valueCell}`, values:[[kpi.formula]] });

    reqs.push({ mergeCells: { range: gridRange(DB,labelRI,labelRI+1,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DB,labelRI,labelRI+1,c1,c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ mergeCells: { range: gridRange(DB,valueRI,valueRI+1,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DB,valueRI,valueRI+1,c1,c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
  });

  // Row heights for KPI rows
  [4,5,7,8].forEach(ri => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:DB, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:ri%3===1?42:28 }, fields:'pixelSize'
    }});
  });

  // Separator rows between KPI groups (index 6)
  reqs.push({ repeatCell: {
    range: gridRange(DB,6,7,0,12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:DB, dimension:'ROWS', startIndex:6, endIndex:7 },
    properties: { pixelSize:8 }, fields:'pixelSize'
  }});

  // ── 8 Secondary KPI cards (rows 10-12) ────────────────────────────────────
  const secondaryKPIs = [
    { label:'Capacity Remaining',     formula:`=IFERROR(${PS}!B19-COUNTIF(${G}!$I$6:$I$305,"Confirmed"),"—")` },
    { label:'Invitations Not Sent',   formula:`=IFERROR(COUNTIF(${G}!$H$6:$H$305,"Invitation Ready")+COUNTIF(${G}!$H$6:$H$305,"Not Invited"),"—")` },
    { label:'Follow-Ups Needed',      formula:`=IFERROR(COUNTIF(${IT}!$I$6:$I$305,"Yes"),"—")` },
    { label:'Entertainment Pending',  formula:`=IFERROR(COUNTA(${EM}!$C$10:$C$40)-COUNTIF(${EM}!$N$10:$N$40,TRUE),"—")` },
    { label:'Speakers Confirmed',     formula:`=IFERROR(COUNTIF(${EM}!$I$44:$I$73,TRUE),"—")` },
    { label:'Music Items Ready',      formula:`=IFERROR(COUNTIF(${EM}!$H$59:$H$100,TRUE),"—")` },
    { label:'Thank-You Cards Left',   formula:`=IFERROR(COUNTIF(${TY}!$I$7:$I$205,"To Write")+COUNTIF(${TY}!$I$7:$I$205,"Drafted")+COUNTIF(${TY}!$I$7:$I$205,"Ready to Send"),"—")` },
    { label:'Prep Tasks Remaining',   formula:`=IFERROR(COUNTA(${PP}!$C$8:$C$207)-COUNTIF(${PP}!$J$8:$J$207,TRUE),"—")` },
  ];

  // Rows 10-11 (label 10, value 11) for group 1; rows 13-14 for group 2
  secondaryKPIs.forEach((kpi, i) => {
    const group = i < 4 ? 0 : 1;
    const pos = i % 4;
    const labelRI = 10 + group * 3;
    const valueRI = labelRI + 1;
    const c1 = pos * 3;
    const c2 = c1 + 3;

    const labelCell = String.fromCharCode(65+c1) + (labelRI+1);
    const valueCell = String.fromCharCode(65+c1) + (valueRI+1);
    data.push({ range:`${S}!${labelCell}`, values:[[kpi.label]] });
    data.push({ range:`${S}!${valueCell}`, values:[[kpi.formula]] });

    reqs.push({ mergeCells: { range: gridRange(DB,labelRI,labelRI+1,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DB,labelRI,labelRI+1,c1,c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ mergeCells: { range: gridRange(DB,valueRI,valueRI+1,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DB,valueRI,valueRI+1,c1,c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:14, foregroundColor:hex(C.primary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
  });

  [10,11,13,14].forEach(ri => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:DB, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:ri%3===1?38:26 }, fields:'pixelSize'
    }});
  });

  // Separator rows (9, 12, 15)
  [9,12,15].forEach(ri => {
    reqs.push({ repeatCell: {
      range: gridRange(DB,ri,ri+1,0,12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:DB, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:8 }, fields:'pixelSize'
    }});
  });

  // ── Section: Upcoming Tasks (rows 16-27) ─────────────────────────────────
  // Section header row 16
  data.push({ range:`${S}!A17`, values:[['NEXT 10 UPCOMING TASKS (by due date)']] });
  reqs.push({ mergeCells: { range: gridRange(DB,16,17,0,9), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DB,16,17,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:DB, dimension:'ROWS', startIndex:16, endIndex:17 },
    properties: { pixelSize:26 }, fields:'pixelSize'
  }});

  // Column headers row 17
  data.push({ range:`${S}!A18`, values:[['#','Phase','Task','Assigned To','Due Date','Priority','Status','Days Left','Overdue?']] });
  reqs.push({ repeatCell: {
    range: gridRange(DB,17,18,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // 10 upcoming task rows via SMALL/IF (rows 18-27, index 18-27)
  for (let n = 1; n <= 10; n++) {
    const row = 18 + n;
    data.push({ range:`${S}!A${row}`, values:[[n]] });
    data.push({ range:`${S}!B${row}`, values:[[`=IFERROR(INDEX(${PP}!$B$8:$B$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });
    data.push({ range:`${S}!C${row}`, values:[[`=IFERROR(INDEX(${PP}!$C$8:$C$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });
    data.push({ range:`${S}!D${row}`, values:[[`=IFERROR(INDEX(${PP}!$E$8:$E$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });
    data.push({ range:`${S}!E${row}`, values:[[`=IFERROR(INDEX(${PP}!$G$8:$G$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });
    data.push({ range:`${S}!F${row}`, values:[[`=IFERROR(INDEX(${PP}!$H$8:$H$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });
    data.push({ range:`${S}!G${row}`, values:[[`=IFERROR(INDEX(${PP}!$I$8:$I$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });
    data.push({ range:`${S}!H${row}`, values:[[`=IFERROR(INDEX(${PP}!$L$8:$L$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });
    data.push({ range:`${S}!I${row}`, values:[[`=IFERROR(INDEX(${PP}!$M$8:$M$207,MATCH(SMALL(IF((${PP}!$J$8:$J$207=FALSE)*(${PP}!$G$8:$G$207<>""),${PP}!$G$8:$G$207),${n}),${PP}!$G$8:$G$207,0)),"")` ]] });

    // Overdue row CF: red tint
    reqs.push({ addConditionalFormatRule: { index:n-1, rule: {
      ranges:[gridRange(DB,18+n,19+n,0,9)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Yes'}] },
        format:{ backgroundColor:hex('#F5E8E8') }
      }
    }}});
    // Row alt tint
    if (n % 2 === 0) {
      reqs.push({ repeatCell: {
        range: gridRange(DB,18+n,19+n,0,9),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
        fields:'userEnteredFormat(backgroundColor)'
      }});
    }
  }

  // ── Entertainment Snapshot (rows 30-37) ─────────────────────────────────
  data.push({ range:`${S}!A31`, values:[['ENTERTAINMENT SNAPSHOT']] });
  reqs.push({ mergeCells: { range: gridRange(DB,30,31,0,9), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DB,30,31,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:DB, dimension:'ROWS', startIndex:30, endIndex:31 },
    properties: { pixelSize:26 }, fields:'pixelSize'
  }});

  const entSnapshot = [
    { label:'Entertainment Items Confirmed', formula:`=IFERROR(COUNTIF(${EM}!$N$10:$N$40,TRUE),"—")` },
    { label:'Total Planned Speaking Time',   formula:`=IFERROR(${EM}!F55&" min","—")` },
    { label:'Speakers Confirmed',            formula:`=IFERROR(COUNTIF(${EM}!$I$44:$I$73,TRUE),"—")` },
    { label:'Music Items Ready',             formula:`=IFERROR(COUNTIF(${EM}!$H$59:$H$100,TRUE),"—")` },
    { label:'AV Equipment Checked',          formula:`=IFERROR(COUNTIF(${EM}!$C$86:$C$97,TRUE),"—")` },
    { label:'Entertainment Budget',          formula:`=IFERROR(SUMIF(${BE}!$C$12:$C$27,"Entertainment",${BE}!$B$12:$B$27)+SUMIF(${BE}!$C$12:$C$27,"Music",${BE}!$B$12:$B$27)+SUMIF(${BE}!$C$12:$C$27,"Audio / Visual",${BE}!$B$12:$B$27),0)` },
  ];

  entSnapshot.forEach(({ label, formula }, i) => {
    const row = 32 + i;
    data.push({ range:`${S}!A${row}`, values:[[label]] });
    data.push({ range:`${S}!D${row}`, values:[[formula]] });
    reqs.push({ repeatCell: {
      range: gridRange(DB,row-1,row,0,3),
      cell: { userEnteredFormat: {
        backgroundColor: i%2===0 ? hex(C.panel) : hex(C.altRow),
        textFormat: { bold:true, fontSize:9, fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DB,row-1,row,3,6),
      cell: { userEnteredFormat: {
        backgroundColor: i%2===0 ? hex(C.panel) : hex(C.altRow),
        textFormat: { bold:true, fontSize:13, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ mergeCells: { range: gridRange(DB,row-1,row,0,3), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DB,row-1,row,3,6), mergeType:'MERGE_ALL' } });
  });

  // ── Planning Progress Weights (rows 40-47) ────────────────────────────────
  data.push({ range:`${S}!A41`, values:[['PLANNING PROGRESS WEIGHTS']] });
  reqs.push({ mergeCells: { range: gridRange(DB,40,41,0,9), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DB,40,41,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:DB, dimension:'ROWS', startIndex:40, endIndex:41 },
    properties: { pixelSize:26 }, fields:'pixelSize'
  }});

  const weights = [
    ['Invitations Sent', '20%', `=IFERROR(COUNTIF(${G}!$H$6:$H$305,"Invitation Sent")/MAX(1,COUNTA(${G}!$C$6:$C$305)),0)`],
    ['RSVPs Resolved', '20%', `=IFERROR((COUNTIF(${G}!$I$6:$I$305,"Confirmed")+COUNTIF(${G}!$I$6:$I$305,"Declined"))/MAX(1,COUNTA(${G}!$C$6:$C$305)),0)`],
    ['Budget Setup', '15%', `=IFERROR(IF(SUM(${BE}!$B$12:$B$27)>0,1,0),0)`],
    ['Entertainment Ready', '20%', `=IFERROR(COUNTIF(${EM}!$N$10:$N$40,TRUE)/MAX(1,COUNTA(${EM}!$C$10:$C$40)),0)`],
    ['Prep Tasks Done', '15%', `=IFERROR(COUNTIF(${PP}!$J$8:$J$207,TRUE)/MAX(1,COUNTA(${PP}!$C$8:$C$207)),0)`],
    ['Thank-You Prep', '10%', `=IFERROR((COUNTIF(${TY}!$I$7:$I$205,"Sent")+COUNTIF(${TY}!$I$7:$I$205,"Delivered")+COUNTIF(${TY}!$I$7:$I$205,"Not Needed"))/MAX(1,COUNTA(${TY}!$C$7:$C$205)),0)`],
  ];

  data.push({ range:`${S}!A42`, values:[['Component','Weight','Progress']] });
  reqs.push({ repeatCell: {
    range: gridRange(DB,41,42,0,3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  weights.forEach(([label, weight, formula], i) => {
    const row = 43 + i;
    data.push({ range:`${S}!A${row}`, values:[[label]] });
    data.push({ range:`${S}!B${row}`, values:[[weight]] });
    data.push({ range:`${S}!C${row}`, values:[[formula]] });
    reqs.push({ repeatCell: {
      range: gridRange(DB,row-1,row,0,3),
      cell: { userEnteredFormat: {
        backgroundColor: i%2===0 ? hex(C.panel) : hex(C.altRow),
        textFormat: { fontSize:9, fontFamily:'Arial' }
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
  });

  // Disclaimer row 50
  data.push({ range:`${S}!A51`, values:[['All values refresh automatically. Update Guest List, Budget, Entertainment, Prep, and Thank-You tabs to see live results here.']] });
  reqs.push({ mergeCells: { range: gridRange(DB,50,51,0,12), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DB,50,51,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { italic:true, fontSize:8, foregroundColor:hex(C.secText), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // ── Pivot tables for charts (cols J-L) ───────────────────────────────────
  // RSVP pivot J1:K9
  data.push({ range:`${S}!J1`, values:[['RSVP Status','Count']] });
  const rsvpStatuses = ['Confirmed','Declined','Maybe','No Response','Invitation Sent','Follow-Up Needed','Invitation Ready','Not Invited'];
  rsvpStatuses.forEach((s, i) => {
    data.push({ range:`${S}!J${2+i}`, values:[[s]] });
    data.push({ range:`${S}!K${2+i}`, values:[[`=IFERROR(COUNTIF(${G}!$I$6:$I$305,"${s}"),0)`]] });
  });

  // Spending by category pivot J11:K26
  data.push({ range:`${S}!J11`, values:[['Category','Actual Spent']] });
  const cats = ['Venue','Invitations','Decorations','Food','Drinks','Cake & Desserts','Entertainment','Music','Audio / Visual','Gifts','Awards','Photography','Supplies','Transportation','Setup & Cleanup','Miscellaneous'];
  cats.forEach((cat, i) => {
    data.push({ range:`${S}!J${12+i}`, values:[[cat]] });
    data.push({ range:`${S}!K${12+i}`, values:[[`=IFERROR(SUMIF(${BE}!$C$31:$C$329,"${cat}",${BE}!$I$31:$I$329),0)`]] });
  });

  // Budget vs Actual pivot J29:L44
  data.push({ range:`${S}!J29`, values:[['Category','Planned','Actual']] });
  cats.forEach((cat, i) => {
    data.push({ range:`${S}!J${30+i}`, values:[[cat]] });
    data.push({ range:`${S}!K${30+i}`, values:[[`=IFERROR(SUMIF(${BE}!$A$12:$A$27,"${cat}",${BE}!$B$12:$B$27),0)`]] });
    data.push({ range:`${S}!L${30+i}`, values:[[`=IFERROR(SUMIF(${BE}!$C$31:$C$329,"${cat}",${BE}!$I$31:$I$329),0)`]] });
  });

  // Task status pivot J47:K51
  data.push({ range:`${S}!J47`, values:[['Task Status','Count']] });
  const taskStatuses = ['Complete','In Progress','Not Started','Waiting','Cancelled'];
  taskStatuses.forEach((s, i) => {
    data.push({ range:`${S}!J${48+i}`, values:[[s]] });
    data.push({ range:`${S}!K${48+i}`, values:[[`=IFERROR(COUNTIF(${PP}!$I$8:$I$207,"${s}"),0)`]] });
  });

  // Entertainment type pivot J54:K65
  data.push({ range:`${S}!J54`, values:[['Entertainment Type','Count']] });
  const entTypes = ['Live Music','DJ','Playlist','Slideshow','Video Tribute','Speaker','Toast','Games','Photo Booth','Performance','Open Mic','Other'];
  entTypes.forEach((t, i) => {
    data.push({ range:`${S}!J${55+i}`, values:[[t]] });
    data.push({ range:`${S}!K${55+i}`, values:[[`=IFERROR(COUNTIF(${EM}!$B$10:$B$40,"${t}"),0)`]] });
  });

  // Thank-You card status pivot J68:K73
  data.push({ range:`${S}!J68`, values:[['Card Status','Count']] });
  const tyStatuses = ['Not Needed','To Write','Drafted','Ready to Send','Sent','Delivered'];
  tyStatuses.forEach((s, i) => {
    data.push({ range:`${S}!J${69+i}`, values:[[s]] });
    data.push({ range:`${S}!K${69+i}`, values:[[`=IFERROR(COUNTIF(${TY}!$I$7:$I$205,"${s}"),0)`]] });
  });

  // ── Column widths ──────────────────────────────────────────────────────────
  const colW = [160,80,200,100,90,80,90,70,70,100,110,110];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:DB, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:DB, tabColorStyle:{ rgbColor:hex(C.secondary) } },
    fields:'tabColorStyle'
  }});

  await batchUpdate(id, reqs);
  await valuesBatchUpdate(id, data);

  // ── 6 Charts ───────────────────────────────────────────────────────────────
  const chartReqs = [];

  // 1. RSVP donut J1:K9
  chartReqs.push({ addChart: { chart: {
    spec: {
      title:'RSVP Status',
      pieChart: {
        legendPosition:'RIGHT_LEGEND', pieHole:0.4,
        domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:0, endRowIndex:9, startColumnIndex:9, endColumnIndex:10 }]}},
        series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:0, endRowIndex:9, startColumnIndex:10, endColumnIndex:11 }]}}
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:5, columnIndex:9 }, widthPixels:380, heightPixels:220 } }
  }}});

  // 2. Spending by category donut J11:K26
  chartReqs.push({ addChart: { chart: {
    spec: {
      title:'Actual Spending by Category',
      pieChart: {
        legendPosition:'RIGHT_LEGEND', pieHole:0.4,
        domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:10, endRowIndex:27, startColumnIndex:9, endColumnIndex:10 }]}},
        series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:10, endRowIndex:27, startColumnIndex:10, endColumnIndex:11 }]}}
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:15, columnIndex:9 }, widthPixels:380, heightPixels:230 } }
  }}});

  // 3. Budget vs Actual column J29:L44
  chartReqs.push({ addChart: { chart: {
    spec: {
      title:'Planned vs. Actual Budget',
      basicChart: {
        chartType:'COLUMN', legendPosition:'RIGHT_LEGEND',
        domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:28, endRowIndex:45, startColumnIndex:9, endColumnIndex:10 }]}}}],
        series:[
          { series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:28, endRowIndex:45, startColumnIndex:10, endColumnIndex:11 }]}}, targetAxis:'LEFT_AXIS' },
          { series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:28, endRowIndex:45, startColumnIndex:11, endColumnIndex:12 }]}}, targetAxis:'LEFT_AXIS' },
        ],
        headerCount:1
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:29, columnIndex:9 }, widthPixels:440, heightPixels:250 } }
  }}});

  // 4. Prep task status bar J47:K51
  chartReqs.push({ addChart: { chart: {
    spec: {
      title:'Prep Tasks by Status',
      basicChart: {
        chartType:'BAR', legendPosition:'BOTTOM_LEGEND',
        domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:47, endRowIndex:52, startColumnIndex:9, endColumnIndex:10 }]}}}],
        series:[{ series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:47, endRowIndex:52, startColumnIndex:10, endColumnIndex:11 }]}}, targetAxis:'BOTTOM_AXIS' }],
        headerCount:1
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:16, columnIndex:9 }, widthPixels:380, heightPixels:210, offsetYPixels:240 } }
  }}});

  // 5. Entertainment type donut J54:K65
  chartReqs.push({ addChart: { chart: {
    spec: {
      title:'Entertainment by Type',
      pieChart: {
        legendPosition:'RIGHT_LEGEND', pieHole:0.4,
        domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:53, endRowIndex:66, startColumnIndex:9, endColumnIndex:10 }]}},
        series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:53, endRowIndex:66, startColumnIndex:10, endColumnIndex:11 }]}}
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:39, columnIndex:9 }, widthPixels:380, heightPixels:200 } }
  }}});

  // 6. Thank-You card status bar J68:K73
  chartReqs.push({ addChart: { chart: {
    spec: {
      title:'Thank-You Card Status',
      basicChart: {
        chartType:'BAR', legendPosition:'BOTTOM_LEGEND',
        domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:67, endRowIndex:74, startColumnIndex:9, endColumnIndex:10 }]}}}],
        series:[{ series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:67, endRowIndex:74, startColumnIndex:10, endColumnIndex:11 }]}}, targetAxis:'BOTTOM_AXIS' }],
        headerCount:1
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:46, columnIndex:9 }, widthPixels:380, heightPixels:200 } }
  }}});

  await batchUpdate(id, chartReqs);
  console.log('Dashboard complete');
})();
