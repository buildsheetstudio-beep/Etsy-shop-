'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DASH = sheetMap['🏡 Dashboard'];

const GL   = "'👨‍👩‍👧‍👦 Guest List & RSVPs'";
const BDG  = "'💰 Budget & Expenses'";
const CKL  = "'✅ Planning Checklist'";
const BRANCHES = ['Henderson Side','Thompson Side','Wilson Side','Smith Side','Out-of-Town','Family Friends'];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🏡 Dashboard'";

  // Column widths
  [30,200,220,30,200,200,30,200,200].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DASH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // ── Row 0: Hero title banner ──
  reqs.push({ mergeCells: { range: gridRange(DASH,0,1,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,0,1,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 20 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 70 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['🏡 The Henderson Family Reunion 2025']] });

  // ── Rows 1-7: Input block (B-C) + KPI block (E-F) ──
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,8,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmCream),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});

  // Input labels (col B, bold)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,8,1,2),
    cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: hex(C.forestGreen), fontSize: 9 },
    }},
    fields: 'userEnteredFormat.textFormat',
  }});
  // Input values (col C, ivory bg)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,8,2,3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.ivory),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  // KPI labels (col E, bold)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,8,4,5),
    cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: hex(C.forestGreen), fontSize: 9 },
    }},
    fields: 'userEnteredFormat.textFormat',
  }});
  // KPI values (col F, amber bold big)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,8,5,6),
    cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: hex(C.warmAmber), fontSize: 14 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
  }});

  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 1, endIndex: 8 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Input data
  vals.push({ range: `${TAB}!B2`, values: [['REUNION NAME']] });
  vals.push({ range: `${TAB}!C2`, values: [['The Henderson Family Reunion 2025']] });
  vals.push({ range: `${TAB}!B3`, values: [['DATE(S)']] });
  vals.push({ range: `${TAB}!C3`, values: [['15–16 August 2025']] });
  vals.push({ range: `${TAB}!B4`, values: [['LOCATION']] });
  vals.push({ range: `${TAB}!C4`, values: [['Lake Eildon Cabin, Victoria']] });
  vals.push({ range: `${TAB}!B5`, values: [['ORGANISER']] });
  vals.push({ range: `${TAB}!C5`, values: [['James Henderson']] });
  vals.push({ range: `${TAB}!B6`, values: [['START DATE (for countdown)']] });
  vals.push({ range: `${TAB}!C6`, values: [['2025-08-15']] });
  // Date format for C6
  reqs.push({ repeatCell: {
    range: gridRange(DASH,5,6,2,3),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }, backgroundColor: hex(C.ivory) }},
    fields: 'userEnteredFormat(numberFormat,backgroundColor)',
  }});
  vals.push({ range: `${TAB}!B7`, values: [['DAYS UNTIL REUNION']] });
  vals.push({ range: `${TAB}!C7`, values: [['=IFERROR(IF(TODAY()<=C6,C6-TODAY()&" days to go 🎉","Reunion complete! 🏡"),"Enter start date in C6")']] });

  // KPI values (formulas)
  vals.push({ range: `${TAB}!E2`, values: [['GUESTS INVITED']] });
  vals.push({ range: `${TAB}!F2`, values: [[`=IFERROR(COUNTA(${GL}!A4:A200),"—")`]] });
  vals.push({ range: `${TAB}!E3`, values: [['CONFIRMED']] });
  vals.push({ range: `${TAB}!F3`, values: [[`=IFERROR(COUNTIF(${GL}!H4:H200,"Confirmed"),"—")`]] });
  vals.push({ range: `${TAB}!E4`, values: [['AWAITING RSVP']] });
  vals.push({ range: `${TAB}!F4`, values: [[`=IFERROR(COUNTIF(${GL}!H4:H200,"Invited")+COUNTIF(${GL}!H4:H200,"No Response"),"—")`]] });
  vals.push({ range: `${TAB}!E5`, values: [['DECLINED']] });
  vals.push({ range: `${TAB}!F5`, values: [[`=IFERROR(COUNTIF(${GL}!H4:H200,"Declined"),"—")`]] });
  vals.push({ range: `${TAB}!E6`, values: [['% CONFIRMED']] });
  vals.push({ range: `${TAB}!F6`, values: [[`=IFERROR(TEXT(COUNTIF(${GL}!H4:H200,"Confirmed")/COUNTA(${GL}!A4:A200),"0%"),"—")`]] });
  vals.push({ range: `${TAB}!E7`, values: [['CO-ORGANISER']] });
  vals.push({ range: `${TAB}!F7`, values: [['Sarah Henderson']] });

  // ── Row 8: Amber section divider ──
  reqs.push({ repeatCell: {
    range: gridRange(DASH,8,9,0,9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmAmber) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // ── Row 9: Budget section header ──
  reqs.push({ mergeCells: { range: gridRange(DASH,9,10,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,9,10,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 9, endIndex: 10 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A10`, values: [['  💰  BUDGET OVERVIEW']] });

  // Budget KPIs (rows 10-14 = GSheets 11-15)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,10,15,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmCream),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,10,15,1,2),
    cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.forestGreen), fontSize: 9 }}},
    fields: 'userEnteredFormat.textFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,10,15,2,3),
    cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.warmAmber), fontSize: 13 }, horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 10, endIndex: 15 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  vals.push({ range: `${TAB}!B11`, values: [['TOTAL BUDGET']] });
  vals.push({ range: `${TAB}!C11`, values: [[`=IFERROR(SUM(${BDG}!C4:C15),"—")`]] });
  vals.push({ range: `${TAB}!B12`, values: [['TOTAL SPENT']] });
  vals.push({ range: `${TAB}!C12`, values: [[`=IFERROR(SUM(${BDG}!F19:F200),"—")`]] });
  vals.push({ range: `${TAB}!B13`, values: [['REMAINING']] });
  vals.push({ range: `${TAB}!C13`, values: [['=IFERROR(C11-C12,"—")']] });
  vals.push({ range: `${TAB}!B14`, values: [['% USED']] });
  vals.push({ range: `${TAB}!C14`, values: [['=IFERROR(TEXT(C12/C11,"0%"),"—")']] });
  vals.push({ range: `${TAB}!B15`, values: [['COST PER CONFIRMED GUEST']] });
  vals.push({ range: `${TAB}!C15`, values: [[`=IFERROR("$"&TEXT(C12/COUNTIF(${GL}!H4:H200,"Confirmed"),"#,##0.00"),"—")`]] });

  // Currency format for C11-C13
  reqs.push({ repeatCell: {
    range: gridRange(DASH,10,13,2,3),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }}},
    fields: 'userEnteredFormat.numberFormat',
  }});

  // ── Row 15: Checklist section header ──
  reqs.push({ repeatCell: {
    range: gridRange(DASH,15,16,0,9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmAmber) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 15, endIndex: 16 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  reqs.push({ mergeCells: { range: gridRange(DASH,16,17,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,16,17,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 16, endIndex: 17 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A17`, values: [['  ✅  PLANNING CHECKLIST OVERVIEW']] });

  // Checklist stats (rows 17-19 = GSheets 18-20)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,17,22,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmCream),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,17,22,1,2),
    cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.forestGreen), fontSize: 9 }}},
    fields: 'userEnteredFormat.textFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,17,22,2,3),
    cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.warmAmber), fontSize: 13 }, horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 17, endIndex: 22 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  vals.push({ range: `${TAB}!B18`, values: [['TASKS DONE']] });
  vals.push({ range: `${TAB}!C18`, values: [[`=IFERROR(COUNTIF(${CKL}!A4:A200,TRUE)&" of "&COUNTA(${CKL}!C4:C200),"—")`]] });
  vals.push({ range: `${TAB}!B19`, values: [['% COMPLETE']] });
  vals.push({ range: `${TAB}!C19`, values: [[`=IFERROR(TEXT(COUNTIF(${CKL}!A4:A200,TRUE)/COUNTA(${CKL}!C4:C200),"0%"),"—")`]] });
  vals.push({ range: `${TAB}!B20`, values: [['OVERDUE TASKS']] });
  vals.push({ range: `${TAB}!C20`, values: [[`=IFERROR(COUNTIFS(${CKL}!A4:A200,FALSE,${CKL}!F4:F200,"<"&TODAY()),"—")`]] });
  vals.push({ range: `${TAB}!B21`, values: [['ACCOMMODATION NEEDED']] });
  vals.push({ range: `${TAB}!C21`, values: [[`=IFERROR(COUNTIF(${GL}!K4:K200,TRUE),"—")`]] });
  vals.push({ range: `${TAB}!B22`, values: [['ACCOMMODATION ASSIGNED']] });
  vals.push({ range: `${TAB}!C22`, values: [[`=IFERROR(COUNTIFS(${GL}!K4:K200,TRUE,${GL}!L4:L200,"<>"&""),"—")`]] });

  // ── Row 22: Divider + Upcoming Tasks ──
  reqs.push({ repeatCell: {
    range: gridRange(DASH,22,23,0,9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmAmber) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 22, endIndex: 23 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  reqs.push({ mergeCells: { range: gridRange(DASH,23,24,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,23,24,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 23, endIndex: 24 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A24`, values: [['  ⚡  UPCOMING TASKS (Next 21 Days)']] });

  // Upcoming tasks block (rows 24-30 = GSheets 25-31)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,24,32,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmCream),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 24, endIndex: 32 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // FILTER formula for upcoming tasks — merged B25:H31
  reqs.push({ mergeCells: { range: gridRange(DASH,24,32,1,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,24,32,1,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleGreen),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'TOP', wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,wrapStrategy)',
  }});
  vals.push({ range: `${TAB}!B25`, values: [[`=IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER(${CKL}!C4:C200,(${CKL}!A4:A200=FALSE)*(${CKL}!F4:F200<=TODAY()+21)*(${CKL}!C4:C200<>""))),"✅ No tasks due in the next 3 weeks!")`]] });

  // ── Family Branch Summary (rows 32-39 = GSheets 33-40) ──
  reqs.push({ repeatCell: {
    range: gridRange(DASH,32,33,0,9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmAmber) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 32, endIndex: 33 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  reqs.push({ mergeCells: { range: gridRange(DASH,33,34,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,33,34,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 33, endIndex: 34 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A34`, values: [['  👨‍👩‍👧‍👦  FAMILY BRANCH SUMMARY']] });

  // Branch summary table headers (0-indexed 34 = GSheets 35)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,34,35,1,5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 34, endIndex: 35 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B35`, values: [['FAMILY BRANCH','TOTAL GUESTS','CONFIRMED','AWAITING']] });

  // Branch summary rows
  reqs.push({ repeatCell: {
    range: gridRange(DASH,35,35+BRANCHES.length,1,5),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 35, endIndex: 35+BRANCHES.length },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  BRANCHES.forEach((branch, i) => {
    const bg = i % 2 === 0 ? C.white : C.warmCream;
    const row = 36 + i;
    reqs.push({ repeatCell: {
      range: gridRange(DASH,35+i,36+i,1,5),
      cell: { userEnteredFormat: { backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat.backgroundColor',
    }});
    vals.push({ range: `${TAB}!B${row}`, values: [[
      branch,
      `=IFERROR(COUNTIF(${GL}!B4:B200,"${branch}"),"—")`,
      `=IFERROR(COUNTIFS(${GL}!B4:B200,"${branch}",${GL}!H4:H200,"Confirmed"),"—")`,
      `=IFERROR(COUNTIFS(${GL}!B4:B200,"${branch}",${GL}!H4:H200,"No Response")+COUNTIFS(${GL}!B4:B200,"${branch}",${GL}!H4:H200,"Invited"),"—")`,
    ]] });
  });

  await batchUpdate(id, reqs, 'dashboard-format');
  await valuesBatchUpdate(id, vals, 'dashboard-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
