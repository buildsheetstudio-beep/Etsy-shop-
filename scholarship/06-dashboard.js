'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🎓 Dashboard'];

const TRK = "'📋 Application Tracker'";

(async () => {
  const reqs = [];
  const vals = [];
  const TAB  = "'🎓 Dashboard'";
  const NCOLS = 10; // A-J

  // Column widths: A=40(label indent), B=180, C=20(gap), D=200, E=180, F=20, G=200, H-J extra
  [40,200,30,200,180,30,200,120,120,120].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // ── Row 0: Title banner ──
  reqs.push({ mergeCells: { range: gridRange(DSH,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  🎓  Scholarship & Bursary Tracker — BuildSheetStudio']] });

  // ── Row 1: Thin gold accent ──
  reqs.push({ repeatCell: {
    range: gridRange(DSH,1,2,0,NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 6 }, fields: 'pixelSize',
  }});

  // ── INPUT BLOCK (rows 2-7, GSheets 3-8, 0-indexed 2-7) ──
  // Left side: labels in col B, inputs in col C-D
  const inputSubRi = 2;
  reqs.push({ mergeCells: { range: gridRange(DSH, inputSubRi, inputSubRi+1, 0, 4), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, inputSubRi, inputSubRi+1, 0, 4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: inputSubRi, endIndex: inputSubRi+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [['  👤  YOUR DETAILS']] });

  const INPUT_ROWS = [
    ['STUDENT NAME',          'Alex Chen'],
    ['INSTITUTION',           'University of Melbourne'],
    ['ACADEMIC YEAR',         '2025–2026'],
    ['ANNUAL FUNDING GOAL',   15000],
  ];
  const blueBorder = { style: 'SOLID', color: hex(C.royalBlue), width: 1 };
  INPUT_ROWS.forEach(([label, value], i) => {
    const ri  = 3 + i;
    const row = ri + 1;
    // Label col B (index 1)
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 1, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.paleGold),
        textFormat: { foregroundColor: hex(C.deepBlue), bold: true, fontSize: 9 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Input cell col C-D merged (index 2-3)
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri+1, 2, 4), mergeType: 'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 2, 4),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.inputBg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 10 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        ...(i === 3 ? { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } : {}),
      }},
      fields: `userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment${i===3?',numberFormat':''})`,
    }});
    reqs.push({ updateBorders: {
      range: gridRange(DSH, ri, ri+1, 2, 4),
      top: blueBorder, bottom: blueBorder, left: blueBorder, right: blueBorder,
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!B${row}`, values: [[label, value]] });
  });

  // ── KPI BLOCK — Right side (cols E-G, rows 3-8) ──
  // Subheader
  const kpiSubRi = 2;
  reqs.push({ mergeCells: { range: gridRange(DSH, kpiSubRi, kpiSubRi+1, 4, 8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, kpiSubRi, kpiSubRi+1, 4, 8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `${TAB}!E3`, values: [['  📊  APPLICATION STATS']] });

  const KPI_ROWS = [
    ['TOTAL APPLICATIONS', `=IFERROR(COUNTA(${TRK}!A4:A200),"—")`],
    ['SUBMITTED',          `=IFERROR(COUNTIF(${TRK}!G:G,"Submitted"),"—")`],
    ['AWARDED',            `=IFERROR(COUNTIF(${TRK}!G:G,"Awarded"),"—")`],
    ['REJECTED',           `=IFERROR(COUNTIF(${TRK}!G:G,"Rejected"),"—")`],
    ['IN PROGRESS',        `=IFERROR(COUNTIF(${TRK}!G:G,"In Progress"),"—")`],
    ['NOT STARTED',        `=IFERROR(COUNTIF(${TRK}!G:G,"Not Started"),"—")`],
  ];
  KPI_ROWS.forEach(([label, formula], i) => {
    const ri  = 3 + i;
    const row = ri + 1;
    // Label col E (index 4)
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.paleGold),
        textFormat: { foregroundColor: hex(C.deepBlue), bold: true, fontSize: 9 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Value col F (index 5)
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 5, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formulaBg),
        textFormat: { foregroundColor: hex(C.gold), bold: true, fontSize: 14 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!E${row}`, values: [[label, formula]] });
  });

  // ── FINANCIAL BLOCK (rows 9-15, 0-indexed 9-15) ──
  // Divider
  reqs.push({ repeatCell: {
    range: gridRange(DSH,8,9,0,NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBlue) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 6 }, fields: 'pixelSize',
  }});

  // Financial subheader
  const finSubRi = 9;
  reqs.push({ mergeCells: { range: gridRange(DSH, finSubRi, finSubRi+1, 0, 4), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, finSubRi, finSubRi+1, 0, 4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: finSubRi, endIndex: finSubRi+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A10`, values: [['  💰  FINANCIAL OVERVIEW']] });

  const FIN_ROWS = [
    ['AMOUNT APPLIED FOR', `=IFERROR(SUMIF(${TRK}!G:G,"<>Not Started",${TRK}!D:D),"—")`, 'CURRENCY'],
    ['TOTAL AWARDED',      `=IFERROR(SUMIF(${TRK}!G:G,"Awarded",${TRK}!M:M),"—")`,       'CURRENCY'],
    ['FUNDING GAP',        `=IFERROR(B6-B11,"—")`,                                         'CURRENCY'],
    ['SUCCESS RATE',       `=IFERROR(COUNTIF(${TRK}!G:G,"Awarded")/(COUNTIF(${TRK}!G:G,"Awarded")+COUNTIF(${TRK}!G:G,"Rejected")),"—")`, 'PERCENT'],
    ['FUNDING PROGRESS',   `=SPARKLINE(IFERROR(B11/B6,0),{"charttype","bar";"color1","#C9A800";"max",1})`, 'SPARKLINE'],
  ];
  FIN_ROWS.forEach(([label, formula, fmt], i) => {
    const ri  = 10 + i;
    const row = ri + 1;
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 1, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.paleGold),
        textFormat: { foregroundColor: hex(C.deepBlue), bold: true, fontSize: 9 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri+1, 2, 4), mergeType: 'MERGE_ALL' }});
    const numFmt = fmt === 'CURRENCY' ? { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } }
                 : fmt === 'PERCENT'  ? { numberFormat: { type: 'PERCENT', pattern: '0.0%' } }
                 : {};
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 2, 4),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formulaBg),
        textFormat: { foregroundColor: hex(C.gold), bold: true, fontSize: fmt === 'SPARKLINE' ? 9 : 14 },
        horizontalAlignment: fmt === 'SPARKLINE' ? 'LEFT' : 'CENTER', verticalAlignment: 'MIDDLE',
        ...numFmt,
      }},
      fields: `userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment${fmt!=='SPARKLINE'&&fmt?',numberFormat':''})`,
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!B${row}`, values: [[label, formula]] });
  });

  // ── DEADLINE BLOCK (rows 15-17 subheader, 17-23 filter) ──
  // Divider
  reqs.push({ repeatCell: {
    range: gridRange(DSH,15,16,0,NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBlue) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 15, endIndex: 16 },
    properties: { pixelSize: 6 }, fields: 'pixelSize',
  }});

  // Next deadline KPI
  const dlSubRi = 16;
  reqs.push({ mergeCells: { range: gridRange(DSH, dlSubRi, dlSubRi+1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, dlSubRi, dlSubRi+1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: dlSubRi, endIndex: dlSubRi+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A17`, values: [['  ⏰  DEADLINE TRACKER']] });

  // Next deadline row
  const dlRows = [
    ['NEXT DEADLINE',       `=IFERROR(TEXT(MINIFS(${TRK}!F:F,${TRK}!F:F,">"&TODAY(),${TRK}!G:G,"Not Started"),"DD MMM YYYY")&" — "&IFERROR(INDEX(${TRK}!A:A,MATCH(MINIFS(${TRK}!F:F,${TRK}!F:F,">"&TODAY(),${TRK}!G:G,"Not Started"),${TRK}!F:F,0)),""),"No upcoming deadlines")`],
    ['DAYS TO NEXT',        `=IFERROR(MINIFS(${TRK}!F:F,${TRK}!F:F,">"&TODAY(),${TRK}!G:G,"Not Started")-TODAY(),"—")`],
  ];
  dlRows.forEach(([label, formula], i) => {
    const ri  = 17 + i;
    const row = ri + 1;
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 1, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.paleGold),
        textFormat: { foregroundColor: hex(C.deepBlue), bold: true, fontSize: 9 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri+1, 2, 6), mergeType: 'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 2, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formulaBg),
        textFormat: { foregroundColor: hex(C.gold), bold: true, fontSize: 12 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!B${row}`, values: [[label, formula]] });
  });

  // Urgent deadlines panel subheader
  const urgSubRi = 19;
  reqs.push({ mergeCells: { range: gridRange(DSH, urgSubRi, urgSubRi+1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, urgSubRi, urgSubRi+1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: urgSubRi, endIndex: urgSubRi+1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A20`, values: [['  🚨  DEADLINES IN THE NEXT 30 DAYS']] });

  // Urgent deadline filter headers
  const urgHdrRi = 20;
  ['SCHOLARSHIP','DEADLINE','DAYS LEFT','STATUS'].forEach((h, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(DSH, urgHdrRi, urgHdrRi+1, i+1, i+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.deepBlue),
        textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: urgHdrRi, endIndex: urgHdrRi+1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Filter formula rows 21-27 (0-indexed, GSheets 22-28) — styled as formula cells
  reqs.push({ repeatCell: {
    range: gridRange(DSH,21,28,1,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 21, endIndex: 28 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  const urgCond = `(${TRK}!F4:F200-TODAY()<=30)*(${TRK}!F4:F200>=TODAY())*(${TRK}!G4:G200<>"Submitted")*(${TRK}!G4:G200<>"Awarded")*(${TRK}!G4:G200<>"Rejected")*(${TRK}!G4:G200<>"Withdrawn")*(${TRK}!A4:A200<>"")`;
  vals.push({ range: `${TAB}!B22`, values: [[`=IFERROR(FILTER(${TRK}!A4:A200,${urgCond}),"No deadlines in next 30 days ✅")`]] });
  vals.push({ range: `${TAB}!C22`, values: [[`=IFERROR(FILTER(${TRK}!F4:F200,${urgCond}),"")`]] });
  vals.push({ range: `${TAB}!D22`, values: [[`=IFERROR(FILTER(${TRK}!F4:F200-TODAY(),${urgCond}),"")`]] });
  vals.push({ range: `${TAB}!E22`, values: [[`=IFERROR(FILTER(${TRK}!G4:G200,${urgCond}),"")`]] });
  // Date format for C22
  reqs.push({ repeatCell: {
    range: gridRange(DSH,21,28,2,3),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } }},
    fields: 'userEnteredFormat.numberFormat',
  }});

  // ── STATUS CHART DATA BLOCK (rows 27-35, cols B-C — for donut chart) ──
  const statusChartRi = 28;
  reqs.push({ repeatCell: {
    range: gridRange(DSH, statusChartRi-1, statusChartRi, 1, 3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: statusChartRi-1, endIndex: statusChartRi },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B28`, values: [['STATUS','COUNT']] });

  const STATUSES = ['Not Started','Researching','In Progress','Submitted','Awarded','Rejected','Waitlisted','Withdrawn'];
  STATUSES.forEach((s, i) => {
    const ri  = statusChartRi + i;
    const row = ri + 1;
    reqs.push({ repeatCell: {
      range: gridRange(DSH, ri, ri+1, 1, 3),
      cell: { userEnteredFormat: {
        backgroundColor: hex(i % 2 === 0 ? C.white : C.paleBlue),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!B${row}`, values: [[s, `=IFERROR(COUNTIF(${TRK}!G:G,"${s}"),0)`]] });
  });

  await batchUpdate(id, reqs, 'dashboard-format');
  await valuesBatchUpdate(id, vals, 'dashboard-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
