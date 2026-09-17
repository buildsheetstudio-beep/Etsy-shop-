'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RES = sheetMap['🏆 Results & Awards Log'];

const TRK = "'📋 Application Tracker'";

const TYPES = ['University','Government','Private Foundation','Corporate','Community','Research Grant','Bursary','Fellowship','Award','Other'];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB  = "'🏆 Results & Awards Log'";
  const NCOLS = 8; // A-H

  // Column widths
  [240,180,120,120,120,120,220,160].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: RES, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // ── Row 0: Title banner ──
  reqs.push({ mergeCells: { range: gridRange(RES,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(RES,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  🏆  Results & Awards Log — Scholarship & Bursary Tracker']] });

  // ── Row 1: Section 1 subheader ──
  reqs.push({ mergeCells: { range: gridRange(RES,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(RES,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['  📨  RESULTS LOG — Auto-populated from Application Tracker']] });

  // ── Row 2: Note ──
  reqs.push({ mergeCells: { range: gridRange(RES,2,3,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(RES,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      textFormat: { foregroundColor: hex(C.slateBlue), italic: true, fontSize: 9 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [['  ℹ️  Results auto-populate from Application Tracker for all Submitted, Awarded, Rejected, and Waitlisted applications.']] });

  // ── Row 3: Column headers ──
  const hdrs = ['SCHOLARSHIP NAME','PROVIDER','AMOUNT APPLIED FOR','STATUS','AMOUNT AWARDED','RESULT DATE','NOTES','TYPE'];
  reqs.push({ repeatCell: {
    range: gridRange(RES,3,4,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A4`, values: [hdrs] });

  // ── FILTER formulas (row 4, 0-indexed, GSheets row 5) ──
  const statusCond = `((${TRK}!G4:G200="Submitted")+(${TRK}!G4:G200="Awarded")+(${TRK}!G4:G200="Rejected")+(${TRK}!G4:G200="Waitlisted"))*(${TRK}!A4:A200<>"")`;

  // Style FILTER output area rows 4-20
  reqs.push({ repeatCell: {
    range: gridRange(RES,4,21,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Amount applied (col C, index 2): currency
  reqs.push({ repeatCell: {
    range: gridRange(RES,4,21,2,3),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } }},
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Amount awarded (col E, index 4): currency
  reqs.push({ repeatCell: {
    range: gridRange(RES,4,21,4,5),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } }},
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Result date (col F, index 5): date
  reqs.push({ repeatCell: {
    range: gridRange(RES,4,21,5,6),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } }},
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: 4, endIndex: 21 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Write FILTER formulas at GSheets row 5 (0-indexed row 4)
  vals.push({ range: `${TAB}!A5`, values: [[`=IFERROR(FILTER(${TRK}!A4:A200,${statusCond}),"No results yet")`]] });
  vals.push({ range: `${TAB}!B5`, values: [[`=IFERROR(FILTER(${TRK}!B4:B200,${statusCond}),"")`]] });
  vals.push({ range: `${TAB}!C5`, values: [[`=IFERROR(FILTER(${TRK}!D4:D200,${statusCond}),"")`]] });
  vals.push({ range: `${TAB}!D5`, values: [[`=IFERROR(FILTER(${TRK}!G4:G200,${statusCond}),"")`]] });
  vals.push({ range: `${TAB}!E5`, values: [[`=IFERROR(FILTER(${TRK}!M4:M200,${statusCond}),"")`]] });
  vals.push({ range: `${TAB}!F5`, values: [[`=IFERROR(FILTER(${TRK}!N4:N200,${statusCond}),"")`]] });
  vals.push({ range: `${TAB}!G5`, values: [[`=IFERROR(FILTER(${TRK}!O4:O200,${statusCond}),"")`]] });
  vals.push({ range: `${TAB}!H5`, values: [[`=IFERROR(FILTER(${TRK}!C4:C200,${statusCond}),"")`]] });

  // ── Divider ──
  const divRi = 21;
  reqs.push({ repeatCell: {
    range: gridRange(RES, divRi, divRi+1, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: divRi, endIndex: divRi+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // ── Row 22: Section 2 subheader ──
  const sec2Ri = 22;
  reqs.push({ mergeCells: { range: gridRange(RES,sec2Ri,sec2Ri+1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(RES,sec2Ri,sec2Ri+1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: sec2Ri, endIndex: sec2Ri+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A23`, values: [['  🏅  AWARD SUMMARY']] });

  // ── KPI rows 23-28 (0-indexed 23-28, GSheets 24-29) ──
  const kpiRows = [
    ['TOTAL AWARDED THIS YEAR',   `=IFERROR(SUMIF(${TRK}!G:G,"Awarded",${TRK}!M:M),0)`],
    ['ANNUAL FUNDING GOAL',       `=IFERROR('🎓 Dashboard'!B5,"—")`],
    ['FUNDING GAP',               `=IFERROR(B24-B23,"—")`],
    ['APPLICATIONS SUBMITTED',    `=IFERROR(COUNTIF(${TRK}!G:G,"Submitted")+COUNTIF(${TRK}!G:G,"Awarded")+COUNTIF(${TRK}!G:G,"Rejected"),"—")`],
    ['SUCCESS RATE',              `=IFERROR(COUNTIF(${TRK}!G:G,"Awarded")/(COUNTIF(${TRK}!G:G,"Awarded")+COUNTIF(${TRK}!G:G,"Rejected")),"—")`],
    ['AVERAGE AWARD SIZE',        `=IFERROR(AVERAGEIF(${TRK}!G:G,"Awarded",${TRK}!M:M),"—")`],
  ];

  kpiRows.forEach(([label, formula], i) => {
    const ri  = 23 + i;
    const row = ri + 1;
    // Label col A
    reqs.push({ repeatCell: {
      range: gridRange(RES, ri, ri+1, 0, 1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.paleGold),
        textFormat: { foregroundColor: hex(C.deepBlue), bold: true, fontSize: 9 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Value col B
    reqs.push({ repeatCell: {
      range: gridRange(RES, ri, ri+1, 1, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formulaBg),
        textFormat: { foregroundColor: hex(C.gold), bold: true, fontSize: 13 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        numberFormat: (i === 0 || i === 1 || i === 2 || i === 5) ? { type: 'CURRENCY', pattern: '$#,##0.00' } :
                      (i === 4) ? { type: 'PERCENT', pattern: '0.0%' } : undefined,
      }},
      fields: `userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment${(i===0||i===1||i===2||i===4||i===5)?',numberFormat':''})`,
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: RES, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [[label, formula]] });
  });

  // ── Row 29: Type breakdown header (0-indexed 29, GSheets 30) ──
  const typeHdrRi = 29;
  reqs.push({ repeatCell: {
    range: gridRange(RES, typeHdrRi, typeHdrRi+1, 0, 3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: typeHdrRi, endIndex: typeHdrRi+1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A30`, values: [['SCHOLARSHIP TYPE','COUNT AWARDED','TOTAL AWARDED ($)']] });

  // Type data rows (0-indexed 30-39, GSheets 31-40)
  TYPES.forEach((type, i) => {
    const ri  = 30 + i;
    const row = ri + 1;
    const bg  = i % 2 === 0 ? C.white : C.paleBlue;
    reqs.push({ repeatCell: {
      range: gridRange(RES, ri, ri+1, 0, 3),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(RES, ri, ri+1, 2, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } }},
      fields: 'userEnteredFormat.numberFormat',
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: RES, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [[
      type,
      `=IFERROR(COUNTIFS(${TRK}!C:C,"${type}",${TRK}!G:G,"Awarded"),0)`,
      `=IFERROR(SUMIFS(${TRK}!M:M,${TRK}!C:C,"${type}",${TRK}!G:G,"Awarded"),0)`,
    ]] });
  });

  // ── Row 40: Monthly sparkline (0-indexed 40, GSheets 41) ──
  const sparkRi = 40;
  reqs.push({ mergeCells: { range: gridRange(RES, sparkRi, sparkRi+1, 0, 3), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(RES, sparkRi, sparkRi+1, 0, 3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: sparkRi, endIndex: sparkRi+1 },
    properties: { pixelSize: 60 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(RES, sparkRi-1, sparkRi, 0, 3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RES, dimension: 'ROWS', startIndex: sparkRi-1, endIndex: sparkRi },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A40`, values: [['MONTHLY AWARDS TREND','','']] });
  vals.push({ range: `${TAB}!A41`, values: [[`=SPARKLINE(ARRAYFORMULA(SUMIF(TEXT(${TRK}!N:N,"MMM"),{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},${TRK}!M:M)),{"charttype","column";"color","#C9A800"})`, '', '']] });

  await batchUpdate(id, reqs, 'results-format');
  await valuesBatchUpdate(id, vals, 'results-values');
  console.log('Results & Awards Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
