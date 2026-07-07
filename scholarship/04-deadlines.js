'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DL = sheetMap['⏰ Deadlines & To-Do'];

const TRK_NAME = "'📋 Application Tracker'";

(async () => {
  const reqs = [];
  const vals = [];
  const TAB  = "'⏰ Deadlines & To-Do'";
  const NCOLS = 8; // A-H

  // Column widths
  [240, 180, 110, 120, 90, 120, 90, 180].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // ── Row 0: Title banner ──
  reqs.push({ mergeCells: { range: gridRange(DL,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DL,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  ⏰  Deadlines & To-Do — Scholarship & Bursary Tracker']] });

  // ── Row 1: Section 1 subheader ──
  reqs.push({ mergeCells: { range: gridRange(DL,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DL,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['  📅  UPCOMING DEADLINES (Next 90 Days) — Auto-updates from Application Tracker']] });

  // ── Row 2: Note ──
  reqs.push({ mergeCells: { range: gridRange(DL,2,3,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DL,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      textFormat: { foregroundColor: hex(C.slateBlue), italic: true, fontSize: 9 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [['  ℹ️  This view auto-updates from the Application Tracker. No manual entry needed. Sorted by deadline in tracker order.']] });

  // ── Row 3: Column headers ──
  const filterHdrs = ['SCHOLARSHIP NAME','PROVIDER','AMOUNT','DEADLINE','DAYS LEFT','STATUS','PRIORITY','REQUIREMENTS'];
  reqs.push({ repeatCell: {
    range: gridRange(DL,3,4,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A4`, values: [filterHdrs] });

  // ── FILTER formulas (row 4, 0-indexed, GSheets row 5) ──
  // Common filter conditions: deadline > TODAY, deadline <= TODAY+90, status not in terminal set, name not empty
  const cond = `(${TRK_NAME}!F4:F200>TODAY())*(${TRK_NAME}!F4:F200<=TODAY()+90)*(${TRK_NAME}!G4:G200<>"Submitted")*(${TRK_NAME}!G4:G200<>"Awarded")*(${TRK_NAME}!G4:G200<>"Rejected")*(${TRK_NAME}!G4:G200<>"Withdrawn")*(${TRK_NAME}!A4:A200<>"")`;

  const filterFormulas = [
    `=IFERROR(FILTER(${TRK_NAME}!A4:A200,${cond}),"No deadlines in next 90 days")`, // A — name
    `=IFERROR(FILTER(${TRK_NAME}!B4:B200,${cond}),"")`,  // B — provider
    `=IFERROR(FILTER(${TRK_NAME}!D4:D200,${cond}),"")`,  // C — amount
    `=IFERROR(FILTER(${TRK_NAME}!F4:F200,${cond}),"")`,  // D — deadline
    `=IFERROR(FILTER(${TRK_NAME}!F4:F200-TODAY(),${cond}),"")`, // E — days left
    `=IFERROR(FILTER(${TRK_NAME}!G4:G200,${cond}),"")`,  // F — status
    `=IFERROR(FILTER(${TRK_NAME}!H4:H200,${cond}),"")`,  // G — priority
    `=IFERROR(FILTER(${TRK_NAME}!J4:J200,${cond}),"")`,  // H — requirements
  ];

  // Style formula rows 4-20 (FILTER output area)
  reqs.push({ repeatCell: {
    range: gridRange(DL,4,20,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Amount (col C, index 2) currency format
  reqs.push({ repeatCell: {
    range: gridRange(DL,4,20,2,3),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } }},
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Deadline (col D, index 3) date format
  reqs.push({ repeatCell: {
    range: gridRange(DL,4,20,3,4),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } }},
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Days left (col E, index 4) number + center
  reqs.push({ repeatCell: {
    range: gridRange(DL,4,20,4,5),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0' } }},
    fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 4, endIndex: 20 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Write filter formulas
  filterFormulas.forEach((f, i) => {
    const col = String.fromCharCode(65 + i);
    vals.push({ range: `${TAB}!${col}5`, values: [[f]] });
  });

  // ── Divider between sections ──
  const divRi = 20; // 0-indexed
  reqs.push({ repeatCell: {
    range: gridRange(DL, divRi, divRi+1, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: divRi, endIndex: divRi+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // ── Row 21: Section 2 subheader ──
  const sec2Ri = 21;
  reqs.push({ mergeCells: { range: gridRange(DL,sec2Ri,sec2Ri+1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DL,sec2Ri,sec2Ri+1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: sec2Ri, endIndex: sec2Ri+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A22`, values: [['  ✅  APPLICATION TO-DO CHECKLIST — Track tasks per scholarship']] });

  // ── Row 22: Note ──
  reqs.push({ mergeCells: { range: gridRange(DL,22,23,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DL,22,23,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      textFormat: { foregroundColor: hex(C.slateBlue), italic: true, fontSize: 9 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 22, endIndex: 23 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A23`, values: [['  ℹ️  Enter scholarship names in the gold cells (row 24). Tick checkboxes as you complete each task.']] });

  // ── Row 23: Scholarship name inputs (0-indexed 23, GSheets 24) ──
  // Col A: label; Cols B-G: 6 scholarship name inputs
  reqs.push({ repeatCell: {
    range: gridRange(DL,23,24,0,1),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Gold border input cells for scholarship names
  const goldBorder = { style: 'SOLID', color: hex(C.gold), width: 1 };
  reqs.push({ repeatCell: {
    range: gridRange(DL,23,24,1,7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleGold),
      textFormat: { foregroundColor: hex(C.bodyText), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Apply gold borders to B24:G24 (0-indexed cols 1-6, row 23)
  for (let ci = 1; ci <= 6; ci++) {
    reqs.push({ updateBorders: {
      range: gridRange(DL,23,24,ci,ci+1),
      top: goldBorder, bottom: goldBorder, left: goldBorder, right: goldBorder,
    }});
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 23, endIndex: 24 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A24`, values: [['SCHOLARSHIP →','Westpac Future Leaders','Aus. Research Training','Ian Potter Foundation','Myer Foundation Grant','Uni Sports Excellence','NAB Village Scholarship']] });

  // ── Rows 24-31: Task rows (0-indexed 24-31, GSheets 25-32) ──
  const TASKS = [
    'Research & eligibility check',
    'Personal statement / essay drafted',
    'Personal statement / essay final',
    'Academic transcript obtained',
    'Reference 1 requested',
    'Reference 2 requested',
    'Application form completed',
    'Application submitted',
  ];

  TASKS.forEach((task, i) => {
    const ri  = 24 + i;
    // Col A: task label
    reqs.push({ repeatCell: {
      range: gridRange(DL, ri, ri+1, 0, 1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.deepBlue),
        textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Cols B-G: checkbox cells (white bg + light blue border)
    reqs.push({ repeatCell: {
      range: gridRange(DL, ri, ri+1, 1, 7),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.white),
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment)',
    }});
    for (let ci = 1; ci <= 6; ci++) {
      const lightBlueBorder = { style: 'SOLID', color: hex('#9EB8F0'), width: 1 };
      reqs.push({ updateBorders: {
        range: gridRange(DL, ri, ri+1, ci, ci+1),
        top: lightBlueBorder, bottom: lightBlueBorder, left: lightBlueBorder, right: lightBlueBorder,
      }});
    }
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DL, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    }});
    vals.push({ range: `${TAB}!A${ri+1}`, values: [[task, 'FALSE','FALSE','FALSE','FALSE','FALSE','FALSE']] });
  });

  // Checkbox validation on B25:G32 (0-indexed rows 24-31, cols 1-6)
  reqs.push({ setDataValidation: {
    range: gridRange(DL, 24, 32, 1, 7),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
  }});

  // ── Progress bar row 32 (0-indexed, GSheets 33) ──
  reqs.push({ repeatCell: {
    range: gridRange(DL,32,33,0,1),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DL,32,33,1,7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 32, endIndex: 33 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A33`, values: [['PROGRESS']] });
  // Progress bar formula for each column B-G (cols 1-6)
  const progFormulas = ['B','C','D','E','F','G'].map(col =>
    `=SPARKLINE(IFERROR(COUNTIF(${col}25:${col}32,TRUE)/8,0),{"charttype","bar";"color1","#1A56C4";"max",1})`
  );
  vals.push({ range: `${TAB}!B33`, values: [progFormulas] });

  // ── % label row 33 (0-indexed, GSheets 34) ──
  reqs.push({ repeatCell: {
    range: gridRange(DL,33,34,0,1),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DL,33,34,1,7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formulaBg),
      textFormat: { foregroundColor: hex(C.royalBlue), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DL, dimension: 'ROWS', startIndex: 33, endIndex: 34 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A34`, values: [['% DONE']] });
  const pctFormulas = ['B','C','D','E','F','G'].map(col =>
    `=IFERROR(TEXT(COUNTIF(${col}25:${col}32,TRUE)/8,"0%")&" complete","")`
  );
  vals.push({ range: `${TAB}!B34`, values: [pctFormulas] });

  await batchUpdate(id, reqs, 'deadlines-format');
  await valuesBatchUpdate(id, vals, 'deadlines-values');
  console.log('Deadlines & To-Do complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
