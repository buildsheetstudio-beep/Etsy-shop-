'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Fitness Dashboard'];
const S = "'Fitness Dashboard'";
const LOG = "'Daily Workout Log'";
const SETUP = "'Workout Setup'";

// Layout:
// ri=0  Title
// ri=1  Subtitle
// ri=2  Year/Month inputs
// ri=3  Spacer
// ri=4  KPI group 1 labels  (Sessions Year | Sessions Month | Sets Year | Volume Year)
// ri=5  KPI group 1 values
// ri=6  Spacer
// ri=7  KPI group 2 labels  (Best Bench | Best Deadlift | Best Squat | Avg RPE)
// ri=8  KPI group 2 values
// ri=9  Spacer
// ri=10 KPI group 3 labels  (Monthly Goal | Monthly Status | Exercises Tracked | Year-End Goal)
// ri=11 KPI group 3 values
// ri=12 Spacer
// ri=13 Section headers (Breakdown | Top PRs)
// ri=14 Sub-headers
// ri=15-19 5 data rows
// ri=20 Spacer
// ri=21 Recent log entries header
// ri=22 Column headers
// ri=23-32 10 log entries

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title — merged across all 20 cols, no frozenColumnCount
  vals.push({ range: `${S}!A1`, values: [['FITNESS DASHBOARD']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 18, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Auto-calculated from your Daily Workout Log. Change the year/month below to update monthly KPIs.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Row 3: Year/Month input
  vals.push({ range: `${S}!A3`, values: [['Reporting Year:']] });
  vals.push({ range: `${S}!B3`, values: [[2026]] });
  vals.push({ range: `${S}!C3`, values: [['Month (1–12):']] });
  vals.push({ range: `${S}!D3`, values: [[10]] });
  vals.push({ range: `${S}!E3`, values: [['← Change year/month numbers to update monthly KPIs']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,4,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,2,3), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,3,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,4,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // ri=3 Spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // KPI grid: 3 groups × 4 boxes each (5 cols per box = 20 cols)
  // labelRi = 4 + g*3, valueRi = labelRi+1, spacerRi = labelRi+2 (except last group)
  const KPIS = [
    // Group 0
    ['Sessions This Year',       `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$B$8:$B$5000,${LOG}!$B$8:$B$5000),0))`],
    ['Sessions This Month',      `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*(MONTH(${LOG}!$C$8:$C$5000)=$D$3)*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$B$8:$B$5000,${LOG}!$B$8:$B$5000),0))`],
    ['Total Sets This Year',     `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*(LEN(${LOG}!$B$8:$B$5000)>0)*1)`],
    ['Total Volume This Year (lb)', `=TEXT(SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*IFERROR(${LOG}!$AB$8:$AB$5000*1,0)),"#,##0")`],
    // Group 1
    ['Best Bench Press (lb)',    `=IFERROR(MAXIFS(${LOG}!$Q$8:$Q$5000,${LOG}!$K$8:$K$5000,"EX-0001"),"")`],
    ['Best Deadlift (lb)',       `=IFERROR(MAXIFS(${LOG}!$Q$8:$Q$5000,${LOG}!$K$8:$K$5000,"EX-0007"),"")`],
    ['Best Back Squat (lb)',     `=IFERROR(MAXIFS(${LOG}!$Q$8:$Q$5000,${LOG}!$K$8:$K$5000,"EX-0036"),"")`],
    ['Avg RPE This Year',        `=IFERROR(TEXT(SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*IFERROR(${LOG}!$W$8:$W$5000*1,0))/SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*(LEN(${LOG}!$W$8:$W$5000)>0)*1),"0.0"),"")`],
    // Group 2
    ['Monthly Sessions Goal',    `=IFERROR(${SETUP}!B12,"")`],
    ['Monthly Status',           `=IFERROR(IF(SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*(MONTH(${LOG}!$C$8:$C$5000)=$D$3)*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$B$8:$B$5000,${LOG}!$B$8:$B$5000),0))>=${SETUP}!B12,"On Track","Below Goal"),"")`],
    ['Exercises Tracked',        `=SUMPRODUCT((LEN(${LOG}!$K$8:$K$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$K$8:$K$5000,${LOG}!$K$8:$K$5000),0))`],
    ['Year-End Goal',            `=IFERROR(${SETUP}!B13,"")`],
  ];

  for (let g = 0; g < 3; g++) {
    const labelRi = 4 + g * 3;
    const valueRi = labelRi + 1;

    for (let b = 0; b < 4; b++) {
      const [label, formula] = KPIS[g * 4 + b];
      const c1 = b * 5;
      const c2 = c1 + 5;

      fmt.push({ mergeCells: { range: gridRange(SID, labelRi, labelRi+1, c1, c2), mergeType: 'MERGE_ALL' }});
      fmt.push({ mergeCells: { range: gridRange(SID, valueRi, valueRi+1, c1, c2), mergeType: 'MERGE_ALL' }});

      vals.push({ range: `${S}!${cl(c1)}${labelRi+1}`, values: [[label]] });
      fmt.push({ repeatCell: { range: gridRange(SID, labelRi, labelRi+1, c1, c2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary), textFormat: { fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat' }});

      vals.push({ range: `${S}!${cl(c1)}${valueRi+1}`, values: [[formula]] });
      fmt.push({ repeatCell: { range: gridRange(SID, valueRi, valueRi+1, c1, c2), cell: { userEnteredFormat: {
        backgroundColor: hex(b % 2 === 0 ? C.white : C.secondaryTint),
        textFormat: { bold: true, fontSize: 22, foregroundColor: hex(C.secondary), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat' }});
    }

    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: labelRi, endIndex: labelRi+1 }, properties: { pixelSize: 18 }, fields: 'pixelSize' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: valueRi, endIndex: valueRi+1 }, properties: { pixelSize: 58 }, fields: 'pixelSize' }});

    // Spacer after group (except last)
    if (g < 2) {
      fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: valueRi+1, endIndex: valueRi+2 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});
    }
  }

  // ri=12: Spacer after KPI block
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 12 }, fields: 'pixelSize' }});

  // ri=13: Section headers — left = Workout Type Breakdown, right = Top Personal Records
  vals.push({ range: `${S}!A14`, values: [['WORKOUT TYPE BREAKDOWN (THIS YEAR)']] });
  vals.push({ range: `${S}!K14`, values: [['TOP PERSONAL RECORDS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,13,14,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ mergeCells: { range: gridRange(SID,13,14,10,20), mergeType: 'MERGE_ALL' }});
  [0, 10].forEach(c => fmt.push({ repeatCell: { range: gridRange(SID,13,14,c,c+10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }}));
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ri=14: Sub-headers
  vals.push({ range: `${S}!A15`, values: [['Type']] });
  vals.push({ range: `${S}!B15`, values: [['Sessions This Year']] });
  vals.push({ range: `${S}!G15`, values: [['Sessions This Month']] });
  vals.push({ range: `${S}!K15`, values: [['Exercise']] });
  vals.push({ range: `${S}!P15`, values: [['Best Value']] });
  fmt.push({ mergeCells: { range: gridRange(SID,14,15,1,6), mergeType: 'MERGE_ALL' }});
  fmt.push({ mergeCells: { range: gridRange(SID,14,15,6,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ mergeCells: { range: gridRange(SID,14,15,10,15), mergeType: 'MERGE_ALL' }});
  fmt.push({ mergeCells: { range: gridRange(SID,14,15,15,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,14,15,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 14, endIndex: 15 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ri=15-19: 5 workout types / 5 top exercises
  const TYPES = ['Push', 'Pull', 'Legs', 'Cardio', 'HIIT'];
  const TOP_EX = [
    ['Bench Press',     'EX-0001', 'Q'],
    ['Deadlift',        'EX-0007', 'Q'],
    ['Back Squat',      'EX-0036', 'Q'],
    ['Overhead Press',  'EX-0014', 'Q'],
    ['Pull-Up',         'EX-0009', 'R'],
  ];

  for (let i = 0; i < 5; i++) {
    const ri = 15 + i;
    const r = ri + 1;
    const type = TYPES[i];
    const bg = i % 2 === 0 ? C.white : C.altRow;

    // Merges for this row
    fmt.push({ mergeCells: { range: gridRange(SID,ri,ri+1,1,6), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,ri,ri+1,6,10), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,ri,ri+1,10,15), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,ri,ri+1,15,20), mergeType: 'MERGE_ALL' }});

    // Left panel: workout type
    const daysYear = `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*(${LOG}!$I$8:$I$5000="${type}")*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$B$8:$B$5000,${LOG}!$B$8:$B$5000),0))`;
    const daysMo  = `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$3)*(MONTH(${LOG}!$C$8:$C$5000)=$D$3)*(${LOG}!$I$8:$I$5000="${type}")*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$B$8:$B$5000,${LOG}!$B$8:$B$5000),0))`;
    vals.push({ range: `${S}!A${r}`, values: [[type]] });
    vals.push({ range: `${S}!B${r}`, values: [[daysYear]] });
    vals.push({ range: `${S}!G${r}`, values: [[daysMo]] });

    // Right panel: exercise PR
    const [exName, exId, qrCol] = TOP_EX[i];
    const bestVal = `=IFERROR(MAXIFS(${LOG}!$${qrCol}$8:$${qrCol}$5000,${LOG}!$K$8:$K$5000,"${exId}"),"")`;
    vals.push({ range: `${S}!K${r}`, values: [[exName]] });
    vals.push({ range: `${S}!P${r}`, values: [[bestVal]] });

    // Formats
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,10), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,1), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat(textFormat)' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,1,6), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.secondary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,6,10), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.accent), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,10,20), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,10,15), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat(textFormat)' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,15,20), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.success), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  }

  // ri=20: Spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 20, endIndex: 21 }, properties: { pixelSize: 12 }, fields: 'pixelSize' }});

  // ri=21: Recent log entries header
  vals.push({ range: `${S}!A22`, values: [['RECENT LOG ENTRIES (LAST 10 SETS)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,21,22,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,21,22,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 21, endIndex: 22 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ri=22: Column headers
  vals.push({ range: `${S}!A23`, values: [['Date','Session ID','Exercise Name','Type','Weight (lb)','Reps','Volume (lb)']] });
  fmt.push({ repeatCell: { range: gridRange(SID,22,23,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 22, endIndex: 23 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // ri=23-32: Last 10 log entries (most recent first)
  // COUNTA(LOG!$B$8:$B$5000)-i gives position of (i+1)th from last entry
  for (let i = 0; i < 10; i++) {
    const ri = 23 + i;
    const r = ri + 1; // spreadsheet row 24..33
    const idx = `COUNTA(${LOG}!$B$8:$B$5000)-${i}`;

    vals.push({ range: `${S}!A${r}`, values: [[`=IFERROR(TEXT(INDEX(${LOG}!$C$8:$C$5000,${idx}),"yyyy-mm-dd"),"")`]] });
    vals.push({ range: `${S}!B${r}`, values: [[`=IFERROR(INDEX(${LOG}!$B$8:$B$5000,${idx}),"")`]] });
    vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(INDEX(${LOG}!$L$8:$L$5000,${idx}),"")`]] });
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(INDEX(${LOG}!$I$8:$I$5000,${idx}),"")`]] });
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(INDEX(${LOG}!$Q$8:$Q$5000,${idx}),"")`]] });
    vals.push({ range: `${S}!F${r}`, values: [[`=IFERROR(INDEX(${LOG}!$R$8:$R$5000,${idx}),"")`]] });
    vals.push({ range: `${S}!G${r}`, values: [[`=IFERROR(INDEX(${LOG}!$AB$8:$AB$5000,${idx}),"")`]] });

    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,20), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  }

  // Column widths (20 cols)
  const WIDTHS = [90,120,160,90,80,60,100,80,80,80,140,80,80,80,80,100,80,80,80,80];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-3 — no frozenColumnCount (title spans all 20 cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '10-dashboard values');
  await batchUpdate(id, fmt, '10-dashboard format');
  console.log('✅  Fitness Dashboard done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function cl(ci) { return String.fromCharCode(65 + ci); }
