'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Monthly Calendar'];
const S = "'Monthly Calendar'";
const LOG = "'Daily Workout Log'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title (no frozenColumnCount — spans all 16 cols)
  vals.push({ range: `${S}!A1`, values: [['MONTHLY WORKOUT CALENDAR']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Month/Year selectors
  vals.push({ range: `${S}!A2`, values: [['Year:']] });
  vals.push({ range: `${S}!B2`, values: [[2026]] });
  vals.push({ range: `${S}!C2`, values: [['Month (1–12):']] });
  vals.push({ range: `${S}!D2`, values: [[10]] }); // default Oct 2026 → shows Oct 2025 sample data
  vals.push({ range: `${S}!E2`, values: [['← Change year/month numbers to navigate']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,4,12), mergeType: 'MERGE_ALL' }});
  [0,1,2,3].forEach(ci => {
    const bg = ci % 2 === 0 ? C.altRow : C.input;
    const bold = true;
    fmt.push({ repeatCell: { range: gridRange(SID,1,2,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { bold, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(ci % 2 === 0 ? C.primary : C.secondary) },
      horizontalAlignment: ci % 2 === 0 ? 'RIGHT' : 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,4,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Row 3: Day-of-week headers (Mon-Sun across cols A-G, col H = Summary)
  vals.push({ range: `${S}!A3`, values: [['Mon','Tue','Wed','Thu','Fri','Sat','Sun','MONTHLY SUMMARY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,7,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,7,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Calendar grid: rows 4-9 (6 weeks × 7 days) = index 3-8
  // First calendar cell (A4): Monday of the week containing the 1st of the month
  const firstCell = `DATE($B$2,$D$2,1)-(WEEKDAY(DATE($B$2,$D$2,1),2)-1)`;
  for (let week = 0; week < 6; week++) {
    const ri = 3 + week;
    for (let day = 0; day < 7; day++) {
      const col = day;
      const cellDate = week === 0
        ? (day === 0 ? `=${firstCell}` : `=$A$4+${day}`)
        : `=$A$4+${week * 7 + day}`;
      // Formula: show day number if in selected month, plus workout indicator
      const displayFml = `=IFERROR(IF(MONTH(${colLetter(col)}${ri+1})<>$D$2,"",DAY(${colLetter(col)}${ri+1})&IF(SUMPRODUCT((${LOG}!$C$8:$C$5000=${colLetter(col)}${ri+1})*(LEN(${LOG}!$B$8:$B$5000)>0)*1)>0,"  ✓","")),"")`;
      // Store date value in a hidden backing area — instead, the cell itself holds the date
      // We use one formula for the date calculation and show day# in the same cell
      // Simpler: store date and display day# + checkmark
      // Use: =IFERROR(IF(MONTH(date_formula)<>$D$2,"", DAY(date_formula)&...), "")
      const dateExpr = week === 0 ? (day === 0 ? firstCell : `$A$4+${day}`) : `$A$4+${week*7+day}`;
      const cellFml = `=IFERROR(IF(MONTH(${dateExpr})<>$D$2,"",DAY(${dateExpr})&IF(SUMPRODUCT((${LOG}!$C$8:$C$5000=${dateExpr})*(LEN(${LOG}!$B$8:$B$5000)>0)*1)>0,"  ✓","")),"")`;
      vals.push({ range: `${S}!${colLetter(col)}${ri+1}`, values: [[cellFml]] });
    }
    // Format calendar row
    const weekBg = week % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,7), cell: { userEnteredFormat: {
      backgroundColor: hex(weekBg), textFormat: { fontSize: 11, fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});
  }

  // Summary stats in rows 4-9, cols H-O (right panel)
  const summaryStats = [
    ['Workouts This Month',  `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=$D$2)*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`],
    ['Total Sets This Month',`=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=$D$2)*(LEN(${LOG}!$B$8:$B$5000)>0)*1)`],
    ['Total Volume (lb)',    `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=$D$2)*IFERROR(${LOG}!$AB$8:$AB$5000*1,0))`],
    ['Avg RPE',              `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=$D$2)*IFERROR(${LOG}!$W$8:$W$5000*1,0))/SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=$D$2)*(LEN(${LOG}!$W$8:$W$5000)>0)*1),"")`],
    ['Monthly Goal',         `=IFERROR('Workout Setup'!B12,"")`],
    ['Goal Status',          `=IFERROR(IF(SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=$D$2)*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))>='Workout Setup'!B12,"On Track","Below Goal"),"")`],
  ];
  summaryStats.forEach(([label, formula], i) => {
    const ri = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,ri,ri+1,7,11), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,ri,ri+1,11,16), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!H${ri+1}`, values: [[label]] });
    vals.push({ range: `${S}!L${ri+1}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,7,11), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,11,16), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });

  // Legend row 10
  vals.push({ range: `${S}!A10`, values: [['✓ = workout logged on that day   |   Blank = rest day or no data']] });
  fmt.push({ mergeCells: { range: gridRange(SID,9,10,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,9,10,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.greenTint), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.accent) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Column widths: A-G calendar cols, H-O summary cols
  const CAL_W = 70;
  for (let ci = 0; ci < 7; ci++) fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: CAL_W }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 7, endIndex: 11 }, properties: { pixelSize: 120 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 11, endIndex: 16 }, properties: { pixelSize: 90 }, fields: 'pixelSize' }});

  // Freeze rows 1:3
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '06-monthly values');
  await batchUpdate(id, fmt, '06-monthly format');
  console.log('✅  Monthly Calendar done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function colLetter(ci) {
  return String.fromCharCode(65 + ci);
}
