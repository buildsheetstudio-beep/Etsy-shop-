'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Yearly Fitness Planner'];
const S = "'Yearly Fitness Planner'";
const LOG = "'Daily Workout Log'";

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title (no frozenColumnCount — spans all 20 cols)
  vals.push({ range: `${S}!A1`, values: [['YEARLY FITNESS PLANNER & ANALYTICS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Year selector
  vals.push({ range: `${S}!A2`, values: [['Reporting Year:']] });
  vals.push({ range: `${S}!B2`, values: [[2026]] });
  vals.push({ range: `${S}!C2`, values: [['← Change year to update all analytics below']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,2,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,2,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Row 3: Column headers for monthly summary table
  vals.push({ range: `${S}!A3`, values: [['Month','Workout Days','Total Sets','Est. Volume (lb)','Avg RPE','Volume / Day','Strength Days','Cardio Days','HIIT Days','Monthly Goal','Goal Met?','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Rows 4-15: 12 months (January-December)
  // SUMPRODUCT formulas use YEAR() and MONTH() — valid in SUMPRODUCT (only restricted in COUNTIFS)
  MONTHS.forEach((month, mi) => {
    const r = 4 + mi;
    const ri = 3 + mi;
    const mo = mi + 1; // month number 1-12

    vals.push({ range: `${S}!A${r}`, values: [[month]] });

    // B: Workout Days (distinct dates with data)
    vals.push({ range: `${S}!B${r}`, values: [[
      `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`
    ]] });

    // C: Total Sets
    vals.push({ range: `${S}!C${r}`, values: [[
      `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*(LEN(${LOG}!$B$8:$B$5000)>0)*1)`
    ]] });

    // D: Total Volume
    vals.push({ range: `${S}!D${r}`, values: [[
      `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*IFERROR(${LOG}!$AB$8:$AB$5000*1,0))`
    ]] });

    // E: Avg RPE
    vals.push({ range: `${S}!E${r}`, values: [[
      `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*IFERROR(${LOG}!$W$8:$W$5000*1,0))/SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*(LEN(${LOG}!$W$8:$W$5000)>0)*1),"")`
    ]] });

    // F: Avg Volume per Workout Day
    vals.push({ range: `${S}!F${r}`, values: [[`=IFERROR(D${r}/B${r},"")`]] });

    // G: Strength Days
    vals.push({ range: `${S}!G${r}`, values: [[
      `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*(${LOG}!$I$8:$I$5000="Strength")*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`
    ]] });

    // H: Cardio Days
    vals.push({ range: `${S}!H${r}`, values: [[
      `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*(${LOG}!$I$8:$I$5000="Cardio")*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`
    ]] });

    // I: HIIT Days
    vals.push({ range: `${S}!I${r}`, values: [[
      `=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=$B$2)*(MONTH(${LOG}!$C$8:$C$5000)=${mo})*(${LOG}!$I$8:$I$5000="HIIT")*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`
    ]] });

    // J: Monthly Goal
    vals.push({ range: `${S}!J${r}`, values: [[`=IFERROR('Workout Setup'!B12,"")`]] });

    // K: Goal Met?
    vals.push({ range: `${S}!K${r}`, values: [[`=IFERROR(IF(B${r}>=J${r},"Yes","No"),"")`]] });

    // Format row
    const bg = mi % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,12), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,1), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat(textFormat)' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 16: Totals row
  vals.push({ range: `${S}!A16`, values: [['YEAR TOTAL']] });
  vals.push({ range: `${S}!B16`, values: [[`=SUM(B4:B15)`]] });
  vals.push({ range: `${S}!C16`, values: [[`=SUM(C4:C15)`]] });
  vals.push({ range: `${S}!D16`, values: [[`=SUM(D4:D15)`]] });
  vals.push({ range: `${S}!E16`, values: [[`=IFERROR(AVERAGE(IFERROR(E4:E15*1,)),"")`]] });
  fmt.push({ repeatCell: { range: gridRange(SID,15,16,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Row 17: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 16 }, fields: 'pixelSize' }});

  // Row 18: Annual Goals section header
  vals.push({ range: `${S}!A18`, values: [['ANNUAL FITNESS GOALS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,17,18,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,17,18,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Rows 19-23: Annual goal inputs
  const GOALS = [
    ['Annual Workout Sessions Target', `=IFERROR('Workout Setup'!B12*12,"")`, ''],
    ['Annual Workout Sessions (Actual)', `=IFERROR(SUM(B4:B15),"")`, ''],
    ['Annual Volume Target (lb)',       '', ''],
    ['Annual Volume (Actual)',          `=IFERROR(SUM(D4:D15),"")`, ''],
    ['Year-End Goal',                   `=IFERROR('Workout Setup'!B13,"")`, ''],
  ];
  GOALS.forEach(([label, formula, notes], i) => {
    const r = 19 + i;
    vals.push({ range: `${S}!A${r}:C${r}`, values: [[label, formula, notes]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Column widths
  const WIDTHS = [120,90,80,120,70,110,90,90,70,90,80,160];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1:3
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '07-yearly values');
  await batchUpdate(id, fmt, '07-yearly format');
  console.log('✅  Yearly Fitness Planner done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
