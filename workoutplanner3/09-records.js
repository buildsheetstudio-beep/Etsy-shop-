'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Personal Records'];
const S = "'Personal Records'";
const LOG = "'Daily Workout Log'";
const SETUP = "'Workout Setup'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,1100,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title (no frozenColumnCount — spans all 16 cols)
  vals.push({ range: `${S}!A1`, values: [['PERSONAL RECORDS & EXERCISE BESTS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Auto-updated from Daily Workout Log. Best values are calculated automatically when you log new workouts.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Row 3: Column headers
  vals.push({ range: `${S}!A3`, values: [['Exercise ID','Exercise Name','Category','Primary Muscle','Tracking Mode','Total Sets','Best Weight (lb)','Best Reps','Best Duration (min)','Best Distance','Best PR Value','Date of Best PR','PR Session ID','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Rows 4-86: 83 exercises (index 3-85)
  // Column A: Exercise ID — formula pulls from Workout Setup row 17 onwards
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,0,1), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(${SETUP}!A17,"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column B: Exercise Name
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,1,2), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(A4,${SETUP}!$A$17:$N$1016,2,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column C: Category
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,2,3), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(A4,${SETUP}!$A$17:$N$1016,3,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column D: Primary Muscle
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,3,4), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(A4,${SETUP}!$A$17:$N$1016,4,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column E: Tracking Mode
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,4,5), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(A4,${SETUP}!$A$17:$N$1016,6,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column F: Total Sets Logged
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,5,6), cell: {
    userEnteredValue: { formulaValue: `=SUMPRODUCT((${LOG}!$K$8:$K$5000=A4)*1)` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column G: Best Weight
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,6,7), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(MAXIFS(${LOG}!$Q$8:$Q$5000,${LOG}!$K$8:$K$5000,A4),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column H: Best Reps
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,7,8), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(MAXIFS(${LOG}!$R$8:$R$5000,${LOG}!$K$8:$K$5000,A4),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column I: Best Duration
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,8,9), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(MAXIFS(${LOG}!$S$8:$S$5000,${LOG}!$K$8:$K$5000,A4),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column J: Best Distance
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,9,10), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(MAXIFS(${LOG}!$T$8:$T$5000,${LOG}!$K$8:$K$5000,A4),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column K: Best PR Value (from AC column in Daily Log)
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,10,11), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(MAXIFS(${LOG}!$AC$8:$AC$5000,${LOG}!$K$8:$K$5000,A4),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.success) }, horizontalAlignment: 'CENTER',
      numberFormat: { type: 'NUMBER', pattern: '#,##0.##' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column L: Date of Best PR (most recent date where PR value was achieved)
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,11,12), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(TEXT(MAXIFS(${LOG}!$C$8:$C$5000,${LOG}!$K$8:$K$5000,A4,${LOG}!$AC$8:$AC$5000,K4),"yyyy-mm-dd"),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Column M: PR Session ID
  fmt.push({ repeatCell: { range: gridRange(SID,3,86,12,13), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(INDEX(${LOG}!$B$8:$B$5000,MATCH(MAXIFS(${LOG}!$C$8:$C$5000,${LOG}!$K$8:$K$5000,A4,${LOG}!$AC$8:$AC$5000,K4),${LOG}!$C$8:$C$5000,0)),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial', fontSize: 8 } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Alternating row backgrounds for data area
  for (let i = 0; i < 83; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,3+i,4+i,13,14), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3+i, endIndex: 4+i }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  }

  // Column widths
  const WIDTHS = [80,180,100,120,120,70,90,70,90,80,90,100,150,180];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1:3 (no column freeze — title spans all 16 cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '09-records values');
  await batchUpdate(id, fmt, '09-records format');
  console.log('✅  Personal Records done — 83 exercise PR rows.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
