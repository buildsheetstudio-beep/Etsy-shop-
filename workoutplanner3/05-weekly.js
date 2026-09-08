'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weekly Planner'];
const S = "'Weekly Planner'";
const LOG = "'Daily Workout Log'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title (merged A:T — no frozenColumnCount, title spans all 20 cols)
  vals.push({ range: `${S}!A1`, values: [['WEEKLY WORKOUT PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Week input row
  vals.push({ range: `${S}!A2`, values: [['Enter any date in your target week:']] });
  vals.push({ range: `${S}!D2`, values: [['2025-10-06']] }); // default: first session week
  vals.push({ range: `${S}!E2`, values: [['← Enter any date (Mon–Sun) to display that week']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,3), mergeType: 'MERGE_ALL' }});
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,4,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,3), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,3,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,4,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Row 3: Column headers
  vals.push({ range: `${S}!A3`, values: [['Day','Date','Workout Name','Type','Status','Sets Logged','Est. Volume (lb)','RPE','Intensity','Location','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Rows 4-10: Mon-Sun (index 3-9)
  const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  DAYS.forEach((day, di) => {
    const r = 4 + di; // spreadsheet row
    const ri = 3 + di; // row index

    vals.push({ range: `${S}!A${r}`, values: [[day]] });

    // B: Date formula (Monday of the week = D2 - (WEEKDAY(D2,2)-1), then +di)
    const dateFml = di === 0
      ? `=IFERROR($D$2-(WEEKDAY($D$2,2)-1),"")`
      : `=IFERROR($B$4+${di},"")`;
    vals.push({ range: `${S}!B${r}`, values: [[dateFml]] });

    // C: Workout Name (first match from Daily Log)
    vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(INDEX(${LOG}!$J$8:$J$5000,MATCH(B${r},${LOG}!$C$8:$C$5000,0)),"")`]] });

    // D: Workout Type
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(INDEX(${LOG}!$I$8:$I$5000,MATCH(B${r},${LOG}!$C$8:$C$5000,0)),"")`]] });

    // E: Status
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(INDEX(${LOG}!$G$8:$G$5000,MATCH(B${r},${LOG}!$C$8:$C$5000,0)),"")`]] });

    // F: Sets Logged
    vals.push({ range: `${S}!F${r}`, values: [[`=SUMPRODUCT((${LOG}!$C$8:$C$5000=B${r})*(LEN(${LOG}!$B$8:$B$5000)>0)*1)`]] });

    // G: Est. Volume
    vals.push({ range: `${S}!G${r}`, values: [[`=SUMPRODUCT((${LOG}!$C$8:$C$5000=B${r})*IFERROR(${LOG}!$AB$8:$AB$5000*1,0))`]] });

    // H: RPE
    vals.push({ range: `${S}!H${r}`, values: [[`=IFERROR(INDEX(${LOG}!$W$8:$W$5000,MATCH(B${r},${LOG}!$C$8:$C$5000,0)),"")`]] });

    // I: Intensity
    vals.push({ range: `${S}!I${r}`, values: [[`=IFERROR(INDEX(${LOG}!$X$8:$X$5000,MATCH(B${r},${LOG}!$C$8:$C$5000,0)),"")`]] });

    // J: Location
    vals.push({ range: `${S}!J${r}`, values: [[`=IFERROR(INDEX(${LOG}!$Y$8:$Y$5000,MATCH(B${r},${LOG}!$C$8:$C$5000,0)),"")`]] });

    // Format row
    const bg = di % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,11), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,1), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat(textFormat)' }});
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,1,2), cell: { userEnteredFormat: {
      numberFormat: { type: 'DATE', pattern: 'ddd, mmm d' },
    }}, fields: 'userEnteredFormat(numberFormat)' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});
  });

  // Row 11: blank spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 12 }, fields: 'pixelSize' }});

  // Row 12: Weekly Summary header
  vals.push({ range: `${S}!A12`, values: [['WEEKLY SUMMARY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,11,12,0,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,11,12,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Rows 13-17: Summary stats
  const WEEK_B4 = `$B$4`;
  const SUMMARIES = [
    ['Workout Days This Week', `=SUMPRODUCT((${LOG}!$C$8:$C$5000>=${WEEK_B4})*(${LOG}!$C$8:$C$5000<=${WEEK_B4}+6)*(${LOG}!$P$8:$P$5000=1)*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`],
    ['Total Sets This Week',   `=SUMPRODUCT((${LOG}!$C$8:$C$5000>=${WEEK_B4})*(${LOG}!$C$8:$C$5000<=${WEEK_B4}+6)*(LEN(${LOG}!$B$8:$B$5000)>0)*1)`],
    ['Total Volume This Week', `=SUMPRODUCT((${LOG}!$C$8:$C$5000>=${WEEK_B4})*(${LOG}!$C$8:$C$5000<=${WEEK_B4}+6)*IFERROR(${LOG}!$AB$8:$AB$5000*1,0))`],
    ['Avg RPE This Week',      `=IFERROR(SUMPRODUCT((${LOG}!$C$8:$C$5000>=${WEEK_B4})*(${LOG}!$C$8:$C$5000<=${WEEK_B4}+6)*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(${LOG}!$W$8:$W$5000*1,0))/SUMPRODUCT((${LOG}!$C$8:$C$5000>=${WEEK_B4})*(${LOG}!$C$8:$C$5000<=${WEEK_B4}+6)*(LEN(${LOG}!$W$8:$W$5000)>0)*1),"")`],
    ['Weekly Goal (Workout Setup)', `=IFERROR('Workout Setup'!B11,"")`],
  ];
  SUMMARIES.forEach(([label, formula], i) => {
    const r = 13 + i;
    vals.push({ range: `${S}!A${r}:B${r}`, values: [[label, formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Column widths
  const WIDTHS = [110,100,180,110,110,80,110,60,100,110,200];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1:3
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '05-weekly values');
  await batchUpdate(id, fmt, '05-weekly format');
  console.log('✅  Weekly Planner done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
