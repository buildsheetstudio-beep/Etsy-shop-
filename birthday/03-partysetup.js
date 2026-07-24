'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PS = sheetMap['Party Setup'];
const S = "'Party Setup'";

(async () => {
  const data = [];

  // Row 1 = title (index 0), Row 2 = title (index 1), Row 3 = guidance (index 2)
  // Freeze rows 1-3 (frozenRowCount:3, indices 0-2)
  // Section headers start at row 4 (index 3)
  // Data starts at row 5 (index 4)

  data.push({ range:`${S}!A1`, values:[['KIDS BIRTHDAY PARTY SETUP']] });
  data.push({ range:`${S}!A3`, values:[['Enter your party details here. All other tabs and the dashboard will update automatically based on this information.']] });

  // Section header labels (row 4, index 3)
  data.push({ range:`${S}!A4`, values:[['PARTY DETAILS']] });
  data.push({ range:`${S}!D4`, values:[['PLANNING SUMMARY']] });

  // Party detail fields — labels in A, values in B, rows 5-20
  const fields = [
    ['Birthday Child:',         'Mia'],
    ['Age Turning:',            8],
    ['Party Date:',             'August 15, 2026'],
    ['Start Time:',             '2:00 PM'],
    ['End Time:',               '5:00 PM'],
    ['Party Theme:',            'Enchanted Garden'],
    ['Party Type:',             'Venue Party'],
    ['Venue / Location:',       'Riverside Community Room'],
    ['Address:',                '45 Riverside Drive, Mapleton'],
    ['Host / Parent Name:',     'Sarah Hendricks'],
    ['Contact Number:',         '(555) 204-8871'],
    ['Total Budget:',           1200],
    ['Guest Capacity:',         30],
    ['RSVP Deadline:',          'August 1, 2026'],
    ['Weather Backup Needed?:', 'No'],
    ['Notes:',                  'Book decorating help 1 week before; confirm headcount with venue by Aug 5.'],
  ];

  fields.forEach(([label, value], i) => {
    const r = 5 + i;
    data.push({ range:`${S}!A${r}`, values:[[label]] });
    data.push({ range:`${S}!B${r}`, values:[[value]] });
  });

  // Summary cards — label in D, value formula in F
  // Cards at rows 5,8,11,14,17,20,23,26 (each spans 2 rows)
  const cards = [
    ['Days Until Party',    `=IFERROR(MAX(0,DATEVALUE("August 15, 2026")-TODAY()),"—")`],
    ['Total Budget',        `=IFERROR(B16,0)`],
    ['Guest Capacity',      `=IFERROR(B17,0)`],
    ['Confirmed Guests',    `=IFERROR(COUNTIF('Guest List & RSVP'!$I$6:$I$205,"Confirmed"),0)`],
    ['Total Attending',     `=IFERROR(SUMIF('Guest List & RSVP'!$I$6:$I$205,"Confirmed",'Guest List & RSVP'!$J$6:$J$205),0)`],
    ['RSVP Deadline',       `="August 1, 2026"`],
    ['Budget Remaining',    `=IFERROR(B16-SUMIFS('Budget & Expenses'!$I$31:$I$229,'Budget & Expenses'!$I$31:$I$229,">"&0),B16)`],
    ['Est. Cost / Guest',   `=IFERROR(B16/MAX(1,SUMIF('Guest List & RSVP'!$I$6:$I$205,"Confirmed",'Guest List & RSVP'!$J$6:$J$205)),0)`],
  ];

  cards.forEach(([label, formula], i) => {
    const row = 5 + (i * 3);
    data.push({ range:`${S}!D${row}`, values:[[label]] });
    data.push({ range:`${S}!F${row}`, values:[[formula]] });
  });

  await valuesBatchUpdate(id, data, 'partysetup-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(PS,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const currF  = { type:'CURRENCY', pattern:'$#,##0.00' };
  const timeF  = { type:'TIME', pattern:'h:mm AM/PM' };

  // Title merge A1:H2 (rows 0-1)
  reqs.push({ mergeCells:{ range:gridRange(PS,0,2,0,8), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:22, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Guidance row 3 (index 2) — single row merge
  reqs.push({ mergeCells:{ range:gridRange(PS,2,3,0,8), mergeType:'MERGE_ALL' } });
  fmt(2,3,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Section header: PARTY DETAILS row 4 (index 3), cols A:B
  reqs.push({ mergeCells:{ range:gridRange(PS,3,4,0,2), mergeType:'MERGE_ALL' } });
  fmt(3,4,0,2,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Fields A5:A20 labels, B5:B20 input (indices 4-19)
  fmt(4,20,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, verticalAlignment:'MIDDLE' } });
  fmt(4,20,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });

  // Date, time, currency formats for specific cells
  reqs.push({ repeatCell:{ range:gridRange(PS,6,7,1,2), cell:{ userEnteredFormat:{ numberFormat:dateF } }, fields:'userEnteredFormat.numberFormat' } }); // Party Date B7
  reqs.push({ repeatCell:{ range:gridRange(PS,7,9,1,2), cell:{ userEnteredFormat:{ numberFormat:timeF } }, fields:'userEnteredFormat.numberFormat' } }); // B8,B9
  reqs.push({ repeatCell:{ range:gridRange(PS,15,16,1,2), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } }); // Budget B16
  reqs.push({ repeatCell:{ range:gridRange(PS,17,18,1,2), cell:{ userEnteredFormat:{ numberFormat:dateF } }, fields:'userEnteredFormat.numberFormat' } }); // RSVP Deadline B18

  reqs.push({ updateBorders:{ range:gridRange(PS,3,20,0,2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Section header: PLANNING SUMMARY row 4 (index 3), cols D:F
  reqs.push({ mergeCells:{ range:gridRange(PS,3,4,3,6), mergeType:'MERGE_ALL' } });
  fmt(3,4,3,6,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Cards: each card spans 2 data rows, label in D (col 3), value in F (col 5)
  // Cards at rows 5,8,11,14,17,20,23,26 → indices 4,7,10,13,16,19,22,25
  const cardIndices = [4,7,10,13,16,19,22,25];
  cardIndices.forEach(ri => {
    reqs.push({ mergeCells:{ range:gridRange(PS,ri,ri+2,3,5), mergeType:'MERGE_ALL' } }); // label D:E
    reqs.push({ mergeCells:{ range:gridRange(PS,ri,ri+2,5,6), mergeType:'MERGE_ALL' } }); // value F
    fmt(ri,ri+2,3,5,{ userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
    fmt(ri,ri+2,5,6,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.primary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  });

  // Currency for Budget, Budget Remaining, Cost per Guest cards (indices 7,22,25)
  reqs.push({ repeatCell:{ range:gridRange(PS,7,9,5,6), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(PS,22,24,5,6), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(PS,25,27,5,6), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });

  reqs.push({ updateBorders:{ range:gridRange(PS,3,27,3,6), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Column widths
  [[0,175],[1,200],[2,20],[3,160],[4,20],[5,120]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:PS, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PS, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:44 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PS, dimension:'ROWS', startIndex:2, endIndex:4 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PS, dimension:'ROWS', startIndex:4, endIndex:28 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Freeze rows 1-3 = frozenRowCount:3
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:PS, gridProperties:{ frozenRowCount:3 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'partysetup-format');
  console.log('Party Setup complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
