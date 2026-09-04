'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BB = sheetMap['💰 Baby Budget'];

// 14 cols, 65 rows
// Row 0: Title, Row 1: Summary labels, Row 2: Summary values
// Row 3: Allocation section header, Row 4: Alloc col headers, Row 5-16: 12 alloc categories
// Row 17: Expenses section header, Row 18: Expense col headers, Row 19-64: expense entries

(async () => {
  const reqs = [];

  // Title
  reqs.push({ mergeCells: { range: gridRange(BB, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BB, 0, 1, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BB, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Summary rows 1-2
  reqs.push({ repeatCell: {
    range: gridRange(BB, 1, 2, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(BB, 2, 3, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BB, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Section header: Allocation row 3
  reqs.push({ mergeCells: { range: gridRange(BB, 3, 4, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BB, 3, 4, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BB, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Alloc col headers row 4
  reqs.push({ repeatCell: {
    range: gridRange(BB, 4, 5, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BB, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Alloc data rows 5-16
  for (let r = 5; r < 17; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(BB, r, r+1, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Section header: Expenses row 17
  reqs.push({ mergeCells: { range: gridRange(BB, 17, 18, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BB, 17, 18, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BB, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Expense col headers row 18
  reqs.push({ repeatCell: {
    range: gridRange(BB, 18, 19, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BB, dimension: 'ROWS', startIndex: 18, endIndex: 19 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Expense alt rows 19-64
  for (let r = 19; r < 65; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(BB, r, r+1, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Column widths
  const colW = [30, 180, 130, 90, 100, 100, 90, 80, 120, 90, 100, 100, 120, 160];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: BB, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Number format: currency columns (D,E,F in alloc = cols 3,4,5; E,F in expense = 4,5)
  const currFmt = { type: 'CURRENCY', pattern: '"$"#,##0.00' };
  [3,4,5,6].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(BB, 5, 17, c, c+1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' }});
    reqs.push({ repeatCell: { range: gridRange(BB, 19, 65, c, c+1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Summary row currency
  [1,2,3,4,5,6].forEach(c => reqs.push({ repeatCell: { range: gridRange(BB, 2, 3, c, c+1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' }}));

  // Wrap + vertical align
  reqs.push({ repeatCell: { range: gridRange(BB, 5, 65, 0, 14), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(BB, 4, 17, 0, 14), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});
  reqs.push({ updateBorders: { range: gridRange(BB, 18, 65, 0, 14), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});

  await batchUpdate(id, reqs, 'bb-format');

  const vdata = [];
  vdata.push({ range: `'💰 Baby Budget'!A1`, values: [['💰 Baby Budget']] });

  vdata.push({ range: `'💰 Baby Budget'!A2:N2`, values: [[
    'Total Budget','Total Allocated','Total Spent','Remaining Budget','% Spent','Savings Goal','Gift Registry Credit','Outstanding','','','','','','',
  ]]});
  vdata.push({ range: `'💰 Baby Budget'!A3:N3`, values: [[
    5000,
    '=SUM(D6:D17)',
    '=SUM(F20:F65)',
    '=A3-C3',
    '=IFERROR(TEXT(C3/A3,"0%"),"—")',
    500,
    '=SUMIF(J20:J65,TRUE,F20:F65)',
    '=C3-G3',
    '','','','','','',
  ]]});

  vdata.push({ range: `'💰 Baby Budget'!A4`, values: [['  📦 Budget Allocation by Category']] });

  vdata.push({ range: `'💰 Baby Budget'!A5:N5`, values: [[
    '#','Category','Priority','Allocated','Actual Spent','Variance','% Used','Paid','Payment Status','Store / Vendor','Due Date','Notes','','',
  ]]});

  const allocRows = [
    [1,'Nursery & Furniture','High',800,'=SUMIF($B$20:$B$65,B6,$F$20:$F$65)','=D6-E6','=IFERROR(TEXT(E6/D6,"0%"),"—")',false,'Pending','Buy Buy Baby','',  'Crib, dresser, glider'],
    [2,'Stroller & Car Seat','High',600,'=SUMIF($B$20:$B$65,B7,$F$20:$F$65)','=D7-E7','=IFERROR(TEXT(E7/D7,"0%"),"—")',true,'Paid in Full','UPPAbaby','',  'Travel system purchased'],
    [3,'Clothing & Gear','Medium',400,'=SUMIF($B$20:$B$65,B8,$F$20:$F$65)','=D8-E8','=IFERROR(TEXT(E8/D8,"0%"),"—")',false,'Pending','Various','',  '0-3M, 3-6M, 6-12M'],
    [4,'Medical & Healthcare','High',500,'=SUMIF($B$20:$B$65,B9,$F$20:$F$65)','=D9-E9','=IFERROR(TEXT(E9/D9,"0%"),"—")',true,'Paid in Full','Insurance / OB','',  'Copays, prenatal vitamins'],
    [5,'Feeding Supplies','High',300,'=SUMIF($B$20:$B$65,B10,$F$20:$F$65)','=D10-E10','=IFERROR(TEXT(E10/D10,"0%"),"—")',false,'Pending','Amazon / Target','',  'Pump, bottles, nursing pads'],
    [6,'Postpartum Care','Medium',250,'=SUMIF($B$20:$B$65,B11,$F$20:$F$65)','=D11-E11','=IFERROR(TEXT(E11/D11,"0%"),"—")',false,'Pending','Target','',  'Recovery essentials'],
    [7,'Education & Books','Low',100,'=SUMIF($B$20:$B$65,B12,$F$20:$F$65)','=D12-E12','=IFERROR(TEXT(E12/D12,"0%"),"—")',true,'Paid in Full','Amazon','',  'Baby books, parenting books'],
    [8,'Entertainment & Toys','Low',200,'=SUMIF($B$20:$B$65,B13,$F$20:$F$65)','=D13-E13','=IFERROR(TEXT(E13/D13,"0%"),"—")',false,'Pending','Target','',  'Play gym, mobiles'],
    [9,'Shower & Celebrations','Medium',300,'=SUMIF($B$20:$B$65,B14,$F$20:$F$65)','=D14-E14','=IFERROR(TEXT(E14/D14,"0%"),"—")',true,'Paid in Full','Restaurant / Décor','',  'Baby shower budget'],
    [10,'Travel & Transport','Low',150,'=SUMIF($B$20:$B$65,B15,$F$20:$F$65)','=D15-E15','=IFERROR(TEXT(E15/D15,"0%"),"—")',false,'Pending','Various','',  'Hospital bag transport, visits'],
    [11,'Tech & Gadgets','Medium',200,'=SUMIF($B$20:$B$65,B16,$F$20:$F$65)','=D16-E16','=IFERROR(TEXT(E16/D16,"0%"),"—")',false,'Pending','Amazon','',  'Baby monitor, white noise machine'],
    [12,'Miscellaneous','Low',200,'=SUMIF($B$20:$B$65,B17,$F$20:$F$65)','=D17-E17','=IFERROR(TEXT(E17/D17,"0%"),"—")',false,'Pending','Various','',  'Overflow / unexpected costs'],
  ];
  vdata.push({ range: `'💰 Baby Budget'!A6:M17`, values: allocRows });

  vdata.push({ range: `'💰 Baby Budget'!A18`, values: [['  🛒 Expense Log']] });

  vdata.push({ range: `'💰 Baby Budget'!A19:N19`, values: [[
    '#','Category','Item / Description','Date','Amount','Actual Cost','Notes','Store / Vendor','Priority','Gift / Registry Item','Payment Method','Status','Receipt / Link','Tags',
  ]]});

  const expenses = [
    [1,'Stroller & Car Seat','UPPAbaby Vista V2 Travel System','2025-07-15',1200,1150,'Found on sale!','Buy Buy Baby','High',false,'Credit Card','Paid in Full','','Stroller'],
    [2,'Medical & Healthcare','Prenatal vitamins (3-month supply)','2025-06-01',45,43.99,'Ritual brand','Amazon','High',false,'Debit','Paid in Full','','Vitamins'],
    [3,'Medical & Healthcare','OB co-pays (Q1)','2025-06-30',90,90,'Insurance co-pays','City OB Clinic','High',false,'Debit','Paid in Full','','Medical'],
    [4,'Nursery & Furniture','Babyletto Hudson Crib','2025-08-20',400,380,'White finish','Buy Buy Baby','High',false,'Credit Card','Paid in Full','','Crib'],
    [5,'Nursery & Furniture','Glider chair (Pottery Barn)','2025-08-20',550,550,'Cream bouclé','Pottery Barn','High',false,'Credit Card','Paid in Full','','Furniture'],
    [6,'Clothing & Gear','Newborn & 0-3M clothing bundle','2025-09-01',150,145,'Mix of brands','Target + Amazon','Medium',false,'Debit','Paid in Full','','Clothing'],
    [7,'Education & Books','What to Expect When Expecting','2025-06-05',18,17.99,'','Amazon','Low',false,'Debit','Paid in Full','','Books'],
    [8,'Education & Books','Ina May\'s Guide to Childbirth','2025-07-01',16,15.99,'','Amazon','Low',false,'Debit','Paid in Full','','Books'],
    [9,'Shower & Celebrations','Baby shower venue & décor','2025-10-01',250,250,'Hosted by bestie','Events by Lily','Medium',false,'Cash','Paid in Full','','Shower'],
    [10,'Feeding Supplies','Spectra S2 breast pump','2025-09-15',180,180,'Covered partially by insurance','Amazon','High',false,'Insurance','Paid in Full','','Pump'],
    [11,'Nursery & Furniture','Dresser / changing table combo','2025-08-25',320,299,'IKEA Hemnes','IKEA','High',false,'Credit Card','Paid in Full','','Furniture'],
    [12,'Tech & Gadgets','Nanit Pro baby monitor','2025-10-15',300,279,'With breathing band','Amazon','Medium',false,'Credit Card','Deposit Paid','','Tech'],
  ];
  vdata.push({ range: `'💰 Baby Budget'!A20:N31`, values: expenses });

  await valuesBatchUpdate(id, vdata, 'bb-values');
  console.log('Baby Budget complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
