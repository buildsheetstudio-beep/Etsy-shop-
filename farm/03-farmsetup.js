'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS, ENTERPRISES, INCOME_CATS, EXPENSE_CATS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const FS = sheetMap['Farm Setup'];
const S = "'Farm Setup'";

// Budget values for income categories (used by Budget vs Actual)
const INCOME_BUDGETS = [20000,25000,5000,1500,8000,10000,6000,0,0,0,0,0,0,0,2000,0,0];
// Budget values for expense categories
const EXPENSE_BUDGETS = [3500,8000,2500,500,1000,1500,8000,0,2000,3000,500,1000,800,1000,8000,2000,2500,1800,2500,0,500,600,500,600,1000,300,500,300,800,500];
// Monthly profit targets (seasonal, sums to $65,000)
const MONTHLY_GOALS  = [1500,1500,2500,4000,6000,7000,8000,9000,6000,8000,5500,6000];
const MONTHLY_NOTES  = ['Winter slow season','Winter slow season','Spring prep begins','Spring crops going in','Strong spring income','Peak early-summer produce','Peak summer stand','Corn & harvest income','Fall fieldwork','Soybean & fall harvest','Post-harvest slowdown','Year-end wrap-up'];

(async () => {
  const data = [];

  // ── Row 1-2: Title ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['FARM SETUP & CONFIGURATION — Cedar Ridge Farm']] });
  data.push({ range: `${S}!A2`, values: [['Configure your farm details, enterprise list, budgets, and monthly profit targets here.']] });

  // ── Row 4-13: Farm Information ─────────────────────────────────────────────
  data.push({ range: `${S}!A4`, values: [['FARM INFORMATION']] });
  data.push({ range: `${S}!A5:B12`, values: [
    ['Farm Name',           'Cedar Ridge Farm'],
    ['Owner / Manager',     'Taylor Morgan'],
    ['Reporting Year',      2026],
    ['Currency Symbol',     '$'],
    ['Farm Type',           'Mixed Farm'],
    ['Primary Enterprise',  'Crops'],
    ['Starting Cash Balance', 15000],
    ['Annual Profit Goal',    65000],
  ]});

  // ── Row 14-31: Enterprise Customization ───────────────────────────────────
  data.push({ range: `${S}!A14`, values: [['ENTERPRISE CUSTOMIZATION']] });
  data.push({ range: `${S}!A15:C15`, values: [['Enterprise Slot','Enterprise Name','Active?']] });
  const entData = ENTERPRISES.map((e, i) => [i+1, e, true]);
  data.push({ range: `${S}!A16:C30`, values: entData });

  // ── Row 32-51: Annual Income Budgets ──────────────────────────────────────
  data.push({ range: `${S}!A32`, values: [['ANNUAL BUDGET INPUTS']] });
  data.push({ range: `${S}!A33:D33`, values: [['Type','Category','Annual Budget','Notes']] });
  // Income rows 34-50
  data.push({ range: `${S}!A34`, values: [['Income']] });
  INCOME_CATS.forEach((cat, i) => {
    data.push({ range: `${S}!B${34+i}:C${34+i}`, values: [[cat, INCOME_BUDGETS[i]]] });
  });
  // Expense rows 52-81
  data.push({ range: `${S}!A52`, values: [['Expense']] });
  EXPENSE_CATS.forEach((cat, i) => {
    data.push({ range: `${S}!B${52+i}:C${52+i}`, values: [[cat, EXPENSE_BUDGETS[i]]] });
  });

  // ── Row 83-97: Monthly Profit Targets ─────────────────────────────────────
  data.push({ range: `${S}!A83`, values: [['MONTHLY PROFIT TARGETS']] });
  data.push({ range: `${S}!A84:C84`, values: [['Month','Profit Goal','Notes']] });
  MONTHS.forEach((m, i) => {
    data.push({ range: `${S}!A${85+i}:C${85+i}`, values: [[m, MONTHLY_GOALS[i], MONTHLY_NOTES[i]]] });
  });

  // ── Row 98: Disclaimer ────────────────────────────────────────────────────
  data.push({ range: `${S}!A98`, values: [['⚠ DISCLAIMER']] });
  data.push({ range: `${S}!A99`, values: [['This spreadsheet is an organizational and educational tool only. It does not provide accounting, tax, legal, or financial advice and does not replace records required by a qualified professional or government authority. Tax treatment varies by location and individual circumstances. Verify all classifications, deductions, and reporting requirements with a qualified tax or accounting professional.']] });

  await valuesBatchUpdate(id, data, 'farmsetup-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];
  const merge = (r1,r2,c1,c2) => reqs.push({ mergeCells: { range: gridRange(FS,r1,r2,c1,c2), mergeType:'MERGE_ALL' } });
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell: { range: gridRange(FS,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });

  // Title merges
  merge(0,1,0,6); merge(1,2,0,6);

  // Title format
  fmt(0,1,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:16, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,2,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Section header helper
  const secHdr = (r) => {
    merge(r-1,r,0,5);
    fmt(r-1,r,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  };
  secHdr(4); secHdr(14); secHdr(32); secHdr(83); secHdr(98);

  // Col header rows
  const colHdr = (r1,r2,c1,c2) => fmt(r1,r2,c1,c2,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER' } });
  colHdr(14,15,0,3); // enterprise header
  colHdr(32,33,0,4); // budget header
  colHdr(83,84,0,3); // monthly goals header

  // Farm info label col A
  fmt(4,13,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Farm info value col B (input)
  fmt(4,13,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ foregroundColor:hex(C.mainText), fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Currency for B11:B12 (rows 10-11, 0-indexed)
  reqs.push({ repeatCell:{ range:gridRange(FS,10,12,1,2), cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' } } }, fields:'userEnteredFormat.numberFormat' } });

  // Enterprise data rows
  for (let r = 15; r < 30; r++) {
    const bg = r%2===0 ? C.bg : C.altRow;
    fmt(r,r+1,0,3,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
    // Enterprise name = input
    reqs.push({ repeatCell:{ range:gridRange(FS,r,r+1,1,2), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }

  // Budget table rows
  fmt(33,82,0,4,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  // Alternate rows in budget
  for (let r = 33; r < 82; r++) {
    if (r%2===1) fmt(r,r+1,0,4,{ userEnteredFormat:{ backgroundColor:hex(C.altRow) } });
  }
  // Budget Amount = input, currency
  fmt(33,82,2,3,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial' }, numberFormat:{ type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' }, horizontalAlignment:'RIGHT' } });

  // Monthly goals rows
  for (let r = 84; r < 96; r++) {
    const bg = r%2===0 ? C.bg : C.altRow;
    fmt(r,r+1,0,3,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(84,96,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), numberFormat:{ type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' }, horizontalAlignment:'RIGHT' } });

  // Disclaimer
  merge(97,98,0,5);
  fmt(97,98,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  merge(98,104,0,5);
  fmt(98,104,0,5,{ userEnteredFormat:{ backgroundColor:hex('#FFF5F0'), textFormat:{ foregroundColor:hex('#5A2A20'), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'TOP', wrapStrategy:'WRAP' } });

  // Borders
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID',        color:hex(C.border) };
  [[4,13,0,5],[14,30,0,3],[32,82,0,4],[83,96,0,3]].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ updateBorders: { range:gridRange(FS,r1,r2,c1,c2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });
  });

  // Col widths: A=160, B=200, C=130, D=180, E=100
  [[0,160],[1,200],[2,130],[3,180],[4,100]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:FS, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:FS, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:38 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:FS, dimension:'ROWS', startIndex:1, endIndex:200 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  // Section headers taller
  for (const r of [3,13,31,82,97]) {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:FS, dimension:'ROWS', startIndex:r, endIndex:r+1 }, properties:{ pixelSize:28 }, fields:'pixelSize' } });
  }

  // Freeze rows 1-3
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:FS, gridProperties:{ frozenRowCount:3, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'farmsetup-format');
  console.log('Farm Setup complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
