'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// All four layouts pull from Invoice Log (row matching by Invoice ID from Business Setup B30)
// and from the line-item archive block in Invoice Log columns AC onward (see archive note)
// For simplicity: pull summary data from Invoice Log columns A-R; line items reference
// a fixed invoice row via INDEX/MATCH on the Invoice ID selector.

const CLASSIC  = sheetMap['Invoice — Classic'];
const MODERN   = sheetMap['Invoice — Modern'];
const ELEGANT  = sheetMap['Invoice — Elegant'];
const COMPACT  = sheetMap['Invoice — Compact'];

const BS  = "'Business Setup'";
const IL  = "'Invoice Log'";
const PT  = "'Payment Tracker'";

// Shared selector formula base — matches selected Invoice ID in Business Setup B30
function ilLookup(col) {
  // col is 1-based column index in Invoice Log (A=1, B=2, ...)
  return `=IFERROR(INDEX('Invoice Log'!$${colLetter(col-1)}$6:$${colLetter(col-1)}$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)),"")`;
}
function bsField(row) { return `=${BS}!B${row}`; }

// Build one layout on a given sheetId
async function buildLayout(sheetId, sheetName, style) {
  const S = `'${sheetName}'`;
  const data = [];
  const reqs = [];

  // Palette per style
  const palette = {
    'Classic': { header: '#2F2C30', accent: '#5A3A52', line: '#CBC5BF', headerText: '#FFFFFF', bg: '#FFFFFF', altRow: '#F5F3F0' },
    'Modern':  { header: '#5A3A52', accent: '#78998B', line: '#78998B', headerText: '#FFFFFF', bg: '#FFFFFF', altRow: '#F7F4EF' },
    'Elegant': { header: '#5A3A52', accent: '#A9A6A2', line: '#CBC5BF', headerText: '#FFFFFF', bg: '#FFFFFF', altRow: '#FBFAF8' },
    'Compact': { header: '#2F2C30', accent: '#5A3A52', line: '#CBC5BF', headerText: '#FFFFFF', bg: '#FFFFFF', altRow: '#F5F3F0' },
  };
  const p = palette[style];

  // ── Row layout (0-indexed) ─────────────────────────────────────────────────
  // Row 0: Edit note
  // Row 1-2: Business name / initials
  // Row 3: Business contact line
  // Row 4: INVOICE label
  // Row 5: Invoice # / Date / Due Date
  // Row 6: Bill To label
  // Row 7-11: Client info
  // Row 12: PO / Project
  // Row 13: Line items header
  // Rows 14-33: 20 line item rows
  // Row 34: Spacer
  // Row 35-42: Totals (subtotal, discount, shipping, tax, grand total, payments, balance)
  // Row 43: Payment instructions
  // Row 44-45: Notes + Terms
  // Row 46: Thank-you

  // Edit note
  reqs.push({ mergeCells: { range: gridRange(sheetId, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 0, 1, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex('#FFF9C4'),
    textFormat: { italic: true, foregroundColor: hex('#5D4037'), fontSize: 8 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  data.push({ range: `${S}!A1`, values: [['This layout pulls from the invoice selected in Business Setup B30. All data is read-only here. Edit source data in Invoice Log or Payment Tracker.']] });

  // Business initials / name block
  reqs.push({ mergeCells: { range: gridRange(sheetId, 1, 2, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 1, 2, 0, 2), cell: { userEnteredFormat: {
    backgroundColor: hex(p.header), textFormat: { bold: true, foregroundColor: hex(p.headerText), fontSize: style === 'Elegant' ? 22 : 20 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  data.push({ range: `${S}!A2`, values: [[`=${BS}!B21`]] }); // Initials

  reqs.push({ mergeCells: { range: gridRange(sheetId, 1, 2, 2, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 1, 2, 2, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.header), textFormat: { bold: true, foregroundColor: hex(p.headerText), fontSize: 15 },
    horizontalAlignment: style === 'Modern' ? 'LEFT' : 'RIGHT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  data.push({ range: `${S}!C2`, values: [[`=${BS}!B5`]] }); // Business Name

  // Contact line
  reqs.push({ mergeCells: { range: gridRange(sheetId, 2, 3, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 2, 3, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.accent), textFormat: { foregroundColor: hex('#FFFFFF'), fontSize: 8 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  data.push({ range: `${S}!A3`, values: [[
    `=IFERROR(TEXTJOIN("  |  ",TRUE,${BS}!B8,${BS}!B9,${BS}!B10),"")` ]] });
  reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // INVOICE label
  reqs.push({ mergeCells: { range: gridRange(sheetId, 3, 4, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 3, 4, 0, 4), cell: { userEnteredFormat: {
    backgroundColor: hex(p.bg), textFormat: { bold: true, foregroundColor: hex(p.header), fontSize: 16 },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  data.push({ range: `${S}!A4`, values: [['INVOICE']] });
  reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Invoice meta (# / Date / Due)
  data.push({ range: `${S}!F4`, values: [['Invoice #']] });
  data.push({ range: `${S}!G4`, values: [[ilLookup(2)]] }); // Invoice Number
  data.push({ range: `${S}!F5`, values: [['Invoice Date']] });
  data.push({ range: `${S}!G5`, values: [[ilLookup(5)]] }); // Invoice Date
  data.push({ range: `${S}!F6`, values: [['Due Date']] });
  data.push({ range: `${S}!G6`, values: [[ilLookup(6)]] }); // Due Date
  data.push({ range: `${S}!F7`, values: [['Currency']] });
  data.push({ range: `${S}!G7`, values: [[ilLookup(10)]] }); // Currency
  reqs.push({ repeatCell: { range: gridRange(sheetId, 3, 8, 5, 6), cell: { userEnteredFormat: {
    textFormat: { bold: true, foregroundColor: hex(p.header) }, horizontalAlignment: 'RIGHT',
  }}, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 3, 8, 6, 9), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.mainText) } } }, fields: 'userEnteredFormat.textFormat' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 3, 4, 5, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 3, 4, 7, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 4, 5, 5, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 4, 5, 7, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 5, 6, 5, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 5, 6, 7, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 6, 7, 5, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 6, 7, 7, 9), mergeType: 'MERGE_ALL' } });

  // Bill To block
  reqs.push({ mergeCells: { range: gridRange(sheetId, 4, 5, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 4, 5, 0, 4), cell: { userEnteredFormat: {
    textFormat: { bold: true, foregroundColor: hex(p.accent), fontSize: 9 },
  }}, fields: 'userEnteredFormat.textFormat' } });
  data.push({ range: `${S}!A5`, values: [['BILL TO']] });

  // Client info lines
  const clientLines = [
    [ilLookup(4)],  // Client Name
    [ilLookup(4)],  // placeholder - use address from CRM via generator
    [`=IFERROR(INDEX('Client CRM'!$F$6:$F$305,MATCH(${ilLookup(3).replace('=','')},'Client CRM'!$C$6:$C$305,0),1),"")`],
    [`=IFERROR(INDEX('Client CRM'!$H$6:$H$305,MATCH(${ilLookup(3).replace('=','')},'Client CRM'!$C$6:$C$305,0),1)&", "&INDEX('Client CRM'!$I$6:$I$305,MATCH(${ilLookup(3).replace('=','')},'Client CRM'!$C$6:$C$305,0),1)&" "&INDEX('Client CRM'!$J$6:$J$305,MATCH(${ilLookup(3).replace('=','')},'Client CRM'!$C$6:$C$305,0),1),"")`],
    [`=IFERROR(INDEX('Client CRM'!$K$6:$K$305,MATCH(${ilLookup(3).replace('=','')},'Client CRM'!$C$6:$C$305,0),1),"")`],
  ];

  // Simplify: pull client CRM via client ID from IL
  const clientIdFormula = `IFERROR(INDEX('Invoice Log'!$C$6:$C$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)),"")`;
  data.push({ range: `${S}!A6`, values: [[ilLookup(4)]] }); // Client Name
  data.push({ range: `${S}!A7`, values: [[`=IFERROR(INDEX('Client CRM'!$F$6:$F$305,MATCH(${clientIdFormula},'Client CRM'!$A$6:$A$305,0)),"")`]] });
  data.push({ range: `${S}!A8`, values: [[`=IFERROR(INDEX('Client CRM'!$G$6:$G$305,MATCH(${clientIdFormula},'Client CRM'!$A$6:$A$305,0)),"")`]] });
  data.push({ range: `${S}!A9`, values: [[`=IFERROR(INDEX('Client CRM'!$H$6:$H$305,MATCH(${clientIdFormula},'Client CRM'!$A$6:$A$305,0))&", "&INDEX('Client CRM'!$I$6:$I$305,MATCH(${clientIdFormula},'Client CRM'!$A$6:$A$305,0))&" "&INDEX('Client CRM'!$J$6:$J$305,MATCH(${clientIdFormula},'Client CRM'!$A$6:$A$305,0)),"")`]] });
  data.push({ range: `${S}!A10`, values: [[`=IFERROR(INDEX('Client CRM'!$K$6:$K$305,MATCH(${clientIdFormula},'Client CRM'!$A$6:$A$305,0)),"")`]] });
  data.push({ range: `${S}!A11`, values: [[`=IFERROR(INDEX('Client CRM'!$D$6:$D$305,MATCH(${clientIdFormula},'Client CRM'!$A$6:$A$305,0)),"")`]] }); // Email

  [5,6,7,8,9,10].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(sheetId, ri, ri + 1, 0, 4), mergeType: 'MERGE_ALL' } });
  });

  // Project / PO
  data.push({ range: `${S}!A12`, values: [['Project / PO:']] });
  data.push({ range: `${S}!C12`, values: [[`=IFERROR(TEXTJOIN("  ",TRUE,${ilLookup(9).replace('=','')},${ilLookup(8).replace('=','')}),"")`]] });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 11, 12, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat.textFormat' } });
  reqs.push({ mergeCells: { range: gridRange(sheetId, 11, 12, 2, 9), mergeType: 'MERGE_ALL' } });

  // Separator line before items
  reqs.push({ repeatCell: { range: gridRange(sheetId, 12, 13, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.header),
  }}, fields: 'userEnteredFormat.backgroundColor' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // Line items header row 14 (index 13)
  const liCols = ['#','Item / Service','Description','Qty','Unit','Unit Price','Disc %','Taxable','Line Total'];
  data.push({ range: `${S}!A14`, values: [liCols] });
  reqs.push({ repeatCell: { range: gridRange(sheetId, 13, 14, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.header), textFormat: { bold: true, foregroundColor: hex('#FFFFFF'), fontSize: 9 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // 20 line item rows (15-34, index 14-33)
  // These pull from a protected line-item archive block in Invoice Log (columns AC onward)
  // For the primary build, we note the manual archive workflow and show placeholders
  // populated from the first 20 payment rows matching the selected invoice
  for (let i = 0; i < 20; i++) {
    const r = 15 + i;
    // Line items are stored in Invoice Log columns AC:AV (line archive)
    // We reference Invoice Log archive cols: AC=Line#, AD=Service, AE=Desc, AF=Qty, AG=Unit, AH=UnitPrice, AI=Disc%, AJ=Taxable, AK=LineTotal
    const archiveBase = `IFERROR(INDEX('Invoice Log'!$AC$6:$AC$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`;
    // Show line number if non-empty service
    data.push({ range: `${S}!A${r}`, values: [[`=IFERROR(IF(INDIRECT("'Invoice Log'!AD"&(MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+5+${i}))="","",${i+1}),"")`]] });
    data.push({ range: `${S}!B${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AD$6:$AD$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AE$6:$AE$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AF$6:$AF$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });
    data.push({ range: `${S}!E${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AG$6:$AG$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });
    data.push({ range: `${S}!F${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AH$6:$AH$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });
    data.push({ range: `${S}!G${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AI$6:$AI$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });
    data.push({ range: `${S}!H${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AJ$6:$AJ$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });
    data.push({ range: `${S}!I${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$AK$6:$AK$1005,MATCH(${BS}!$B$30,'Invoice Log'!$A$6:$A$1005,0)+${i}),"")`]] });

    reqs.push({ repeatCell: { range: gridRange(sheetId, r - 1, r, 0, 9), cell: { userEnteredFormat: {
      backgroundColor: i % 2 === 0 ? hex(p.altRow) : hex(p.bg),
    }}, fields: 'userEnteredFormat.backgroundColor' } });
    // Merge B:C for service+desc, merge C is description
    reqs.push({ mergeCells: { range: gridRange(sheetId, r - 1, r, 2, 4), mergeType: 'MERGE_ALL' } });
  }

  // Totals (rows 36-44, index 35-43)
  const totRow = 36;
  reqs.push({ repeatCell: { range: gridRange(sheetId, totRow - 2, totRow - 1, 0, 9), cell: { userEnteredFormat: { backgroundColor: hex(p.header) } }, fields: 'userEnteredFormat.backgroundColor' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: totRow - 2, endIndex: totRow - 1 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  const totals = [
    ['Subtotal',              ilLookup(11)],
    ['Discount',              `=IFERROR(-1*${ilLookup(12).replace('=','')}+0,"")`],
    ['Shipping',              ilLookup(13)],
    ['Tax',                   ilLookup(14)],
    ['Grand Total',           ilLookup(15)],
    ['Deposit Required',      ilLookup(16)],
    ['Payments Received',     ilLookup(17)],
    ['Balance Due',           ilLookup(18)],
  ];
  totals.forEach(([label, formula], i) => {
    const r = totRow + i;
    reqs.push({ mergeCells: { range: gridRange(sheetId, r - 1, r, 0, 6), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(sheetId, r - 1, r, 6, 9), mergeType: 'MERGE_ALL' } });
    data.push({ range: `${S}!A${r}`, values: [[label]] });
    data.push({ range: `${S}!G${r}`, values: [[formula]] });
    const isBold = label === 'Grand Total' || label === 'Balance Due';
    reqs.push({ repeatCell: { range: gridRange(sheetId, r - 1, r, 0, 6), cell: { userEnteredFormat: {
      textFormat: { bold: isBold, foregroundColor: hex(label === 'Balance Due' ? p.accent : C.mainText) },
      horizontalAlignment: 'RIGHT',
      backgroundColor: isBold ? hex(label === 'Balance Due' ? p.accent : p.header) : hex(p.bg),
    }}, fields: 'userEnteredFormat(textFormat,horizontalAlignment,backgroundColor)' } });
    reqs.push({ repeatCell: { range: gridRange(sheetId, r - 1, r, 6, 9), cell: { userEnteredFormat: {
      textFormat: { bold: isBold, foregroundColor: isBold ? hex('#FFFFFF') : hex(C.mainText) },
      horizontalAlignment: 'RIGHT',
      backgroundColor: isBold ? hex(label === 'Balance Due' ? p.accent : p.header) : hex(p.altRow),
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00;[Red]-$#,##0.00' },
    }}, fields: 'userEnteredFormat(textFormat,horizontalAlignment,backgroundColor,numberFormat)' } });
  });

  // Payment instructions
  const piRow = totRow + totals.length + 2;
  reqs.push({ mergeCells: { range: gridRange(sheetId, piRow - 1, piRow, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, piRow - 1, piRow, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.header), textFormat: { bold: true, foregroundColor: hex('#FFFFFF'), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  data.push({ range: `${S}!A${piRow}`, values: [['PAYMENT INSTRUCTIONS']] });

  const piFields = [
    ['Payee', `=${BS}!B34`],
    ['Bank',  `=${BS}!B35`],
    ['Account',`=${BS}!B36`],
    ['PayPal',`=${BS}!B38`],
    ['Other', `=${BS}!B39`],
  ];
  piFields.forEach(([label, val], i) => {
    const r = piRow + 1 + i;
    data.push({ range: `${S}!A${r}`, values: [[label]] });
    data.push({ range: `${S}!C${r}`, values: [[val]] });
    reqs.push({ repeatCell: { range: gridRange(sheetId, r - 1, r, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 8 } } }, fields: 'userEnteredFormat.textFormat' } });
    reqs.push({ mergeCells: { range: gridRange(sheetId, r - 1, r, 2, 9), mergeType: 'MERGE_ALL' } });
  });

  // Notes + Terms
  const notesRow = piRow + piFields.length + 2;
  reqs.push({ mergeCells: { range: gridRange(sheetId, notesRow - 2, notesRow - 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, notesRow - 2, notesRow - 1, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.header), textFormat: { bold: true, foregroundColor: hex('#FFFFFF'), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  data.push({ range: `${S}!A${notesRow - 1}`, values: [['NOTES & TERMS']] });

  reqs.push({ mergeCells: { range: gridRange(sheetId, notesRow - 1, notesRow + 1, 0, 9), mergeType: 'MERGE_ALL' } });
  data.push({ range: `${S}!A${notesRow}`, values: [[ilLookup(28)]] }); // Notes from IL col AB
  reqs.push({ repeatCell: { range: gridRange(sheetId, notesRow - 1, notesRow + 1, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.altRow), wrapStrategy: 'WRAP', textFormat: { fontSize: 8 },
  }}, fields: 'userEnteredFormat(backgroundColor,wrapStrategy,textFormat)' } });

  reqs.push({ mergeCells: { range: gridRange(sheetId, notesRow + 1, notesRow + 3, 0, 9), mergeType: 'MERGE_ALL' } });
  data.push({ range: `${S}!A${notesRow + 2}`, values: [[`=${BS}!B20`]] }); // Terms from Business Setup
  reqs.push({ repeatCell: { range: gridRange(sheetId, notesRow + 1, notesRow + 3, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.bg), wrapStrategy: 'WRAP', textFormat: { italic: true, fontSize: 8 },
  }}, fields: 'userEnteredFormat(backgroundColor,wrapStrategy,textFormat)' } });

  // Thank-you row
  const tyRow = notesRow + 4;
  reqs.push({ mergeCells: { range: gridRange(sheetId, tyRow - 1, tyRow, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(sheetId, tyRow - 1, tyRow, 0, 9), cell: { userEnteredFormat: {
    backgroundColor: hex(p.accent), textFormat: { bold: true, foregroundColor: hex('#FFFFFF'), fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: tyRow - 1, endIndex: tyRow }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  data.push({ range: `${S}!A${tyRow}`, values: [['Thank you for your business!']] });

  // Sheet bg
  reqs.push({ repeatCell: { range: gridRange(sheetId, 0, 200, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex('#FFFFFF') } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Column widths
  [30, 160, 200, 50, 60, 80, 60, 60, 90].forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // Date format for invoice dates
  [4, 5].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(sheetId, ri, ri + 1, 6, 9), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  await valuesBatchUpdate(id, data, `${style.toLowerCase()}-values`);
  await batchUpdate(id, reqs, `${style.toLowerCase()}-format`);
  console.log(`✓ Invoice — ${style}`);
}

(async () => {
  await buildLayout(CLASSIC, 'Invoice — Classic', 'Classic');
  await buildLayout(MODERN,  'Invoice — Modern',  'Modern');
  await buildLayout(ELEGANT, 'Invoice — Elegant', 'Elegant');
  await buildLayout(COMPACT, 'Invoice — Compact', 'Compact');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function colLetter(idx) {
  let s = ''; idx += 1;
  while (idx > 0) { const r = (idx - 1) % 26; s = String.fromCharCode(65 + r) + s; idx = Math.floor((idx - 1) / 26); }
  return s;
}
