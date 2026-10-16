'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IG = sheetMap['Invoice Generator'];
const S  = "'Invoice Generator'";
const CRM = "'Client CRM'";
const PS  = "'Products & Services'";
const PT  = "'Payment Tracker'";

(async () => {
  const data = [];

  // Title
  data.push({ range: `${S}!A1`, values: [['INVOICE GENERATOR']] });
  data.push({ range: `${S}!A2`, values: [['Prepare one invoice here. When finalized, manually copy the summary row and line items into Invoice Log and save each payment into Payment Tracker. The generator resets when you start a new invoice.']] });
  data.push({ range: `${S}!A3`, values: [['Without Apps Script, the workbook cannot auto-append records. Copy row data to Invoice Log to archive each invoice permanently.']] });

  // ── Header fields (rows 6-24) ─────────────────────────────────────────────
  data.push({ range: `${S}!A5:F5`, values: [['INVOICE HEADER', '', 'CLIENT DETAILS', '', 'BILLING CONTROLS', '']] });

  const headerLeft = [
    ['Invoice ID',        '=IFERROR(\'Business Setup\'!B16&"-"&TEXT(TODAY(),"YYYY")&"-"&TEXT(\'Business Setup\'!B18,"0000"),"")'],
    ['Invoice Number',    '=IFERROR(\'Business Setup\'!B18,"")'],
    ['Client ID',         ''],
    ['Client Name',       '=IFERROR(INDEX(\'Client CRM\'!$B$6:$B$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Invoice Date',      ''],
    ['Due Date',          '=IFERROR(IF(B12="Due on Receipt",B10,IF(B12="Net 7",B10+7,IF(B12="Net 14",B10+14,IF(B12="Net 15",B10+15,IF(B12="Net 30",B10+30,IF(B12="Net 45",B10+45,IF(B12="Net 60",B10+60,B10))))))),"")'  ],
    ['Payment Terms',     '=IFERROR(IF(INDEX(\'Client CRM\'!$M$6:$M$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0))<>"",INDEX(\'Client CRM\'!$M$6:$M$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),\'Business Setup\'!B14),"Net 30")'],
    ['Currency',          '=IFERROR(INDEX(\'Client CRM\'!$O$6:$O$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"USD")'],
    ['PO / Reference',    ''],
    ['Project / Job',     ''],
    ['Delivery Method',   '=IFERROR(INDEX(\'Client CRM\'!$P$6:$P$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"Email")'],
    ['Invoice Status',    'Draft'],
    ['Tax Rate',          '=IFERROR(IF(INDEX(\'Client CRM\'!$N$6:$N$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0))>0,INDEX(\'Client CRM\'!$N$6:$N$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),\'Business Setup\'!B13),\'Business Setup\'!B13)'],
    ['Discount Type',     'None'],
    ['Discount Value',    0],
    ['Shipping',          0],
    ['Deposit Required',  0],
    ['Notes',             '=IFERROR(\'Business Setup\'!B19,"")'],
    ['Terms & Conditions','=IFERROR(\'Business Setup\'!B20,"")'],
  ];

  headerLeft.forEach(([label, val], i) => {
    data.push({ range: `${S}!A${i + 6}`, values: [[label]] });
    data.push({ range: `${S}!B${i + 6}`, values: [[val]] });
  });

  // Client detail lookups (cols D-E)
  const clientDetails = [
    ['Contact Name',      '=IFERROR(INDEX(\'Client CRM\'!$C$6:$C$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Email',             '=IFERROR(INDEX(\'Client CRM\'!$D$6:$D$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Phone',             '=IFERROR(INDEX(\'Client CRM\'!$E$6:$E$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Address 1',         '=IFERROR(INDEX(\'Client CRM\'!$F$6:$F$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Address 2',         '=IFERROR(INDEX(\'Client CRM\'!$G$6:$G$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['City',              '=IFERROR(INDEX(\'Client CRM\'!$H$6:$H$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['State / Province',  '=IFERROR(INDEX(\'Client CRM\'!$I$6:$I$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Postal Code',       '=IFERROR(INDEX(\'Client CRM\'!$J$6:$J$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Country',           '=IFERROR(INDEX(\'Client CRM\'!$K$6:$K$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
    ['Tax ID',            '=IFERROR(INDEX(\'Client CRM\'!$L$6:$L$305,MATCH(B10,\'Client CRM\'!$A$6:$A$305,0)),"")'],
  ];
  clientDetails.forEach(([label, val], i) => {
    data.push({ range: `${S}!D${i + 6}`, values: [[label]] });
    data.push({ range: `${S}!E${i + 6}`, values: [[val]] });
  });

  // ── Line Items header (row 27) ────────────────────────────────────────────
  const liHeaders = ['Line #','Item ID','Item / Service','Description','Qty','Unit','Unit Price','Discount %','Taxable?','Line Subtotal','Line Discount','Taxable Base','Line Tax','Line Total'];
  data.push({ range: `${S}!A27`, values: [['LINE ITEMS']] });
  data.push({ range: `${S}!A28`, values: [liHeaders] });

  // Line item rows 29-48 (20 lines)
  for (let i = 0; i < 20; i++) {
    const r = 29 + i;
    // Line #
    data.push({ range: `${S}!A${r}`, values: [[`=IF(C${r}="","",${ i + 1 })`]] });
    // Item ID input — col B (editable)
    // Item Name lookup
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(IF(B${r}="","",INDEX('Products & Services'!$B$6:$B$305,MATCH(B${r},'Products & Services'!$A$6:$A$305,0))),"")`]] });
    // Description lookup
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(IF(B${r}="","",INDEX('Products & Services'!$D$6:$D$305,MATCH(B${r},'Products & Services'!$A$6:$A$305,0))),"")`]] });
    // Qty — col E editable (default 1 if item set)
    data.push({ range: `${S}!E${r}`, values: [[`=IF(B${r}="","",1)`]] });
    // Unit lookup
    data.push({ range: `${S}!F${r}`, values: [[`=IFERROR(IF(B${r}="","",INDEX('Products & Services'!$E$6:$E$305,MATCH(B${r},'Products & Services'!$A$6:$A$305,0))),"")`]] });
    // Unit Price lookup
    data.push({ range: `${S}!G${r}`, values: [[`=IFERROR(IF(B${r}="","",INDEX('Products & Services'!$F$6:$F$305,MATCH(B${r},'Products & Services'!$A$6:$A$305,0))),"")`]] });
    // Discount % — col H editable (default 0)
    data.push({ range: `${S}!H${r}`, values: [[`=IF(B${r}="","",0)`]] });
    // Taxable? lookup
    data.push({ range: `${S}!I${r}`, values: [[`=IFERROR(IF(B${r}="","",IF(INDEX('Products & Services'!$G$6:$G$305,MATCH(B${r},'Products & Services'!$A$6:$A$305,0))="Taxable","Yes","No")),"")`]] });
    // Line Subtotal
    data.push({ range: `${S}!J${r}`, values: [[`=IFERROR(IF(B${r}="","",E${r}*G${r}),"")`]] });
    // Line Discount
    data.push({ range: `${S}!K${r}`, values: [[`=IFERROR(IF(J${r}="","",J${r}*H${r}),"")`]] });
    // Taxable Base
    data.push({ range: `${S}!L${r}`, values: [[`=IFERROR(IF(J${r}="","",MAX(0,J${r}-K${r})),"")`]] });
    // Line Tax
    data.push({ range: `${S}!M${r}`, values: [[`=IFERROR(IF(I${r}="Yes",L${r}*$B$18,0),0)`]] });
    // Line Total
    data.push({ range: `${S}!N${r}`, values: [[`=IFERROR(IF(L${r}="","",L${r}+M${r}),"")`]] });
  }

  // ── Totals section (rows 51-61) ───────────────────────────────────────────
  data.push({ range: `${S}!L51`, values: [['TOTALS']] });
  const totalsFormulas = [
    ['Subtotal',                 '=IFERROR(SUM(J29:J48),0)'],
    ['Line Discounts',           '=IFERROR(SUM(K29:K48),0)'],
    ['Invoice-Level Discount',   '=IFERROR(IF(B19="None",0,IF(B19="Percentage",(N52-N53)*B20,IF(B19="Fixed Amount",B20,0))),0)'],
    ['Shipping',                 '=IFERROR(B21,0)'],
    ['Tax',                      '=IFERROR(SUM(M29:M48),0)'],
    ['Grand Total',              '=IFERROR(N52-N53-N54+N55+N56,0)'],
    ['Deposit Required',         '=IFERROR(B22,0)'],
    ['Payments Received',        `=IFERROR(SUMIFS('Payment Tracker'!$N$6:$N$1505,'Payment Tracker'!$C$6:$C$1505,$B$6),0)`],
    ['Balance Due',              '=IFERROR(MAX(0,N57-N59),0)'],
    ['Amount Overpaid',          '=IFERROR(MAX(0,N59-N57),0)'],
  ];
  totalsFormulas.forEach(([label, formula], i) => {
    data.push({ range: `${S}!L${i + 52}`, values: [[label]] });
    data.push({ range: `${S}!N${i + 52}`, values: [[formula]] });
  });

  // ── Ready-to-Log Check (row 63) ───────────────────────────────────────────
  data.push({ range: `${S}!L63`, values: [['Ready-to-Log Check']] });
  data.push({ range: `${S}!N63`, values: [[
    '=IFERROR(IF(AND(B10<>"",B8<>"",B11<>"",B12<>"",COUNTA(B29:B48)>0,B57>0),"Ready","Missing Information"),"")'
  ]] });
  data.push({ range: `${S}!A65`, values: [['Prepare one invoice here, then save its finalized summary and line items into the Invoice Log/archive before reusing the generator. Without Apps Script, the workbook cannot automatically append records. Copy the data manually to archive each invoice.']] });

  await valuesBatchUpdate(id, data, 'generator-values');

  const reqs = [];

  reqs.push({ repeatCell: { range: gridRange(IG, 0, 70, 0, 15), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Title
  reqs.push({ mergeCells: { range: gridRange(IG, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 0, 1, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: IG, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // Subtitles
  [1,2].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(IG, ri, ri + 1, 0, 14), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(IG, ri, ri + 1, 0, 14), cell: { userEnteredFormat: {
      backgroundColor: ri === 1 ? hex(C.secondary) : hex(C.mutedBlue),
      textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  });

  // Section header row 5 (A5 and C5 and E5)
  reqs.push({ repeatCell: { range: gridRange(IG, 4, 5, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Label cols A and D
  reqs.push({ repeatCell: { range: gridRange(IG, 5, 26, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold: true, foregroundColor: hex(C.mainText) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 5, 16, 3, 4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold: true, foregroundColor: hex(C.mainText) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Editable inputs B6:B24 (except formula cells)
  reqs.push({ repeatCell: { range: gridRange(IG, 5, 25, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  // Formula cells: B6(ID), B7(num), B9(client name), B12(due date), B13(terms), B14(currency), B17(delivery), B18(tax), B24(notes), B25(terms)
  [5,6,8,11,12,13,16,17,23,24].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(IG, ri, ri + 1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });
  // Client detail lookups (col E)
  reqs.push({ repeatCell: { range: gridRange(IG, 5, 16, 4, 5), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Line items section header
  reqs.push({ mergeCells: { range: gridRange(IG, 26, 27, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 26, 27, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Line items col headers row 28
  reqs.push({ repeatCell: { range: gridRange(IG, 27, 28, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Line item data rows 29-48
  reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.panel) } }, fields: 'userEnteredFormat.backgroundColor' } });
  for (let r = 0; r < 20; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(IG, 28 + r, 29 + r, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Editable cols: B(Item ID), E(Qty), H(Disc%)
  reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, 4, 5), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, 6, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  // Formula cols: C,D,F,G,I,J,K,L,M,N
  [2,3,5,8,9,10,11,12,13].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, ci, ci + 1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Totals section
  reqs.push({ mergeCells: { range: gridRange(IG, 50, 51, 11, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 50, 51, 11, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 51, 62, 11, 12), cell: { userEnteredFormat: {
    textFormat: { bold: true, foregroundColor: hex(C.mainText) }, backgroundColor: hex(C.white),
  }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 51, 62, 13, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  // Grand total and Balance Due highlighted
  [56, 60].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(IG, ri, ri + 1, 11, 14), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  });
  // Ready check
  reqs.push({ repeatCell: { range: gridRange(IG, 62, 63, 11, 12), cell: { userEnteredFormat: {
    textFormat: { bold: true }, backgroundColor: hex(C.white),
  }}, fields: 'userEnteredFormat(textFormat,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 62, 63, 13, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.warning) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Instruction note (row 65)
  reqs.push({ mergeCells: { range: gridRange(IG, 64, 67, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 64, 67, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex('#E8E5DF'), textFormat: { italic: true, foregroundColor: hex(C.secText), fontSize: 9 },
    wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  // Number formats
  const currFmt = { type: 'CURRENCY', pattern: '$#,##0.00;[Red]-$#,##0.00' };
  reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, 6, 7), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, 9, 14), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 28, 48, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 51, 62, 13, 14), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell: { range: gridRange(IG, 17, 18, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths: A(50) B(90) C(180) D(200) E(50) F(70) G(100) H(80) I(70) J(100) K(100) L(110) M(90) N(100)
  [50,90,180,200,50,70,100,80,70,100,100,110,90,100].forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: IG, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  reqs.push({ updateSheetProperties: { properties: { sheetId: IG, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'generator-format');
  console.log('✓ Invoice Generator');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
