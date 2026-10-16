'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BS = sheetMap['Business Setup'];
const S  = "'Business Setup'";

(async () => {
  const data = [];

  // Title
  data.push({ range: `${S}!A1`, values: [['BUSINESS SETUP']] });
  data.push({ range: `${S}!A2`, values: [['Enter your business information here. These fields populate all printable invoice layouts.']] });

  // Business Information
  data.push({ range: `${S}!A4`, values: [['BUSINESS INFORMATION']] });
  const bizFields = [
    ['Business Name',          'Skyline Creative Studio'],
    ['Owner / Contact Name',   'Morgan Ellis'],
    ['Address',                '4820 Birchwood Lane'],
    ['City',                   'Austin'],
    ['State / Province',       'TX'],
    ['Postal Code',            '78701'],
    ['Country',                'United States'],
    ['Phone',                  '(512) 555-0182'],
    ['Email',                  'hello@skylinecreative.com'],
    ['Website',                'www.skylinecreative.com'],
    ['Tax / Registration ID',  '47-8821053'],
    ['Currency',               'USD'],
    ['Default Tax Rate',       0.085],
    ['Default Payment Terms',  'Net 30'],
    ['Default Late Fee',       '1.5% per month after due date'],
    ['Invoice Prefix',         'INV'],
    ['Starting Invoice Number',1],
    ['Next Invoice Number',    '=B21+COUNTA(\'Invoice Log\'!A6:A1005)'],
    ['Default Notes',          'Thank you for your business! Payment is due within the specified terms. For questions, contact us at hello@skylinecreative.com.'],
    ['Default Terms & Conditions','Late payments are subject to a 1.5% monthly fee. Client is responsible for any costs incurred in collection. Rates subject to change with 30 days notice.'],
    ['Text Logo / Initials',   'SC'],
  ];
  bizFields.forEach(([label, val], i) => {
    data.push({ range: `${S}!A${i + 5}`, values: [[label]] });
    data.push({ range: `${S}!B${i + 5}`, values: [[val]] });
  });

  // Invoice Number Preview
  data.push({ range: `${S}!A27`, values: [['Invoice Number Preview']] });
  data.push({ range: `${S}!B27`, values: [['=IFERROR(B16&"-"&TEXT(TODAY(),"YYYY")&"-"&TEXT(B18,"0000"),"")']] });

  // Invoice Selector (for layouts)
  data.push({ range: `${S}!A29`, values: [['INVOICE SELECTOR']] });
  data.push({ range: `${S}!A30`, values: [['Selected Invoice ID']] });
  data.push({ range: `${S}!B30`, values: [['INV-2026-0001']] });
  data.push({ range: `${S}!A31`, values: [['Selected Layout']] });
  data.push({ range: `${S}!B31`, values: [['Classic']] });

  // Payment Instructions
  data.push({ range: `${S}!A33`, values: [['PAYMENT INSTRUCTIONS']] });
  const payFields = [
    ['Payee Name',            'Skyline Creative Studio'],
    ['Bank Name',             'First National Bank'],
    ['Masked Account Number', '****4821'],
    ['Routing Placeholder',   'Contact for ACH routing details'],
    ['PayPal Email',          'payments@skylinecreative.com'],
    ['Other Instructions',    'Checks payable to Skyline Creative Studio. Please include invoice number on check.'],
  ];
  payFields.forEach(([label, val], i) => {
    data.push({ range: `${S}!A${i + 34}`, values: [[label]] });
    data.push({ range: `${S}!B${i + 34}`, values: [[val]] });
  });

  // Disclaimer
  data.push({ range: `${S}!A42`, values: [['DISCLAIMER']] });
  data.push({ range: `${S}!A43`, values: [['This workbook is an organizational invoicing tool only. It does not provide accounting, tax, legal, or financial advice and does not replace professional bookkeeping records, contracts, tax filings, or payment-processing systems. Verify invoice requirements, tax treatment, payment terms, late fees, and recordkeeping obligations for your location and business.']] });

  await valuesBatchUpdate(id, data, 'businesssetup-values');

  const reqs = [];

  // Sheet bg
  reqs.push({ repeatCell: { range: gridRange(BS, 0, 200, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Title row
  reqs.push({ mergeCells: { range: gridRange(BS, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BS, 0, 1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BS, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // Subtitle
  reqs.push({ mergeCells: { range: gridRange(BS, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BS, 1, 2, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Section headers helper
  const secHeaderRows = [3, 28, 32, 41];
  secHeaderRows.forEach(r => {
    reqs.push({ mergeCells: { range: gridRange(BS, r, r + 1, 0, 8), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(BS, r, r + 1, 0, 8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  });

  // Label column
  reqs.push({ repeatCell: { range: gridRange(BS, 4, 42, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold: true, foregroundColor: hex(C.mainText) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Input cells (col B, rows 5-27 and 30-40)
  [[4, 27], [29, 40]].forEach(([r1, r2]) => {
    reqs.push({ repeatCell: { range: gridRange(BS, r1, r2, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Formula cell B27 (invoice preview) and B18 (next invoice number)
  reqs.push({ repeatCell: { range: gridRange(BS, 26, 27, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  reqs.push({ repeatCell: { range: gridRange(BS, 17, 18, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Tax rate format
  reqs.push({ repeatCell: { range: gridRange(BS, 12, 13, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Disclaimer box
  reqs.push({ mergeCells: { range: gridRange(BS, 42, 46, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(BS, 42, 46, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex('#E8E5DF'),
    textFormat: { italic: true, foregroundColor: hex(C.secText), fontSize: 9 },
    wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BS, dimension: 'ROWS', startIndex: 42, endIndex: 46 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId: BS, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 220 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BS, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 400 }, fields: 'pixelSize' } });

  // Freeze rows 1:3
  reqs.push({ updateSheetProperties: { properties: { sheetId: BS, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'businesssetup-format');
  console.log('✓ Business Setup');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
