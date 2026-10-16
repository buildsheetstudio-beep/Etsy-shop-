'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PS = sheetMap['🏛️ Professional Services & Fees'];

// 5 professional service entries
// Cols: A=# B=Name/Firm C=Type D=Contact E=Phone F=Email G=Engagement Date H=Scope of Work I=Quoted J=Invoiced K=Paid L=Balance M=Fee Status N=Invoice Ref O=Notes
const SERVICES = [
  [1, 'Mitchell & Partners Solicitors', 'Solicitor/Lawyer', 'Andrew Mitchell', '+61 2 9222 3456', 'a.mitchell@mitchellpartners.com.au', '2025-04-22', 'Probate application, executor advice, conveyancing, distribution', 18500, 9200, 9200, 9300, 'Invoiced', 'INV-2025-0441', 'Probate grant expected August 2025'],
  [2, 'Green & Associates Accountants', 'Accountant', 'Sandra Green', '+61 2 9888 5555', 's.green@greenassoc.com.au', '2025-05-01', 'Final personal tax return, estate tax return, ATO liaison', 6800, 3400, 3400, 3400, 'Invoiced', 'GRN-2025-0118', 'Final account due upon lodgement'],
  [3, 'Bay Valuation Pty Ltd', 'Valuer', 'Marcus Chen', '+61 2 9450 7890', 'm.chen@bayvaluation.com.au', '2025-05-02', 'Property valuations (2 properties) as at date of death', 3200, 3200, 3200, 0, 'Paid', 'BVP-3241', 'Valuations completed 2025-05-20'],
  [4, 'Coastal Real Estate — Mosman', 'Real Estate Agent', 'Julie Hargreaves', '+61 2 9456 1234', 'j.hargreaves@coastalre.com.au', '2025-06-01', 'Sale of 42 Wattle Street, Mosman NSW — estate property', 28000, 0, 0, 28000, 'Quoted', 'COA-EST-001', 'Commission 1.9% + GST on expected $1.48M sale'],
  [5, 'Southern Cross Funerals', 'Funeral Director', 'Patricia Lee', '+61 2 9200 0099', 'p.lee@southerncrossfunerals.com.au', '2025-04-13', 'Funeral service, cremation, death notices, certificates', 12400, 12400, 12400, 0, 'Paid', 'SCF-2025-0812', 'Paid from estate account 2025-05-02'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: PS, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  reqs.push({ updateDimensionProperties: { range: { sheetId: PS, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PS, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PS, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PS, dimension: 'ROWS', startIndex: 3, endIndex: 25 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Col widths: A=40 B=220 C=130 D=160 E=130 F=220 G=110 H=280 I=110 J=110 K=110 L=110 M=100 N=130 O=250
  const colW = [40, 220, 130, 160, 130, 220, 110, 280, 110, 110, 110, 110, 100, 130, 250];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: PS, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title
  reqs.push({ mergeCells: { range: gridRange(PS, 0, 1, 0, 15), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(PS, 0, 1, 0, 15), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Disclaimer
  reqs.push({ mergeCells: { range: gridRange(PS, 1, 2, 0, 15), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(PS, 1, 2, 0, 15), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.disclaimerBg),
      textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    },
  }, fields: 'userEnteredFormat' }});

  // Headers
  reqs.push({ repeatCell: { range: gridRange(PS, 2, 3, 0, 15), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Data rows
  for (let r = 0; r < 5; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(PS, r+3, r+4, 0, 15), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Totals row
  reqs.push({ repeatCell: { range: gridRange(PS, 8, 9, 0, 15), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.paleSage),
      textFormat: { foregroundColor: hex(C.mutedSage), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { top: { style: 'SOLID', color: hex(C.mutedSage), width: 2 }, bottom: { style: 'SOLID', color: hex(C.mutedSage), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Currency format for I-L cols
  reqs.push({ repeatCell: { range: gridRange(PS, 3, 9, 8, 12), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  // Center cols: A, C, F, G, L, M
  [0, 2, 5, 6, 11, 12].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(PS, 3, 9, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Borders
  reqs.push({ updateBorders: { range: gridRange(PS, 2, 9, 0, 15),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'prof-fmt');

  const vals = [];
  vals.push({ range: `'🏛️ Professional Services & Fees'!A1`, values: [['🏛️ PROFESSIONAL SERVICES & FEES — Estate of Margaret Eleanor Williams']] });
  vals.push({ range: `'🏛️ Professional Services & Fees'!A2`, values: [['IMPORTANT DISCLAIMER: Fee estimates are subject to change. All professional fees are an expense of the estate. Obtain written quotes and retain all invoices. Seek independent advice before engaging professional services.']] });
  vals.push({
    range: `'🏛️ Professional Services & Fees'!A3:O3`,
    values: [['#', 'Name / Firm', 'Service Type', 'Contact Person', 'Phone', 'Email', 'Engagement Date', 'Scope of Work', 'Quoted ($)', 'Invoiced ($)', 'Paid ($)', 'Balance ($)', 'Fee Status', 'Invoice Ref', 'Notes']],
  });
  vals.push({ range: `'🏛️ Professional Services & Fees'!A4`, values: SERVICES });
  vals.push({
    range: `'🏛️ Professional Services & Fees'!A9:O9`,
    values: [['TOTALS', '', '', '', '', '', '', '', '=SUM(I4:I8)', '=SUM(J4:J8)', '=SUM(K4:K8)', '=SUM(L4:L8)', '', '', '']],
  });

  await valuesBatchUpdate(id, vals, 'prof-vals');
  console.log('Professional Services complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
