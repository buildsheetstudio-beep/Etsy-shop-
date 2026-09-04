'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DR = sheetMap['📁 Documents Register'];

// 12 document entries
// Cols: A=# B=Document Name C=Category D=Date of Document E=Status F=Copies Held G=Location H=Notes
const DOCS = [
  [1, 'Original Will — Margaret Eleanor Williams', 'Will & Probate', '2022-08-15', 'Located', 3, 'Family home safe (original); Mitchell & Partners (certified copy); Estate file', 'Original Will dated 2022, signed by testatrix + 2 witnesses'],
  [2, 'Codicil No. 1 to Will', 'Will & Probate', '2023-11-20', 'Located', 2, 'Mitchell & Partners; Estate file', 'Minor amendment to charitable bequest percentage'],
  [3, 'Grant of Probate — Supreme Court NSW', 'Will & Probate', '', 'Required', 0, 'To be lodged with Mitchell & Partners', 'Application lodged — ref PROB-2025-0442'],
  [4, 'Death Certificate (original + certified copies)', 'Identity', '2025-04-14', 'Certified Copy', 8, 'Estate file (8 certified copies)', 'Ordered from NSW Registry of Births Deaths & Marriages'],
  [5, 'Deceased\'s Birth Certificate', 'Identity', '1950-09-05', 'Located', 1, 'Family home documents folder', ''],
  [6, 'Certificate of Title — 42 Wattle St Mosman NSW', 'Property', '2015-03-01', 'Located', 1, 'Heritage Bank (mortgagee holds); Mitchell & Partners (copy)', 'CT/1234567 — lodged with mortgagee Heritage Bank'],
  [7, 'Certificate of Title — 8 Harbour View Manly NSW', 'Property', '2018-07-12', 'Located', 1, 'Estate file', 'CT/7654321 — no mortgage'],
  [8, 'Property Valuations (both properties)', 'Property', '2025-05-20', 'Located', 2, 'Estate file; Green & Associates', 'Bay Valuation Pty Ltd — dated 12 April 2025 (DoD value)'],
  [9, 'Superannuation Member Statement & Nomination', 'Financial', '2024-06-30', 'Located', 2, 'Estate file; AustralianSuper', 'Binding nomination confirmed — estate as beneficiary'],
  [10, 'AIA Life Insurance Policy Documents', 'Insurance', '2019-01-15', 'Submitted', 1, 'AIA Australia (claim lodged); Estate file (copy)', 'Policy POL-0099887 — claim submitted 2025-05-05'],
  [11, 'Final Tax Return Documents 2024-25', 'Tax', '', 'Required', 0, 'With Green & Associates', 'All income records to be gathered for final personal return'],
  [12, 'Funeral Notice & Service Program', 'Funeral', '2025-04-19', 'Located', 3, 'Family copies; Estate file', 'Southern Cross Funerals — service 19 April 2025'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DR, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  reqs.push({ updateDimensionProperties: { range: { sheetId: DR, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DR, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DR, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DR, dimension: 'ROWS', startIndex: 3, endIndex: 25 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Col widths: A=40 B=280 C=130 D=120 E=120 F=90 G=300 H=300
  const colW = [40, 280, 130, 120, 120, 90, 300, 300];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: DR, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title
  reqs.push({ mergeCells: { range: gridRange(DR, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DR, 0, 1, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Quick summary bar
  reqs.push({ mergeCells: { range: gridRange(DR, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DR, 1, 2, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Headers
  reqs.push({ repeatCell: { range: gridRange(DR, 2, 3, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Data rows
  for (let r = 0; r < 12; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(DR, r+3, r+4, 0, 8), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Center cols: A, C, D, E, F
  [0, 2, 3, 4, 5].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(DR, 3, 15, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Borders
  reqs.push({ updateBorders: { range: gridRange(DR, 2, 15, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'docs-fmt');

  const vals = [];
  vals.push({ range: `'📁 Documents Register'!A1`, values: [['📁 DOCUMENTS REGISTER — Estate of Margaret Eleanor Williams (Ref: PROB-2025-0442)']] });
  vals.push({ range: `'📁 Documents Register'!A2`, values: [['=COUNTA(B4:B25)&" documents registered | "&COUNTIF(E4:E25,"Required")&" still required | "&COUNTIF(E4:E25,"Located")&" located"']] });
  vals.push({ range: `'📁 Documents Register'!A3:H3`, values: [['#', 'Document Name', 'Category', 'Document Date', 'Status', 'Copies Held', 'Location / Storage', 'Notes']] });
  vals.push({ range: `'📁 Documents Register'!A4`, values: DOCS });

  await valuesBatchUpdate(id, vals, 'docs-vals');
  console.log('Documents Register complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
