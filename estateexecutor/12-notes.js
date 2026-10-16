'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const NC = sheetMap['📝 Notes & Contacts'];

// Key contacts (rows 4-13, 10 contacts)
// Cols: A=# B=Name C=Role D=Organisation E=Phone F=Email G=Address H=Notes
const CONTACTS = [
  [1, 'Robert James Williams', 'Beneficiary', 'N/A (Executor)', '+61 2 9345 6789', 'robert.williams@email.com.au', '12 Elm Street, Mosman NSW 2088', 'Executor & beneficiary (40%)'],
  [2, 'Catherine Anne Williams', 'Beneficiary', 'N/A', '+61 2 9876 5432', 'catherine.williams@email.com.au', '7 Rose Garden, Killara NSW 2071', 'Beneficiary (30%)'],
  [3, 'Thomas Edward Williams', 'Beneficiary', 'N/A', '+61 417 234 567', 'thomas.williams@gmail.com', '22 Banksia Avenue, Dee Why NSW 2099', 'Beneficiary (20%)'],
  [4, 'Andrew Mitchell', 'Solicitor', 'Mitchell & Partners Solicitors', '+61 2 9222 3456', 'a.mitchell@mitchellpartners.com.au', 'Level 12, 1 Macquarie Place, Sydney NSW 2000', 'Lead solicitor for estate — probate & distribution'],
  [5, 'Sandra Green', 'Accountant', 'Green & Associates', '+61 2 9888 5555', 's.green@greenassoc.com.au', 'Suite 5, 200 Pacific Highway, Crows Nest NSW 2065', 'Tax returns, ATO liaison'],
  [6, 'Marcus Chen', 'Financial Adviser', 'Bay Valuation Pty Ltd', '+61 2 9450 7890', 'm.chen@bayvaluation.com.au', '88 Bay Street, Mosman NSW 2088', 'Property valuations'],
  [7, 'Julie Hargreaves', 'Property Agent', 'Coastal Real Estate', '+61 2 9456 1234', 'j.hargreaves@coastalre.com.au', '42 Military Road, Mosman NSW 2088', 'Sale of 42 Wattle Street'],
  [8, 'Margaret Grace Williams Charitable Trust', 'Other', 'Trust', '+61 2 9000 1234', 'grants@mwcharitabletrust.org', 'Suite 4, 88 Collins Street Melbourne VIC 3000', 'Charitable beneficiary (10%) — per clause 8 of Will'],
  [9, 'NAB Estate Banking', 'Bank Contact', 'National Australia Bank', '+61 13 22 65', 'estates@nab.com.au', 'NAB Mosman Branch, 727 Military Rd, Mosman NSW 2088', 'Estate bank account — NAB BSB: 082-001'],
  [10, 'AIA Australia', 'Insurance Contact', 'AIA Australia', '+61 1800 333 613', 'claims@aia.com.au', 'PO Box 6111, Melbourne VIC 3004', 'Life insurance claim — POL-0099887'],
];

// General notes (rows 18-27, 10 notes)
// Cols: A=# B=Date C=Topic D=Note E=Action Required F=Status
const NOTES = [
  [1, '2025-04-12', 'General', 'Margaret Eleanor Williams passed at St Vincent\'s Hospital Sydney at 14:23. Death certificate obtained on 14 April 2025.', 'None', 'Resolved'],
  [2, '2025-04-22', 'Legal', 'Mitchell & Partners confirmed probate application to be lodged by 15 May 2025. Estimate Grant of Probate in 12-14 weeks from lodgement.', 'Monitor probate status', 'In Progress'],
  [3, '2025-04-30', 'Financial', 'ATO notified of death. TFN to be cancelled upon finalisation. Estate registered as a new taxpayer (TFN applied for via accountant).', 'Estate TFN application with accountant', 'In Progress'],
  [4, '2025-05-05', 'Financial', 'AIA Life Insurance claim submitted. Estimated 4-8 weeks for payment. Claim officer: Sarah Kim (+61 1800 333 613 ext 421).', 'Follow up with AIA by 3 June 2025', 'In Progress'],
  [5, '2025-05-28', 'Financial', 'AustralianSuper death benefit of $385,000 paid to estate bank account. Notify accountant for tax treatment assessment.', 'Accountant to assess tax implications', 'In Progress'],
  [6, '2025-06-01', 'Property', '42 Wattle Street listed for sale with Coastal Real Estate at $1.48M. Agent Julie Hargreaves managing. Open homes commence 7 June 2025.', 'Monitor sale progress weekly', 'In Progress'],
  [7, '2025-06-15', 'Legal', 'Probate ref PROB-2025-0442 confirmed. Court advised processing time approx 10-12 weeks. Solicitor to lodge conveyancing upon grant.', 'Await probate grant', 'In Progress'],
  [8, '2025-06-15', 'Family', 'Family meeting held. All 3 beneficiaries present. Discussed timeline, property sales, and interim distribution expectations. Meeting notes filed.', 'None', 'Resolved'],
  [9, '2025-06-20', 'General', 'Personal effects at family home to be divided by family agreement before property settlement. Catherine & Thomas to advise by 30 June 2025.', 'Beneficiaries to advise preference by 30 June', 'Open'],
  [10, '2025-06-20', 'Tax', 'Accountant advises estate income from AustralianSuper death benefit may be partially taxable depending on tax-free component. Seeking ATO ruling.', 'ATO ruling requested via Green & Associates', 'In Progress'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: NC, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 3, endIndex: 13 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 14, endIndex: 15 }, properties: { pixelSize: 16 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 17, endIndex: 27 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});

  // Col widths for contacts: A=40 B=200 C=140 D=200 E=140 F=230 G=250 H=260
  const colW = [40, 200, 140, 200, 140, 230, 250, 260];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: NC, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title
  reqs.push({ mergeCells: { range: gridRange(NC, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(NC, 0, 1, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // CONTACTS section header (row 1)
  reqs.push({ mergeCells: { range: gridRange(NC, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(NC, 1, 2, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      padding: { left: 12 },
    },
  }, fields: 'userEnteredFormat' }});

  // Contact col headers (row 2)
  reqs.push({ repeatCell: { range: gridRange(NC, 2, 3, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Contact data rows (3-12)
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(NC, r+3, r+4, 0, 8), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Notes section header (row 15)
  reqs.push({ mergeCells: { range: gridRange(NC, 15, 16, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(NC, 15, 16, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      padding: { left: 12 },
    },
  }, fields: 'userEnteredFormat' }});

  // Notes col headers (row 16)
  reqs.push({ repeatCell: { range: gridRange(NC, 16, 17, 0, 6), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Notes data rows (17-26)
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(NC, r+17, r+18, 0, 6), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Center: A, C, E in contacts; A, B, C, E, F in notes
  [0, 2, 4].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(NC, 3, 13, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));
  [0, 1, 2, 4, 5].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(NC, 17, 27, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Borders
  reqs.push({ updateBorders: { range: gridRange(NC, 2, 13, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});
  reqs.push({ updateBorders: { range: gridRange(NC, 16, 27, 0, 6),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'notes-fmt');

  const vals = [];
  vals.push({ range: `'📝 Notes & Contacts'!A1`, values: [['📝 NOTES & CONTACTS — Estate of Margaret Eleanor Williams (Ref: PROB-2025-0442)']] });
  vals.push({ range: `'📝 Notes & Contacts'!A2`, values: [['📇 KEY CONTACTS']] });
  vals.push({ range: `'📝 Notes & Contacts'!A3:H3`, values: [['#', 'Name', 'Role', 'Organisation', 'Phone', 'Email', 'Address', 'Notes']] });
  vals.push({ range: `'📝 Notes & Contacts'!A4`, values: CONTACTS });
  vals.push({ range: `'📝 Notes & Contacts'!A16`, values: [['📋 ESTATE NOTES']] });
  vals.push({ range: `'📝 Notes & Contacts'!A17:F17`, values: [['#', 'Date', 'Topic', 'Note', 'Action Required', 'Status']] });
  vals.push({ range: `'📝 Notes & Contacts'!A18`, values: NOTES });

  await valuesBatchUpdate(id, vals, 'notes-vals');
  console.log('Notes & Contacts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
