'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CL = sheetMap['📞 Beneficiary Communication Log'];

// 6 communication entries
// Cols: A=# B=Date C=Beneficiary/Party D=Type E=Method F=Summary G=Follow-up Required H=Follow-up Date I=Follow-up Done K=Logged By
const COMMS = [
  [1, '2025-04-15', 'Robert James Williams', 'Initial Notification', 'In Person', 'Met with Robert at family home. Discussed role as executor and initial steps. Provided copy of Will and death certificate.', true, '2025-04-22', true, 'Solicitor — A. Mitchell'],
  [2, '2025-04-15', 'Catherine Anne Williams', 'Initial Notification', 'Phone', 'Called Catherine to advise of mother\'s passing and Will provisions. Confirmed she received copy by email.', false, '', true, 'Robert Williams'],
  [3, '2025-04-15', 'Thomas Edward Williams', 'Initial Notification', 'In Person', 'Visited Thomas. Discussed funeral arrangements and role as beneficiary. Provided copy of relevant Will clauses.', false, '', true, 'Robert Williams'],
  [4, '2025-04-20', 'All Beneficiaries', 'Initial Notification', 'Letter', 'Formal executor notification letters sent via registered mail as required by law. Letters included summary of estate process and timeline.', true, '2025-05-10', true, 'Mitchell & Partners Solicitors'],
  [5, '2025-05-30', 'Catherine Anne Williams', 'Update', 'Email', 'Updated Catherine on probate lodgement status and estimated timeline for first distribution. Responded to questions about investment portfolio.', false, '', true, 'Robert Williams'],
  [6, '2025-06-15', 'Robert James Williams', 'Meeting', 'In Person', 'Meeting with solicitor and accountant to review estate accounts. Robert as executor reviewed preliminary estate values and creditor list.', true, '2025-07-01', false, 'Mitchell & Partners Solicitors'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: CL, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  reqs.push({ updateDimensionProperties: { range: { sheetId: CL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: CL, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: CL, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: CL, dimension: 'ROWS', startIndex: 3, endIndex: 30 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Col widths: A=40 B=100 C=200 D=140 E=100 F=340 G=120 H=110 I=100 J=170
  const colW = [40, 100, 200, 140, 100, 340, 120, 110, 100, 170];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: CL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title
  reqs.push({ mergeCells: { range: gridRange(CL, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CL, 0, 1, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Quick summary bar
  reqs.push({ mergeCells: { range: gridRange(CL, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CL, 1, 2, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Headers
  reqs.push({ repeatCell: { range: gridRange(CL, 2, 3, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Data rows
  for (let r = 0; r < 6; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(CL, r+3, r+4, 0, 10), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Center cols A, B, D, E, G, H, I
  [0, 1, 3, 4, 6, 7, 8].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(CL, 3, 9, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Borders
  reqs.push({ updateBorders: { range: gridRange(CL, 2, 9, 0, 10),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'comm-fmt');

  const vals = [];
  vals.push({ range: `'📞 Beneficiary Communication Log'!A1`, values: [['📞 BENEFICIARY COMMUNICATION LOG — Estate of Margaret Eleanor Williams']] });
  vals.push({ range: `'📞 Beneficiary Communication Log'!A2`, values: [['=COUNTA(B4:B30)&" entries logged | "&COUNTIF(I4:I30,FALSE)&" follow-ups outstanding"']] });
  vals.push({ range: `'📞 Beneficiary Communication Log'!A3:J3`, values: [['#', 'Date', 'Beneficiary / Party', 'Comm Type', 'Method', 'Summary', 'Follow-up Required?', 'Follow-up Date', 'Follow-up Done?', 'Logged By']] });
  vals.push({ range: `'📞 Beneficiary Communication Log'!A4`, values: COMMS });

  await valuesBatchUpdate(id, vals, 'comm-vals');
  console.log('Communication Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
