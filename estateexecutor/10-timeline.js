'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TL = sheetMap['📅 Estate Timeline & Probate Progress'];

// 20 milestones
// Cols: A=# B=Milestone C=Phase D=Target Date (formula from Dashboard!B3 + offset) E=Actual Date F=Status G=Responsible H=Days Remaining I=Notes
// Dashboard!B3 = Date of Death = 2025-04-12. Offsets are calendar days from DoD.
const MILESTONES = [
  // Completed (1-8)
  [1, 'Date of Death & Immediate Actions', 'Phase 1: Immediate', '=IFERROR(\'📋 Dashboard\'!$B$3,"—")', '2025-04-12', 'Completed', 'Robert Williams', '=IFERROR(D4-TODAY(),"—")', 'DoD confirmed — Margaret E. Williams'],
  [2, 'Funeral Arrangements Made', 'Phase 1: Immediate', '=IFERROR(\'📋 Dashboard\'!$B$3+3,"—")', '2025-04-15', 'Completed', 'Robert Williams', '=IFERROR(D5-TODAY(),"—")', 'Service held 2025-04-19'],
  [3, 'Will Located & Solicitor Engaged', 'Phase 1: Immediate', '=IFERROR(\'📋 Dashboard\'!$B$3+10,"—")', '2025-04-22', 'Completed', 'Robert Williams', '=IFERROR(D6-TODAY(),"—")', 'Mitchell & Partners engaged'],
  [4, 'Estate Bank Account Opened', 'Phase 2: First Month', '=IFERROR(\'📋 Dashboard\'!$B$3+14,"—")', '2025-04-22', 'Completed', 'Robert Williams', '=IFERROR(D7-TODAY(),"—")', 'NAB Estate Account'],
  [5, 'Asset & Liability Inventory Complete', 'Phase 2: First Month', '=IFERROR(\'📋 Dashboard\'!$B$3+21,"—")', '2025-04-30', 'Completed', 'Robert Williams', '=IFERROR(D8-TODAY(),"—")', '12 assets, 5 debts identified'],
  [6, 'Property Valuations Ordered', 'Phase 2: First Month', '=IFERROR(\'📋 Dashboard\'!$B$3+21,"—")', '2025-05-02', 'Completed', 'Robert Williams', '=IFERROR(D9-TODAY(),"—")', 'Bay Valuation Pty Ltd'],
  [7, 'Superannuation Death Benefit Claimed', 'Phase 2: First Month', '=IFERROR(\'📋 Dashboard\'!$B$3+30,"—")', '2025-05-28', 'Completed', 'Robert Williams', '=IFERROR(D10-TODAY(),"—")', '$385,000 received from AustralianSuper'],
  [8, 'Creditor Advertisements Published', 'Phase 3: 1-3 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+38,"—")', '2025-05-20', 'Completed', 'Mitchell & Partners', '=IFERROR(D11-TODAY(),"—")', 'NSW Gazette notice published'],
  // In Progress (9-12)
  [9, 'Property Valuations Completed', 'Phase 3: 1-3 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+45,"—")', '', 'In Progress', 'Bay Valuation Pty Ltd', '=IFERROR(D12-TODAY(),"—")', 'Draft valuations received, final pending'],
  [10, 'Probate Application Lodged', 'Phase 3: 1-3 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+50,"—")', '2025-06-01', 'In Progress', 'Mitchell & Partners', '=IFERROR(D13-TODAY(),"—")', 'Awaiting court processing — ref PROB-2025-0442'],
  [11, 'Final Personal Tax Return Filed', 'Phase 3: 1-3 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+90,"—")', '', 'In Progress', 'Green & Associates', '=IFERROR(D14-TODAY(),"—")', '2024-25 return with accountant'],
  [12, 'Real Property Listed for Sale', 'Phase 3: 1-3 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+90,"—")', '2025-06-15', 'In Progress', 'Coastal Real Estate', '=IFERROR(D15-TODAY(),"—")', '42 Wattle St listed at $1.48M'],
  // Not Started (13-20)
  [13, 'Grant of Probate Received', 'Phase 4: 3-6 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+120,"—")', '', 'Not Started', 'Mitchell & Partners', '=IFERROR(D16-TODAY(),"—")', 'Expected approx Aug 2025'],
  [14, 'Real Estate Sold — 42 Wattle St', 'Phase 4: 3-6 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+150,"—")', '', 'Not Started', 'Robert Williams', '=IFERROR(D17-TODAY(),"—")', ''],
  [15, 'Investment Property — Manly Transferred/Sold', 'Phase 4: 3-6 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+180,"—")', '', 'Not Started', 'Robert Williams', '=IFERROR(D18-TODAY(),"—")', ''],
  [16, 'All Estate Debts Paid', 'Phase 4: 3-6 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+180,"—")', '', 'Not Started', 'Robert Williams', '=IFERROR(D19-TODAY(),"—")', ''],
  [17, 'Estate Distribution Accounts Prepared', 'Phase 5: 6-12 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+210,"—")', '', 'Not Started', 'Mitchell & Partners', '=IFERROR(D20-TODAY(),"—")', ''],
  [18, 'Interim Distribution Made to Beneficiaries', 'Phase 5: 6-12 Months', '=IFERROR(\'📋 Dashboard\'!$B$3+240,"—")', '', 'Not Started', 'Robert Williams', '=IFERROR(D21-TODAY(),"—")', ''],
  [19, 'ATO Tax Clearance Obtained', 'Phase 6: Finalisation', '=IFERROR(\'📋 Dashboard\'!$B$3+300,"—")', '', 'Not Started', 'Green & Associates', '=IFERROR(D22-TODAY(),"—")', ''],
  [20, 'Final Distribution & Estate Wound Up', 'Phase 6: Finalisation', '=IFERROR(\'📋 Dashboard\'!$B$3+365,"—")', '', 'Not Started', 'Robert Williams', '=IFERROR(D23-TODAY(),"—")', 'Target: April 2026'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TL, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  reqs.push({ updateDimensionProperties: { range: { sheetId: TL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: TL, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: TL, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: TL, dimension: 'ROWS', startIndex: 3, endIndex: 25 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Col widths: A=40 B=280 C=170 D=110 E=110 F=110 G=160 H=110 I=260
  const colW = [40, 280, 170, 110, 110, 110, 160, 110, 260];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: TL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title
  reqs.push({ mergeCells: { range: gridRange(TL, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(TL, 0, 1, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Progress bar row
  reqs.push({ mergeCells: { range: gridRange(TL, 1, 2, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(TL, 1, 2, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Headers
  reqs.push({ repeatCell: { range: gridRange(TL, 2, 3, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Data rows
  for (let r = 0; r < 20; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(TL, r+3, r+4, 0, 9), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Date format for cols D, E
  reqs.push({ repeatCell: { range: gridRange(TL, 3, 23, 3, 5), cell: {
    userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  // Center cols A, C, D, E, F, G, H
  [0, 2, 3, 4, 5, 6, 7].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(TL, 3, 23, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Wrap col B, I
  [1, 8].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(TL, 3, 23, c, c+1),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat.wrapStrategy',
  }}));

  // Borders
  reqs.push({ updateBorders: { range: gridRange(TL, 2, 23, 0, 9),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'timeline-fmt');

  const vals = [];
  vals.push({ range: `'📅 Estate Timeline & Probate Progress'!A1`, values: [['📅 ESTATE TIMELINE & PROBATE PROGRESS — Estate of Margaret Eleanor Williams']] });
  vals.push({ range: `'📅 Estate Timeline & Probate Progress'!A2`, values: [['=COUNTIF(F4:F23,"Completed")&" / "&COUNTA(B4:B23)&" milestones complete | "&COUNTIF(F4:F23,"In Progress")&" in progress | Estate Ref: PROB-2025-0442"']] });
  vals.push({ range: `'📅 Estate Timeline & Probate Progress'!A3:I3`, values: [['#', 'Milestone', 'Phase', 'Target Date', 'Actual Date', 'Status', 'Responsible', 'Days Remaining', 'Notes']] });
  vals.push({ range: `'📅 Estate Timeline & Probate Progress'!A4`, values: MILESTONES });

  await valuesBatchUpdate(id, vals, 'timeline-vals');
  console.log('Timeline complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
