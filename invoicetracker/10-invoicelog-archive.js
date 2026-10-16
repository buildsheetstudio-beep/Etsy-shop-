'use strict';
// Adds line-item archive columns AC:AK to Invoice Log.
// These columns hold archived line items for the printable layouts to reference.
// Columns (0-indexed from A=0):
//   AC(28) = Line # archved
//   AD(29) = Item / Service
//   AE(30) = Description
//   AF(31) = Qty
//   AG(32) = Unit
//   AH(33) = Unit Price
//   AI(34) = Disc %
//   AJ(35) = Taxable?
//   AK(36) = Line Total
//
// Sample invoice INV-2026-0001 (row 6) gets 3 line items.
// Other invoices get at least one populated line so layouts show data.

const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IL = sheetMap['Invoice Log'];
const S  = "'Invoice Log'";

// Archive header row
const archiveHeaders = ['Line #','Item / Service','Description','Qty','Unit','Unit Price','Disc %','Taxable?','Line Total'];

// Sample line items per invoice row (row index in data, 0=row6)
// Format: [lineNum, service, description, qty, unit, unitPrice, discPct, taxable, lineTotal]
const lineItems = {
  0: [ // INV-2026-0001: Apex Digital — Brand Strategy
    [1,'Brand Strategy Session','Q1 brand positioning and discovery workshop',4,'Hour',195,0,'Yes',780],
    [2,'Logo Design — Premium','Full brand identity: logo, colors, type guide',1,'Project',1800,0,'Yes',1800],
    [3,'Social Media Post Pack','20 branded templates for Instagram/Facebook/LinkedIn',1,'Package',420,0,'Yes',420],
  ],
  1: [ // INV-2026-0002: BlueSky Architecture — Logo
    [1,'Logo Design — Standard','Primary logo + 2 alternates, vector files',1,'Project',850,0,'No',850],
  ],
  2: [ // INV-2026-0003: Driftwood Publishing — Books
    [1,'Logo Design — Premium','Book cover series — 3 volumes',3,'Project',850,0,'Yes',2550],
  ],
  3: [ // INV-2026-0004: Emerald — Marketing
    [1,'Social Media Post Pack','Q1 branded marketing templates',1,'Package',420,0,'Yes',420],
    [2,'Email Newsletter Design','Monthly newsletter template',1,'Each',240,0,'Yes',240],
    [3,'Early Payment Discount','5% early payment discount',-1,'Each',-24,0,'No',-24],
  ],
  4: [ // INV-2026-0005: Gilded Thread — Lookbook
    [1,'Photography — Full Day','Winter lookbook photo shoot',1,'Day',980,0,'Yes',980],
    [2,'Shipping & Handling','Delivery of printed proofs',1,'Each',45,0,'Yes',45],
  ],
  5: [ // INV-2026-0006: Juniper — Staging
    [1,'Photography — Half Day','Residential staging photos',4,'Hour',225,0,'Yes',900],
  ],
  6: [ // INV-2026-0007: Luminary — Slides
    [1,'Slide Deck Design','Custom branded course presentation',2,'Project',650,0.10,'Yes',1170],
  ],
  7: [ // INV-2026-0008: Northgate — Report
    [1,'Slide Deck Design','Q1 consulting annual report design',1,'Project',650,0,'Yes',650],
  ],
  8: [ // INV-2026-0009: Apex — Brand Refresh
    [1,'Logo Design — Premium','Full brand identity refresh',1,'Project',1800,0,'Yes',1800],
    [2,'Brand Strategy Session','Brand audit and positioning workshop',4,'Hour',195,0,'Yes',780],
    [3,'Business Card Design','Refreshed business card print files',3,'Each',150,0,'Yes',450],
    [4,'Social Media Post Pack','Refreshed brand social templates',1,'Package',420,0,'Yes',420],
    [5,'Brochure Design — Trifold','Company overview brochure',2,'Each',380,0,'Yes',760],
  ],
  9: [ // INV-2026-0010: Harbor — Annual Refresh
    [1,'Brand Strategy Session','Annual brand strategy workshops',8,'Hour',195,0,'Yes',1560],
    [2,'Logo Design — Premium','Complete brand identity refresh',1,'Project',1800,0,'Yes',1800],
    [3,'Slide Deck Design','Investor presentation deck',2,'Project',650,0,'Yes',1300],
    [4,'Website Copywriting','Website copy refresh — 8 pages',8,'Hour',120,0,'Yes',960],
    [5,'Photography — Full Day','Brand and property photography',1,'Day',980,0,'Yes',980],
  ],
};

(async () => {
  const data = [];

  // Archive header in row 5 (index 4)
  data.push({ range: `${S}!AC5`, values: [archiveHeaders] });

  // Write line items into archive columns for each invoice that has data
  Object.entries(lineItems).forEach(([rowOffset, lines]) => {
    const baseRow = 6 + parseInt(rowOffset); // IL data starts at row 6
    lines.forEach((line, li) => {
      const r = baseRow + li;
      data.push({ range: `${S}!AC${r}:AK${r}`, values: [line] });
    });
  });

  await valuesBatchUpdate(id, data, 'archive-values');

  const reqs = [];

  // Format archive header
  reqs.push({ repeatCell: { range: gridRange(IL, 4, 5, 28, 37), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Archive data styling
  reqs.push({ repeatCell: { range: gridRange(IL, 5, 1005, 28, 37), cell: { userEnteredFormat: {
    backgroundColor: hex('#F0EEF5'),
  }}, fields: 'userEnteredFormat.backgroundColor' } });

  // Archive column widths
  [40, 180, 200, 50, 70, 90, 60, 70, 90].forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: IL, dimension: 'COLUMNS', startIndex: 28 + i, endIndex: 29 + i },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Currency format for archive unit price and line total
  [33, 36].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(IL, 5, 1005, ci, ci + 1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Percent for disc %
  reqs.push({ repeatCell: { range: gridRange(IL, 5, 1005, 34, 35), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  await batchUpdate(id, reqs, 'archive-format');
  console.log('✓ Invoice Log archive columns (line items for printable layouts)');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
