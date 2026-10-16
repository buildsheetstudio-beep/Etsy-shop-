'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RD = sheetMap['Reference Data'];
const S = "'Reference Data'";

(async () => {
  const data = [];

  // ── Column A: Vehicle Condition ──────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['Vehicle Condition']] });
  ['New','Certified Pre-Owned','Used','Private Party'].forEach((v,i) => {
    data.push({ range: `${S}!A${i+2}`, values: [[v]] });
  });

  // ── Column B: Vehicle Type ───────────────────────────────────────────────
  data.push({ range: `${S}!B1`, values: [['Vehicle Type']] });
  ['Sedan','Hatchback','SUV','Crossover','Pickup Truck','Minivan','Coupe','Wagon',
   'Electric Vehicle','Plug-In Hybrid','Hybrid','Luxury','Sports Car','Other'].forEach((v,i) => {
    data.push({ range: `${S}!B${i+2}`, values: [[v]] });
  });

  // ── Column C: Fuel Type ──────────────────────────────────────────────────
  data.push({ range: `${S}!C1`, values: [['Fuel Type']] });
  ['Gasoline','Diesel','Hybrid','Plug-In Hybrid','Electric','Other'].forEach((v,i) => {
    data.push({ range: `${S}!C${i+2}`, values: [[v]] });
  });

  // ── Column D: Purchase Method ─────────────────────────────────────────────
  data.push({ range: `${S}!D1`, values: [['Purchase Method']] });
  ['Cash','Finance','Lease'].forEach((v,i) => {
    data.push({ range: `${S}!D${i+2}`, values: [[v]] });
  });

  // ── Column E: Loan Terms (months) ────────────────────────────────────────
  data.push({ range: `${S}!E1`, values: [['Loan Term (Months)']] });
  [24,36,48,60,72,84].forEach((v,i) => {
    data.push({ range: `${S}!E${i+2}`, values: [[v]] });
  });

  // ── Column F: Ownership Periods (years) ───────────────────────────────────
  data.push({ range: `${S}!F1`, values: [['Ownership Period (Years)']] });
  [1,3,5,7].forEach((v,i) => {
    data.push({ range: `${S}!F${i+2}`, values: [[v]] });
  });

  // ── Column G: Quote Status ────────────────────────────────────────────────
  data.push({ range: `${S}!G1`, values: [['Quote Status']] });
  ['Researching','Quote Requested','Quote Received','Negotiating','Final Offer','Rejected','Selected'].forEach((v,i) => {
    data.push({ range: `${S}!G${i+2}`, values: [[v]] });
  });

  // ── Column H: Affordability Status ───────────────────────────────────────
  data.push({ range: `${S}!H1`, values: [['Affordability Status']] });
  ['Comfortable','Manageable','Tight','Over Budget','Insufficient Data'].forEach((v,i) => {
    data.push({ range: `${S}!H${i+2}`, values: [[v]] });
  });

  // ── Column I: Recommendation Status ──────────────────────────────────────
  data.push({ range: `${S}!I1`, values: [['Recommendation']] });
  ['Best Overall','Best Monthly Cost','Best Long-Term Value','Best Cash Option',
   'Best Finance Option','Best Lease Option','Review Carefully'].forEach((v,i) => {
    data.push({ range: `${S}!I${i+2}`, values: [[v]] });
  });

  // ── Column J: Yes/No ─────────────────────────────────────────────────────
  data.push({ range: `${S}!J1`, values: [['Yes / No']] });
  ['Yes','No'].forEach((v,i) => {
    data.push({ range: `${S}!J${i+2}`, values: [[v]] });
  });

  // ── Guided Assumptions: Columns L-N ──────────────────────────────────────
  data.push({ range: `${S}!L1:N1`, values: [['Assumption','Default Value','Note']] });

  const assumptions = [
    ['Annual Insurance (New)',             1800,   'Editable estimate — replace with your own quote or local figure.'],
    ['Annual Insurance (Used/CPO)',        1440,   'Editable estimate — replace with your own quote or local figure.'],
    ['Annual Maintenance (New, Yr 1–3)',    600,   'Editable estimate — replace with your own quote or local figure.'],
    ['Annual Maintenance (Used, Yr 1–3)',  1200,   'Editable estimate — replace with your own quote or local figure.'],
    ['Annual Maintenance (Any, Yr 4+)',    1800,   'Editable estimate — replace with your own quote or local figure.'],
    ['Annual Registration & Fees',          300,   'Editable estimate — replace with your own quote or local figure.'],
    ['Depreciation Rate — New, Year 1',    0.20,   'Editable estimate — replace with your own quote or local figure.'],
    ['Depreciation Rate — New, Yr 2–5',   0.12,   'Editable estimate — replace with your own quote or local figure.'],
    ['Depreciation Rate — Used/CPO Annual',0.10,   'Editable estimate — replace with your own quote or local figure.'],
    ['Fuel Price (per gallon, $)',          3.50,   'Editable estimate — replace with your own quote or local figure.'],
    ['Electricity Price (per kWh, $)',      0.16,   'Editable estimate — replace with your own quote or local figure.'],
    ['EV Efficiency (miles per kWh)',       3.5,    'Editable estimate — replace with your own quote or local figure.'],
    ['Default Annual Miles',              12000,   'Editable estimate — replace with your own quote or local figure.'],
    ['Lease Disposition Fee ($)',           395,   'Editable estimate — replace with your own quote or local figure.'],
    ['Excess Mileage Fee (per mile, $)',   0.25,   'Editable estimate — replace with your own quote or local figure.'],
    ['Monthly Parking Estimate ($)',          0,   'Editable estimate — replace with your own quote or local figure.'],
    ['Monthly Tolls Estimate ($)',            0,   'Editable estimate — replace with your own quote or local figure.'],
    ['Affordability Threshold — Manageable',0.15,  'Editable estimate — replace with your own quote or local figure.'],
    ['Affordability Threshold — Tight',    0.20,   'Editable estimate — replace with your own quote or local figure.'],
  ];

  assumptions.forEach(([label, val, note], i) => {
    data.push({ range: `${S}!L${i+2}:N${i+2}`, values: [[label, val, note]] });
  });

  await valuesBatchUpdate(id, data, 'reference-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Background for whole sheet
  reqs.push({ repeatCell: {
    range: gridRange(RD, 0, 500, 0, 35),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor'
  }});

  // Header row style
  const headerStyle = {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
    horizontalAlignment: 'CENTER',
  };
  reqs.push({ repeatCell: {
    range: gridRange(RD, 0, 1, 0, 11),
    cell: { userEnteredFormat: headerStyle },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
  }});
  reqs.push({ repeatCell: {
    range: gridRange(RD, 0, 1, 11, 14),
    cell: { userEnteredFormat: headerStyle },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
  }});

  // Assumption label col (L) — bold
  reqs.push({ repeatCell: {
    range: gridRange(RD, 1, 21, 11, 12),
    cell: { userEnteredFormat: { textFormat: { bold: true } } },
    fields: 'userEnteredFormat.textFormat'
  }});

  // Assumption value col (M) — input style
  reqs.push({ repeatCell: {
    range: gridRange(RD, 1, 21, 12, 13),
    cell: { userEnteredFormat: { backgroundColor: hex(C.input) } },
    fields: 'userEnteredFormat.backgroundColor'
  }});

  // Note col (N) — italic, muted
  reqs.push({ repeatCell: {
    range: gridRange(RD, 1, 21, 13, 14),
    cell: { userEnteredFormat: { textFormat: { italic: true, foregroundColor: hex(C.secText) }, wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(textFormat,wrapStrategy)'
  }});

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId: RD, dimension: 'COLUMNS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 260 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: RD, dimension: 'COLUMNS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 110 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: RD, dimension: 'COLUMNS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 360 }, fields: 'pixelSize' }});

  // Number format for percentages in assumptions
  reqs.push({ repeatCell: {
    range: gridRange(RD, 7, 10, 12, 13),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat'
  }});
  reqs.push({ repeatCell: {
    range: gridRange(RD, 17, 19, 12, 13),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat'
  }});

  // Freeze row 1
  reqs.push({ updateSheetProperties: { properties: { sheetId: RD, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' }});

  // Hide this tab
  reqs.push({ updateSheetProperties: { properties: { sheetId: RD, hidden: true }, fields: 'hidden' }});

  await batchUpdate(id, reqs, 'reference-format');
  console.log('✓ Reference Data');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
