'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

const LISTS = {
  'Home Types':              ['Single-Family Home','Condo','Townhome','Duplex','Mobile / Manufactured Home','Small Residential Property','Other'],
  'Home Areas':              ['Whole Home','Kitchen','Bathroom','Bedroom','Living Room','Dining Room','Laundry','Basement','Attic','Garage','Exterior','Roof','Yard / Landscaping','Driveway','Patio / Deck','HVAC','Plumbing','Electrical','Safety','Other'],
  'Maintenance Categories':  ['HVAC','Plumbing','Electrical','Roof & Gutters','Exterior','Interior','Appliances','Safety','Landscaping','Cleaning','Pest Prevention','Water Systems','Energy Efficiency','Seasonal Preparation','General Inspection','Other'],
  'Task Statuses':           ['Not Started','Scheduled','In Progress','Waiting','Complete','Skipped','Cancelled'],
  'Priorities':              ['Low','Medium','High','Urgent'],
  'Seasons':                 ['Winter','Spring','Summer','Fall','Year-Round'],
  'Recurrence Frequencies':  ['Weekly','Monthly','Every 2 Months','Quarterly','Every 4 Months','Semiannual','Annual','Every 18 Months','Every 2 Years','Every 3 Years','Every 5 Years','Custom Months','Custom Years','Seasonal','Usage-Based','One-Time'],
  'Usage Units':             ['Hours','Months','Miles','Cycles','Filters','Gallons','Other'],
  'Asset Types':             ['HVAC System','Furnace','Boiler','Air Conditioner','Heat Pump','Water Heater','Refrigerator','Freezer','Oven / Range','Dishwasher','Microwave','Washer','Dryer','Garbage Disposal','Sump Pump','Water Softener','Well Pump','Septic System','Roof','Gutters','Windows','Exterior Doors','Garage Door','Smoke Detector','Carbon Monoxide Detector','Fire Extinguisher','Electrical Panel','Plumbing Fixture','Irrigation System','Lawn Equipment','Generator','Other'],
  'Asset Conditions':        ['Excellent','Good','Fair','Monitor','Service Recommended','Replacement Planning','Out of Service'],
  'Repair Statuses':         ['Reported','Diagnosing','Scheduled','In Progress','Waiting for Parts','Complete','Monitoring','Cancelled'],
  'Repair Types':            ['Preventive','Corrective','Emergency','Cosmetic','Replacement','Inspection','Other'],
  'Service Provider Types':  ['HVAC','Plumber','Electrician','Roofer','Handyman','Appliance Repair','Landscaper','Pest Control','Cleaning','Locksmith','Garage Door','General Contractor','Inspector','Other'],
  'Cost Categories':         ['Preventive Maintenance','Repair Labor','Parts & Materials','Contractor Service','Inspection','Replacement','Permit / Fee','Tool / Equipment','Landscaping','Cleaning','Other'],
  'Payment Statuses':        ['Planned','Estimate Received','Due','Paid','Partially Paid','Refunded','Cancelled'],
  'Warranty Statuses':       ['Active','Expiring Soon','Expired','Unknown','Not Applicable'],
  'Yes / No':                ['Yes','No'],
  'DIY or Professional':     ['DIY','Professional','Mixed','Not Decided'],
};

// Seasonal anchor dates for recurrence planning
const SEASONAL_ANCHORS = [
  ['Season','Anchor Date','Note'],
  ['Winter','1/15','Planning anchor only — verify for your property'],
  ['Spring','4/15','Planning anchor only — verify for your property'],
  ['Summer','7/15','Planning anchor only — verify for your property'],
  ['Fall','10/15','Planning anchor only — verify for your property'],
];

(async () => {
  const fmt = [];
  const vals = [];

  const listNames = Object.keys(LISTS);
  const numCols = listNames.length;

  // Ensure enough columns (default 26, need up to ~18)
  // No expansion needed since 18 < 26

  // ── Header row ─────────────────────────────────────────────────────────────
  listNames.forEach((name, ci) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+ci)}1`, values: [[name]] });
  });

  // ── List values ────────────────────────────────────────────────────────────
  listNames.forEach((name, ci) => {
    const items = LISTS[name];
    const col = String.fromCharCode(65+ci);
    items.forEach((item, ri) => {
      vals.push({ range: `${S}!${col}${ri+2}`, values: [[item]] });
    });
    fmt.push({ repeatCell: { range: gridRange(SID, 1, 1+items.length, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  });

  // ── Seasonal anchors section (starting col 0, row after longest list + 3) ──
  const maxItems = Math.max(...Object.values(LISTS).map(a => a.length));
  const anchorStartRow = maxItems + 3; // 0-indexed
  fmt.push({ repeatCell: { range: gridRange(SID, anchorStartRow, anchorStartRow+1, 0, 3), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  SEASONAL_ANCHORS.forEach((row, ri) => {
    vals.push({ range: `${S}!A${anchorStartRow+ri+1}:C${anchorStartRow+ri+1}`, values: [row] });
  });

  // ── Freeze row 1 ──────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Column widths ──────────────────────────────────────────────────────────
  for (let i = 0; i < numCols; i++) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: 160 }, fields: 'pixelSize' } });
  }

  await batchUpdate(id, fmt, 'ref-fmt');
  await valuesBatchUpdate(id, vals, 'ref-vals');

  // ── Hide the tab ──────────────────────────────────────────────────────────
  await batchUpdate(id, [{ updateSheetProperties: { properties: { sheetId: SID, hidden: true }, fields: 'hidden' } }], 'ref-hide');

  console.log('✓ Reference Data complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
