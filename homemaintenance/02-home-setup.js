'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Home & Property Setup'];
const S = "'Home & Property Setup'";
const REF = "'Reference Data'";
const COSTS = "'Maintenance Costs'";

const NC = 8; // columns A-H

const DISCLAIMER = 'Maintenance intervals, useful-life estimates, service recommendations, and replacement timing in this workbook are planning aids only. Follow manufacturer instructions, local codes, inspection findings, warranty requirements, and qualified professional guidance for your specific home and equipment.';

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['HOME & PROPERTY SETUP']] });

  // ── Subtitle ───────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Enter your property details, maintenance preferences, and dashboard controls here.']] });

  // ── Disclaimer ─────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.clay), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [[DISCLAIMER]] });

  // Helper: section header
  const secHdr = (row, label) => {
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, NC), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!A${row+1}`, values: [[`  ${label}`]] });
  };

  // Helper: label+value row
  const field = (row, label, value='', isFormula=false, isInput=true) => {
    // col A-B = label (merged), col C-H = value
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 2), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 2, NC), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    const valueBg = isFormula ? C.formula : C.input;
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 2, NC), cell: { userEnteredFormat: { backgroundColor: hex(valueBg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });
    if (value !== '') vals.push({ range: `${S}!C${row+1}`, values: [[value]] });
  };

  // Helper: bool field
  const boolField = (row, label) => {
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 2), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 2, 4), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
    fmt.push({ setDataValidation: { range: gridRange(SID, row, row+1, 2, 3), rule: { condition: { type: 'BOOLEAN' }, strict: true } } });
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });
  };

  // ── Section: Property Information (rows 3-24) ────────────────────────────
  secHdr(3, 'PROPERTY INFORMATION');
  field(4,  'Property ID',                 'PROP-001');
  field(5,  'Property Name',               'The Hartwell Residence');
  field(6,  'Home Type',                   'Single-Family Home');
  field(7,  'Address / Location Label',    '142 Maple Creek Drive');
  field(8,  'Year Built',                  '1998');
  field(9,  'Purchase Date',               '3/15/2019');
  field(10, 'Square Footage (approx.)',    '2,280');
  field(11, 'Number of Floors',            '2');
  field(12, 'Bedrooms',                    '3');
  field(13, 'Bathrooms',                   '2.5');
  field(14, 'Garage Type',                 'Attached 2-Car');
  boolField(15, 'Basement?');    vals.push({ range: `${S}!C16`, values: [[true]] });
  boolField(16, 'Attic?');       vals.push({ range: `${S}!C17`, values: [[true]] });
  boolField(17, 'Septic?');      vals.push({ range: `${S}!C18`, values: [[false]] });
  boolField(18, 'Well Water?');  vals.push({ range: `${S}!C19`, values: [[false]] });
  boolField(19, 'Central HVAC?'); vals.push({ range: `${S}!C20`, values: [[true]] });
  field(20, 'Primary Heating Type',        'Forced Air / Gas Furnace');
  field(21, 'Primary Cooling Type',        'Central Air Conditioner');
  field(22, 'Roof Type',                   'Asphalt Shingle');
  field(23, 'Primary Water Heater Type',   'Tank / Natural Gas');
  field(24, 'Notes',                       'Built 1998, extensively renovated 2019. New roof 2021.');

  // ── Section: Maintenance Preferences (rows 25-38) ────────────────────────
  secHdr(25, 'MAINTENANCE PREFERENCES');
  field(26, 'Preferred Week Start Day',    'Monday');
  field(27, 'Default Maintenance Day',     'Saturday');
  field(28, 'Reminder Lead Time (Days)',   '7');
  field(29, 'Annual Maintenance Budget',   '4800');
  field(30, 'Default Labor Rate ($/hr)',   '75');
  field(31, 'Default Tax Rate',            '8.5%');
  field(32, 'Homeowner / Primary Contact', 'Jamie Hartwell');
  field(33, 'Emergency Service Contact',   '(555) 911-0000');
  field(34, 'Preferred General Contractor','Ridgeline Home Services');
  field(35, 'Preferred HVAC Provider',     'AirPro HVAC LLC');
  field(36, 'Preferred Plumber',           'ClearFlow Plumbing');
  field(37, 'Preferred Electrician',       'Volt & Wire Electric');
  field(38, 'Preferred Roofer',            'Summit Ridge Roofing');
  field(39, 'Notes',                       'All vendors are vetted. HVAC contract renewed Jan 2026.');

  // ── Section: Dashboard Controls (rows 40-47) ─────────────────────────────
  secHdr(40, 'DASHBOARD CONTROLS');
  field(41, 'Selected Property ID',        'PROP-001');
  field(42, 'Reporting Year',              '2026');
  field(43, 'Reporting Month / Full Year', 'Full Year');
  field(44, 'Selected Season / All',       'All');
  field(45, 'Selected Home Area / All',    'All');
  field(46, 'Selected Category / All',     'All');
  field(47, 'Selected Asset / All',        'All');

  // Dropdowns on controls
  const dvList = (row, srcRange) => ({
    setDataValidation: {
      range: gridRange(SID, row, row+1, 2, NC),
      rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${srcRange}` }] }, showCustomUi: true },
    },
  });
  fmt.push(dvList(41, `${REF}!A2:A2`));     // dummy — Property ID is manual

  // ── Section: Summary Cards (rows 48-57) ──────────────────────────────────
  secHdr(48, 'PROPERTY SUMMARY');

  const MTL  = "'Maintenance Task Log'";
  const APP  = "'Appliances & Systems'";
  const REP  = "'Home Repair Log'";
  const prop = "'Home & Property Setup'!C42";  // Selected Property ID = row 42 (0-indexed 41), col C

  const YEAR_BUILT = `${S}!C9`; // row 9 = Year Built (0-indexed 8)
  const BUDGET     = `${S}!C30`; // Annual Budget (0-indexed 29)

  const cards = [
    ['Home Age (Years)',            `=IFERROR(YEAR(TODAY())-${YEAR_BUILT},"")`],
    ['Active Assets / Systems',     `=COUNTIFS(${APP}!$X$8:$X$507,TRUE,${APP}!$B$8:$B$507,C42)`],
    ['Open Maintenance Tasks',      `=COUNTIFS(${MTL}!$O$8:$O$3007,"<>Complete",${MTL}!$O$8:$O$3007,"<>Cancelled",${MTL}!$O$8:$O$3007,"<>Skipped",${MTL}!$B$8:$B$3007,C42,${MTL}!$F$8:$F$3007,"<>")`],
    ['Overdue Tasks',               `=COUNTIFS(${MTL}!$U$8:$U$3007,TRUE,${MTL}!$B$8:$B$3007,C42)`],
    ['Upcoming (Next 30 Days)',      `=COUNTIFS(${MTL}!$M$8:$M$3007,">="&TODAY(),${MTL}!$M$8:$M$3007,"<="&(TODAY()+30),${MTL}!$O$8:$O$3007,"<>Complete",${MTL}!$O$8:$O$3007,"<>Cancelled",${MTL}!$B$8:$B$3007,C42)`],
    ['Repairs This Year',           `=SUMPRODUCT((YEAR(IF(${REP}!$J$8:$J$2007="",DATE(1900,1,1),${REP}!$J$8:$J$2007))=C43)*(${REP}!$B$8:$B$2007=C42)*(${REP}!$G$8:$G$2007<>""))`],
    ['Maintenance Cost YTD',        `=IFERROR(SUMPRODUCT((YEAR(IF(${COSTS}!$C$8:$C$3007="",DATE(1900,1,1),${COSTS}!$C$8:$C$3007))=C43)*(${COSTS}!$B$8:$B$3007=C42)*(${COSTS}!$C$8:$C$3007<>"")*${COSTS}!$M$8:$M$3007),0)`],
    ['Annual Budget Remaining',     `=IFERROR(VALUE(${BUDGET})-SUMPRODUCT((YEAR(IF(${COSTS}!$C$8:$C$3007="",DATE(1900,1,1),${COSTS}!$C$8:$C$3007))=C43)*(${COSTS}!$B$8:$B$3007=C42)*(${COSTS}!$C$8:$C$3007<>"")*${COSTS}!$M$8:$M$3007),"")`],
  ];

  cards.forEach(([label, fml], i) => {
    const row = 49 + i;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 2), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 2, NC), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 2, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!C${row+1}`, values: [[fml]] });
  });
  // Currency format on rows 56, 57 (0-indexed 55, 56)
  [55, 56].forEach(r => {
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 2, NC), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Freeze rows 1-3 ────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row heights ────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 48 }, fields: 'pixelSize' } });
  // Section headers
  [3, 25, 40, 48].forEach(r => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  });
  // Field rows
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 58 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // ── Column widths ──────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 2 }, properties: { pixelSize: 160 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 2, endIndex: NC }, properties: { pixelSize: 130 }, fields: 'pixelSize' } });

  // Dropdowns for Home Type (row 6, 0-indexed 5)
  fmt.push({ setDataValidation: { range: gridRange(SID, 5, 6, 2, NC), rule: { condition: { type: 'ONE_OF_LIST', values: ['Single-Family Home','Condo','Townhome','Duplex','Mobile / Manufactured Home','Small Residential Property','Other'].map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });

  await batchUpdate(id, fmt, 'hs-fmt');
  await valuesBatchUpdate(id, vals, 'hs-vals');
  console.log('✓ Home & Property Setup complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
