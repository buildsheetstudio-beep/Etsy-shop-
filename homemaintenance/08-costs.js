'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Maintenance Costs'];
const S = "'Maintenance Costs'";
const TASK = "'Maintenance Task Log'";
const REP = "'Home Repair Log'";
const SETUP = "'Home & Property Setup'";
const NC = 20; // A-T

// Col mapping:
// A(0)=Cost ID    B(1)=Property ID   C(2)=Date           D(3)=Category
// E(4)=Home Area  F(5)=Task/Repair   G(6)=Asset ID       H(7)=Description
// I(8)=Type       J(9)=Vendor/Provider K(10)=Amount      L(11)=Tax/Fees
// M(12)=Total(fml) N(13)=DIY or Pro  O(14)=Payment Status P(15)=Year(fml)
// Q(16)=Month(fml) R(17)=Budget Category S(18)=Reimbursable T(19)=Notes

const E = (propId, date, cat, area, taskRepair, assetId, desc, type, vendor, amount, tax, diyPro, payStatus, budgetCat, reimbursable, notes) => {
  const row = Array(NC).fill('');
  row[1] = propId;
  row[2] = date;
  row[3] = cat;
  row[4] = area;
  row[5] = taskRepair;
  row[6] = assetId;
  row[7] = desc;
  row[8] = type;
  row[9] = vendor;
  row[10] = amount;
  row[11] = tax;
  // M(12) = Total = formula
  row[13] = diyPro;
  row[14] = payStatus;
  // P(15) = Year formula
  // Q(16) = Month formula
  row[17] = budgetCat;
  row[18] = reimbursable;
  row[19] = notes;
  return row;
};

const COSTS = [
  // ── 2023 costs ─────────────────────────────────────────────────────────────
  E('PROP-001','3/3/2023','Plumbing','Plumbing','REP-0001','AST-0003','Water heater thermocouple — labor','Repair Labor','ClearFlow Plumbing',140,0,'Professional','Paid','Repair Labor',false,'Thermocouple replacement labor'),
  E('PROP-001','3/3/2023','Plumbing','Plumbing','REP-0001','AST-0003','Water heater thermocouple — parts','Parts & Materials','Supply House',45,3.60,'Professional','Paid','Parts & Materials',false,'Thermocouple part — Rheem compatible'),
  E('PROP-001','8/12/2023','HVAC','HVAC','REP-0007','AST-0002','AC condensate drain flush — parts','Parts & Materials','Self',0,0,'DIY','Paid','Preventive Maintenance',false,'Bleach solution (household supplies)'),
  E('PROP-001','9/15/2023','Interior','Whole Home','REP-0018','AST-0021','Window IGU replacement — parts','Parts & Materials','Andersen',280,22.40,'Professional','Paid','Replacement',true,'3 IGU replacements — covered by warranty'),
  E('PROP-001','9/18/2023','Interior','Whole Home','REP-0018','AST-0021','Window IGU replacement — labor','Repair Labor','Ridgeline Home Services',320,0,'Professional','Paid','Repair Labor',true,'Installation labor — covered by warranty'),
  E('PROP-001','11/15/2023','Plumbing','Bathroom','','','Toilet flappers (x2) — parts','Parts & Materials','Hardware Store',12,0.96,'DIY','Paid','Parts & Materials',false,'Standard flappers for 1.6 GPF toilets'),
  // ── 2024 costs ─────────────────────────────────────────────────────────────
  E('PROP-001','2/15/2024','Interior','Garage','REP-0008','AST-0018','Garage door torsion springs (x2) — parts','Parts & Materials','Supply House',95,7.60,'Professional','Paid','Parts & Materials',false,'Both springs replaced as pair'),
  E('PROP-001','2/15/2024','Interior','Garage','REP-0008','AST-0018','Garage door spring replacement — labor','Repair Labor','Ridgeline Home Services',165,0,'Professional','Paid','Repair Labor',false,'Spring replacement + door balance adjustment'),
  E('PROP-001','3/10/2024','Appliances','Kitchen','REP-0013','AST-0007','Microwave door latch replacement — parts','Parts & Materials','AppliancePro Repair',28,2.24,'Professional','Paid','Parts & Materials',false,'Door latch assembly'),
  E('PROP-001','3/11/2024','Appliances','Kitchen','REP-0013','AST-0007','Microwave door latch — labor','Repair Labor','AppliancePro Repair',65,0,'Professional','Paid','Repair Labor',false,'Technician labor — 1 hour'),
  E('PROP-001','4/5/2024','Plumbing','Basement','REP-0015','AST-0011','Sump pump backup battery replacement','Replacement','Hardware Store',95,7.60,'DIY','Paid','Replacement',false,'AGM 12V battery for battery backup unit'),
  E('PROP-001','5/18/2024','Interior','Bathroom','REP-0012','','Bathtub caulk replacement — materials','Parts & Materials','Hardware Store',8,0.64,'DIY','Paid','Parts & Materials',false,'100% silicone caulk — GE Advanced'),
  E('PROP-001','6/10/2024','Exterior','Basement','REP-0009','','Foundation crack sealing — materials','Parts & Materials','Hardware Store',35,2.80,'DIY','Paid','Parts & Materials',false,'Hydraulic cement and waterproof sealant'),
  E('PROP-001','7/20/2024','Appliances','Kitchen','REP-0011','AST-0004','Refrigerator condenser fan service — labor','Repair Labor','AppliancePro Repair',95,0,'Professional','Paid','Repair Labor',false,'Diagnostic + cleaning + fan tighten'),
  E('PROP-001','7/22/2024','Appliances','Kitchen','REP-0011','AST-0004','Refrigerator condenser fan — parts','Parts & Materials','AppliancePro Repair',35,2.80,'Professional','Paid','Parts & Materials',false,'Fan mount hardware'),
  E('PROP-001','11/1/2024','HVAC','HVAC','','AST-0001','Furnace annual tune-up — labor','Contractor Service','AirPro HVAC LLC',0,0,'Professional','Paid','Contractor Service',false,'Annual furnace service — included in maintenance contract'),
  // ── 2025 costs ─────────────────────────────────────────────────────────────
  E('PROP-001','3/14/2025','Energy Efficiency','Whole Home','','','Home energy audit','Inspection','UtilityFirst Energy',0,0,'Professional','Paid','Inspection',false,'Blower door test; utility-subsidized'),
  E('PROP-001','4/1/2025','Landscaping','Yard / Landscaping','','','Pre-emergent weed control (Scotts Halts)','Preventive Maintenance','Garden Center',40,3.20,'DIY','Paid','Preventive Maintenance',false,'50 lb bag — covers ~5,000 sq ft'),
  E('PROP-001','4/1/2025','Interior','Attic','REP-0020','','Attic air sealing — materials','Parts & Materials','Hardware Store',35,2.80,'DIY','Paid','Parts & Materials',false,'Foam sealant + caulk for top plate sealing'),
  E('PROP-001','4/3/2025','HVAC','HVAC','','AST-0002','AC condensate drain flush — supplies','Preventive Maintenance','Self',0,0,'DIY','Paid','Preventive Maintenance',false,'Bleach from household supplies'),
  E('PROP-001','4/6/2025','Landscaping','Yard / Landscaping','','AST-0025','Lawn mower spring service — parts','Parts & Materials','Hardware Store',38,3.04,'DIY','Paid','Parts & Materials',false,'Oil, spark plug, air filter'),
  E('PROP-001','4/16/2025','Landscaping','Yard / Landscaping','REP-0003','AST-0023','Irrigation head replacement — parts','Parts & Materials','Irrigation Supply',18,1.44,'DIY','Paid','Parts & Materials',false,'Rain Bird 5000 series head'),
  E('PROP-001','4/18/2025','Landscaping','Yard / Landscaping','','','Lawn fertilizer — spring (Milorganite)','Preventive Maintenance','Garden Center',55,4.40,'DIY','Paid','Preventive Maintenance',false,'32 lb bag Milorganite'),
  E('PROP-001','4/22/2025','Pest Prevention','Exterior','','','Pest control — spring perimeter treatment','Contractor Service','HomeShield Pest Control',150,0,'Professional','Paid','Contractor Service',false,'Exterior spray for ants, spiders, wasps'),
  E('PROP-001','4/26/2025','Exterior','Exterior','REP-0004','','Rear door bottom sweep replacement','Parts & Materials','Hardware Store',14,1.12,'DIY','Paid','Parts & Materials',false,'Vinyl door sweep'),
  E('PROP-001','4/30/2025','Exterior','Patio / Deck','','','Deck rail post hardware — bolts','Parts & Materials','Hardware Store',5,0.40,'DIY','Paid','Parts & Materials',false,'Carriage bolts and washers'),
  E('PROP-001','4/30/2025','Exterior','Exterior','','','Exterior caulking — window & door seals','Parts & Materials','Hardware Store',18,1.44,'DIY','Paid','Parts & Materials',false,'3 tubes paintable exterior caulk'),
  E('PROP-001','5/1/2025','Landscaping','Yard / Landscaping','','','Mulch — 3 cubic yards hardwood','Landscaping','Landscape Supply',72,5.76,'DIY','Paid','Landscaping',false,'Brown hardwood mulch — delivery included'),
  E('PROP-001','5/4/2025','Landscaping','Yard / Landscaping','','','Mulch delivery charge','Contractor Service','Landscape Supply',0,0,'Professional','Paid','Contractor Service',false,'Delivery included in mulch price'),
  E('PROP-001','5/15/2025','HVAC','HVAC','','AST-0002','AC tune-up — spring 2025','Contractor Service','AirPro HVAC LLC',160,0,'Professional','Paid','Contractor Service',false,'Refrigerant check, coil cleaning, drain flush'),
  E('PROP-001','9/1/2025','Landscaping','Yard / Landscaping','','','Lawn overseeding — seed & supplies','Landscaping','Garden Center',92,7.36,'DIY','Paid','Landscaping',false,'Tall fescue seed blend + starter fertilizer'),
  E('PROP-001','10/4/2025','Pest Prevention','Exterior','REP-0006','','Pest exclusion — foam sealant','Parts & Materials','Hardware Store',12,0.96,'DIY','Paid','Parts & Materials',false,'Great Stuff foam + caulk for pipe gaps'),
  E('PROP-001','10/5/2025','Landscaping','Yard / Landscaping','','AST-0026','Snow blower service — oil and parts','Parts & Materials','Hardware Store',22,1.76,'DIY','Paid','Parts & Materials',false,'2-cycle oil + spark plug'),
  E('PROP-001','10/6/2025','Interior','Garage','','AST-0018','Garage door lubricant — silicone spray','Parts & Materials','Hardware Store',5,0.40,'DIY','Paid','Parts & Materials',false,'WD-40 Specialist silicone spray'),
  E('PROP-001','10/11/2025','Cleaning','Kitchen','','','Kitchen deep clean — supplies','Cleaning','Self',0,0,'DIY','Paid','Cleaning',false,'Degreaser + scrub pads (household supplies)'),
  E('PROP-001','10/14/2025','Safety','Safety','','AST-0012','Smoke detector batteries (x6 9V)','Parts & Materials','Hardware Store',14,1.12,'DIY','Paid','Parts & Materials',false,'Duracell 9V CopperTop'),
  E('PROP-001','10/16/2025','Landscaping','Yard / Landscaping','','AST-0023','Irrigation winterization — blow-out','Contractor Service','GreenRun Irrigation',80,0,'Professional','Paid','Contractor Service',false,'Compressor blow-out, all 6 zones'),
  E('PROP-001','10/18/2025','Pest Prevention','Exterior','','','Pest control — fall perimeter treatment','Contractor Service','HomeShield Pest Control',150,0,'Professional','Paid','Contractor Service',false,'Fall perimeter spray + attic gap sealed'),
  E('PROP-001','10/25/2025','HVAC','Laundry','','AST-0009','Dryer vent brush kit','Parts & Materials','Hardware Store',0,0,'DIY','Paid','Parts & Materials',false,'Already owned — no new purchase'),
  E('PROP-001','11/1/2025','HVAC','HVAC','','AST-0001','Furnace tune-up — labor 2025','Contractor Service','AirPro HVAC LLC',180,0,'Professional','Paid','Contractor Service',false,'Annual furnace service including ignitor replacement'),
  E('PROP-001','11/1/2025','HVAC','HVAC','','AST-0001','Furnace ignitor — parts 2025','Parts & Materials','AirPro HVAC LLC',45,3.60,'Professional','Paid','Parts & Materials',false,'Hot surface ignitor — replaced during tune-up'),
  // ── 2026 costs ─────────────────────────────────────────────────────────────
  E('PROP-001','1/1/2026','HVAC','HVAC','','AST-0001','HVAC air filter — January 2026','Parts & Materials','Hardware Store',5,0.40,'DIY','Paid','Preventive Maintenance',false,'1-inch MERV 8 filter'),
  E('PROP-001','1/8/2026','Water Systems','Plumbing','','AST-0027','Water softener salt — January 2026','Parts & Materials','Hardware Store',20,1.60,'DIY','Paid','Preventive Maintenance',false,'40 lb Morton Clean & Protect'),
  E('PROP-001','1/8/2026','Appliances','Laundry','','AST-0008','Washer drum-clean tablet','Parts & Materials','Hardware Store',3,0.24,'DIY','Paid','Preventive Maintenance',false,'Affresh tablet — 3-pack'),
  E('PROP-001','2/1/2026','HVAC','HVAC','','AST-0001','HVAC air filter — February 2026','Parts & Materials','Hardware Store',5,0.40,'DIY','Paid','Preventive Maintenance',false,'1-inch MERV 8 filter'),
  E('PROP-001','2/1/2026','Safety','Safety','REP-0021','AST-0013','CO detector replacements (x3) — ordered','Replacement','Online Retailer',60,4.80,'DIY','Paid','Replacement',false,'Kidde 3-pack CO detectors'),
  E('PROP-001','3/1/2026','HVAC','HVAC','','AST-0001','HVAC air filter — March 2026','Parts & Materials','Hardware Store',5,0.40,'DIY','Paid','Preventive Maintenance',false,'1-inch MERV 8 filter'),
  E('PROP-001','4/1/2026','HVAC','HVAC','','AST-0001','HVAC air filter — April 2026','Parts & Materials','Hardware Store',5,0.40,'DIY','Paid','Preventive Maintenance',false,'1-inch MERV 8 filter'),
  // ── Planned / Estimated future costs ──────────────────────────────────────
  E('PROP-001','5/1/2026','HVAC','HVAC','','AST-0002','AC tune-up — spring 2026 (estimate)','Contractor Service','AirPro HVAC LLC',165,0,'Professional','Planned','Contractor Service',false,'Annual AC service estimate'),
  E('PROP-001','4/15/2026','Pest Prevention','Exterior','','','Pest control — spring 2026 (estimate)','Contractor Service','HomeShield Pest Control',155,0,'Professional','Planned','Contractor Service',false,'Spring perimeter treatment estimate'),
  E('PROP-001','4/1/2026','Landscaping','Yard / Landscaping','','','Pre-emergent weed control — 2026 (estimate)','Preventive Maintenance','Garden Center',40,3.20,'DIY','Planned','Preventive Maintenance',false,'Annual pre-emergent estimate'),
  E('PROP-001','5/1/2026','Cleaning','HVAC','','','Air purifier filters — HEPA + carbon (estimate)','Preventive Maintenance','Online Retailer',80,6.40,'DIY','Planned','Preventive Maintenance',false,'Replacement filter set estimate'),
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title row ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 13, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['Maintenance Costs']] });

  // ── Subtitle row ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Track every maintenance and repair expense — see annual totals and cost trends on the Dashboard']] });

  // ── KPI cards ─────────────────────────────────────────────────────────────────
  const kpiDefs = [
    { label: 'Total Paid (All Time)', col: 0, span: 4, color: C.primary },
    { label: `Paid in ${new Date().getFullYear()}`, col: 5, span: 3, color: C.secondary },
    { label: 'DIY Savings (No Labor)', col: 9, span: 4, color: C.success },
    { label: 'Total Planned', col: 14, span: 3, color: C.warning },
    { label: 'Reimbursable', col: 18, span: 2, color: C.info },
  ];

  const D1 = 6;
  const DE = 1000;

  for (const k of kpiDefs) {
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(k.color), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }

  const kpiFormulas = [
    { col: 'A', formula: `=SUMPRODUCT((O${D1+1}:O${DE}="Paid")*(ISNUMBER(M${D1+1}:M${DE}))*M${D1+1}:M${DE})` },
    { col: 'F', formula: `=SUMPRODUCT((O${D1+1}:O${DE}="Paid")*(P${D1+1}:P${DE}=${new Date().getFullYear()})*(ISNUMBER(M${D1+1}:M${DE}))*M${D1+1}:M${DE})` },
    { col: 'J', formula: `=SUMPRODUCT((N${D1+1}:N${DE}="DIY")*(ISNUMBER(M${D1+1}:M${DE}))*M${D1+1}:M${DE})` },
    { col: 'O', formula: `=SUMPRODUCT((O${D1+1}:O${DE}="Planned")*(ISNUMBER(M${D1+1}:M${DE}))*M${D1+1}:M${DE})` },
    { col: 'S', formula: `=SUMPRODUCT((S${D1+1}:S${DE}=TRUE)*(ISNUMBER(M${D1+1}:M${DE}))*M${D1+1}:M${DE})` },
  ];
  for (const { col, formula } of kpiFormulas) {
    vals.push({ range: `${S}!${col}4`, values: [[formula]] });
  }

  for (const [c1, c2] of [[0,4],[5,8],[9,13],[14,17],[18,20]]) {
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, c1, c2), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  for (const [ri, px] of [[0,34],[1,24],[2,22],[3,42]]) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  }

  // ── Divider ───────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ── Header row ────────────────────────────────────────────────────────────────
  const HEADERS = [
    'Cost ID','Property ID','Date','Category','Home Area','Task / Repair Ref','Asset ID','Description',
    'Type','Vendor / Provider','Amount ($)','Tax / Fees ($)','Total (fml)','DIY or Pro','Payment Status',
    'Year (fml)','Month (fml)','Budget Category','Reimbursable?','Notes'
  ];
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A6:T6`, values: [HEADERS] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Data rows ────────────────────────────────────────────────────────────────
  const DR = COSTS.length;
  const DS = D1;

  const dataRows = [];
  COSTS.forEach((row, ri) => {
    const r = ri + D1 + 1;
    row[0] = `=IF(B${r}="","","CST-"&TEXT(ROW()-6,"0000"))`;
    row[12] = `=IF(AND(K${r}="",L${r}=""),"",IFERROR(VALUE(K${r}),0)+IFERROR(VALUE(L${r}),0))`;
    row[15] = `=IF(C${r}<>"",YEAR(C${r}),"")`;
    row[16] = `=IF(C${r}<>"",MONTH(C${r}),"")`;
    dataRows.push(row);
  });

  vals.push({ range: `${S}!A7:T${6+DR}`, values: dataRows });

  // ── Data formatting ────────────────────────────────────────────────────────────
  for (let ri = 0; ri < DR; ri++) {
    const rowIdx = DS + ri;
    const bg = ri % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Formula cols: A(0), M(12), P(15), Q(16)
  for (const ci of [0, 12, 15, 16]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Input bg
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 1, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 13, 15), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 17, 20), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Date format: C(2)
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 2, 3), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Currency: K(10), L(11), M(12)
  for (const ci of [10, 11, 12]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // Checkbox: S(18) Reimbursable
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 18, 19), cell: { dataValidation: { condition: { type: 'BOOLEAN' }, showCustomUi: true } }, fields: 'dataValidation' } });

  // Wrap notes (T=19)
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 19, 20), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat.wrapStrategy' } });

  // ── Data validations ─────────────────────────────────────────────────────────
  const valDefs = [
    [3, ['HVAC','Plumbing','Electrical','Roof & Gutters','Exterior','Interior','Appliances','Safety','Landscaping','Cleaning','Pest Prevention','Water Systems','Energy Efficiency','Seasonal Preparation','General Inspection','Other']],
    [4, ['Whole Home','Kitchen','Bathroom','Bedroom','Living Room','Dining Room','Laundry','Basement','Attic','Garage','Exterior','Roof','Yard / Landscaping','Driveway','Patio / Deck','HVAC','Plumbing','Electrical','Safety','Other']],
    [8, ['Preventive Maintenance','Repair Labor','Parts & Materials','Contractor Service','Inspection','Replacement','Permit / Fee','Tool / Equipment','Landscaping','Cleaning','Other']],
    [13, ['DIY','Professional','Mixed','Not Decided']],
    [14, ['Planned','Estimate Received','Due','Paid','Partially Paid','Refunded','Cancelled']],
    [17, ['Preventive Maintenance','Repair Labor','Parts & Materials','Contractor Service','Inspection','Replacement','Permit / Fee','Tool / Equipment','Landscaping','Cleaning','Other']],
  ];
  for (const [col, list] of valDefs) {
    fmt.push({ setDataValidation: { range: gridRange(SID, DS, DS+DR, col, col+1), rule: { condition: { type: 'ONE_OF_LIST', values: list.map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: true } } });
  }

  // ── Row heights ───────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DS, endIndex: DS+DR }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const colWidths = [75,70,90,110,100,90,80,220,120,150,90,80,85,80,110,55,55,130,80,160];
  colWidths.forEach((px, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'costs-fmt');
  await valuesBatchUpdate(id, vals, 'costs-vals');

  console.log(`✓ Maintenance Costs complete — ${DR} cost records written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
