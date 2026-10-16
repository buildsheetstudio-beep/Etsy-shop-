'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Maintenance Task Log'];
const S = "'Maintenance Task Log'";
const REC = "'Recurring Maintenance'";
const AST = "'Appliances & Systems'";
const NC = 24; // A-X

// Col mapping:
// A(0)=Task ID   B(1)=Property ID    C(2)=Task Name       D(3)=Category
// E(4)=Home Area  F(5)=Asset ID       G(6)=Priority        H(7)=Status
// I(8)=Season     J(9)=Scheduled Date K(10)=Due Date       L(11)=Completed Date
// M(12)=DIY or Pro N(13)=Provider     O(14)=Template ID    P(15)=Est Duration Hr
// Q(16)=Est Cost   R(17)=Actual Cost  S(18)=Notes          T(19)=Overdue(fml)
// U(20)=Days Until Due(fml)           V(21)=Month(fml)     W(22)=Year(fml)
// X(23)=Cost Logged(fml)

function colL(i) {
  if (i < 26) return String.fromCharCode(65 + i);
  return String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26));
}

const T = (propId, name, cat, area, assetId, pri, status, season, scheduledDate, dueDate, completedDate, diyPro, provider, templateId, estHr, estCost, actualCost, notes) => {
  const row = Array(NC).fill('');
  row[1] = propId;
  row[2] = name;
  row[3] = cat;
  row[4] = area;
  row[5] = assetId;
  row[6] = pri;
  row[7] = status;
  row[8] = season;
  row[9] = scheduledDate;
  row[10] = dueDate;
  row[11] = completedDate;
  row[12] = diyPro;
  row[13] = provider;
  row[14] = templateId;
  row[15] = estHr;
  row[16] = estCost;
  row[17] = actualCost;
  row[18] = notes;
  // cols T-X are formulas, set per-row below
  return row;
};

const TASKS = [
  // ── Completed 2025 tasks ─────────────────────────────────────────────────────
  T('PROP-001','Replace HVAC Air Filter — January','HVAC','HVAC','AST-0001','High','Complete','Winter','1/1/2026','1/15/2026','1/15/2026','DIY','Self','REC-001',0.25,5,5,'1-inch filter replaced. MERV 8.'),
  T('PROP-001','Replace HVAC Air Filter — February','HVAC','HVAC','AST-0001','High','Complete','Winter','2/1/2026','2/15/2026','2/14/2026','DIY','Self','REC-001',0.25,5,5,''),
  T('PROP-001','Replace HVAC Air Filter — March','HVAC','HVAC','AST-0001','High','Complete','Spring','3/1/2026','3/15/2026','3/15/2026','DIY','Self','REC-001',0.25,5,5,''),
  T('PROP-001','Replace HVAC Air Filter — April','HVAC','HVAC','AST-0001','High','Complete','Spring','4/1/2026','4/15/2026','4/12/2026','DIY','Self','REC-001',0.25,5,5,''),
  T('PROP-001','Central AC Tune-Up — Spring 2025','HVAC','HVAC','AST-0002','High','Complete','Spring','5/1/2025','5/15/2025','5/15/2025','Professional','AirPro HVAC LLC','REC-003',2,160,160,'Refrigerant check OK. Coils cleaned. Drain flushed.'),
  T('PROP-001','Clean AC Condensate Drain Line','HVAC','HVAC','AST-0002','Medium','Complete','Spring','5/1/2025','5/1/2025','5/3/2025','DIY','Self','REC-004',0.5,0,0,'Flushed with 1 cup diluted bleach solution.'),
  T('PROP-001','Test Sump Pump — Spring 2025','Plumbing','Basement','AST-0011','High','Complete','Spring','4/1/2025','4/1/2025','4/2/2025','DIY','Self','REC-007',0.5,0,0,'Float switch tested. Discharge line clear.'),
  T('PROP-001','Test Pressure Relief Valve — Water Heater','Plumbing','Plumbing','AST-0003','Medium','Complete','Spring','4/1/2025','4/1/2025','4/3/2025','DIY','Self','REC-010',0.25,0,0,'Valve released water normally. Reseated OK.'),
  T('PROP-001','Inspect All Plumbing Supply Lines','Plumbing','Whole Home','','Medium','Complete','Spring','4/1/2025','4/1/2025','4/5/2025','DIY','Self','REC-011',0.5,0,0,'All braided hoses in good condition. No bulging found.'),
  T('PROP-001','Test GFCI & AFCI Outlets — Spring 2025','Electrical','Electrical','','High','Complete','Spring','4/1/2025','4/1/2025','4/5/2025','DIY','Self','REC-016',0.5,0,0,'All 8 GFCI outlets tested. 2 AFCI breakers in panel tested OK.'),
  T('PROP-001','Clean & Inspect Gutters — Spring 2025','Roof & Gutters','Exterior','AST-0017','High','Complete','Spring','4/15/2025','4/15/2025','4/19/2025','DIY','Self','REC-023',2,0,0,'Removed winter debris. Downspouts flushed. No sags found.'),
  T('PROP-001','Roof Visual Inspection — Spring 2025','Roof & Gutters','Roof','AST-0016','Medium','Complete','Spring','4/15/2025','4/15/2025','4/20/2025','DIY','Self','REC-025',0.5,0,0,'No missing shingles. Minor granule loss in south gutter — note for monitoring.'),
  T('PROP-001','Inspect & Caulk Window & Door Seals','Exterior','Exterior','','Medium','Complete','Spring','4/15/2025','4/15/2025','4/26/2025','DIY','Self','REC-027',2,20,18,'Recaulked 3 windows and rear door threshold.'),
  T('PROP-001','Activate Irrigation System — Spring 2025','Landscaping','Yard / Landscaping','AST-0023','Medium','Complete','Spring','4/15/2025','4/15/2025','4/16/2025','DIY','Self','REC-040',1,0,0,'Zone 3 head replaced — cracked over winter.'),
  T('PROP-001','Service Lawn Mower — Spring 2025','Landscaping','Yard / Landscaping','AST-0025','Medium','Complete','Spring','4/1/2025','4/1/2025','4/6/2025','DIY','Self','REC-042',1.5,30,38,'Oil changed, blade sharpened, air filter replaced.'),
  T('PROP-001','Apply Pre-Emergent Weed Control — 2025','Landscaping','Yard / Landscaping','','Medium','Complete','Spring','4/1/2025','4/1/2025','4/3/2025','DIY','Self','REC-046',1,40,40,'Applied Scotts Halts per label. Soil temp 52°F.'),
  T('PROP-001','Apply Lawn Fertilizer — Spring 2025','Landscaping','Yard / Landscaping','','Medium','Complete','Spring','4/15/2025','4/15/2025','4/18/2025','DIY','Self','REC-045',1,60,55,'Milorganite slow-release applied after soil temp 57°F.'),
  T('PROP-001','Inspect Exterior Doors & Weatherstripping','Exterior','Exterior','AST-0022','Low','Complete','Spring','4/15/2025','4/15/2025','4/26/2025','DIY','Self','',1,0,0,'Replaced rear door bottom sweep. Front door good.'),
  T('PROP-001','Inspect & Clean Bathroom Exhaust Fans','Interior','Bathroom','','Medium','Complete','Spring','4/15/2025','4/15/2025','4/27/2025','DIY','Self','REC-034',1,0,0,'3 fans cleaned. Master bath fan running at reduced speed — may need replacement.'),
  T('PROP-001','Clean Refrigerator Condenser Coils','Appliances','Kitchen','AST-0004','Medium','Complete','Spring','4/15/2025','4/15/2025','4/27/2025','DIY','Self','REC-035',0.5,0,0,'Pulled fridge 12 inches, vacuumed coils. Running noticeably quieter.'),
  T('PROP-001','Inspect Washer Hoses for Cracks','Appliances','Laundry','AST-0008','High','Complete','Spring','4/1/2025','4/1/2025','4/5/2025','DIY','Self','REC-038',0.25,0,0,'Braided stainless hoses in good condition. No action needed.'),
  T('PROP-001','Power Wash Siding & Driveway — 2025','Exterior','Exterior','','Low','Complete','Spring','4/15/2025','4/15/2025','4/29/2025','DIY','Self','REC-030',3,0,0,'Used extension wand. Minor efflorescence on east foundation wall noted.'),
  T('PROP-001','Inspect Foundation & Grading — 2025','Exterior','Exterior','','Medium','Complete','Spring','4/15/2025','4/15/2025','4/29/2025','DIY','Self','REC-029',1,0,0,'No foundation cracks found. Grading adequate on all sides.'),
  T('PROP-001','Pest Control Perimeter Treatment — Spring','Pest Prevention','Exterior','','Medium','Complete','Spring','4/15/2025','4/15/2025','4/22/2025','Professional','HomeShield Pest Control','REC-049',1,150,150,'Exterior spray applied. Reapply in October.'),
  T('PROP-001','Mulch Garden Beds — Spring 2025','Landscaping','Yard / Landscaping','','Low','Complete','Spring','5/1/2025','5/1/2025','5/4/2025','DIY','Self','REC-047',3,60,72,'3 cubic yards brown hardwood mulch applied.'),
  T('PROP-001','Furnace Tune-Up & Safety Check — Fall 2025','HVAC','HVAC','AST-0001','High','Complete','Fall','11/1/2025','11/1/2025','11/1/2025','Professional','AirPro HVAC LLC','REC-002',2,180,180,'Burners cleaned, heat exchanger clear, flue tested. New ignitor installed — $45 parts.'),
  T('PROP-001','Inspect & Clean Dryer Vent Duct — 2025','HVAC','Laundry','AST-0009','High','Complete','Fall','10/1/2025','10/1/2025','10/4/2025','DIY','Self','REC-005',1,0,0,'Removed substantial lint accumulation. Duct runs 18 feet — consider professional cleaning next year.'),
  T('PROP-001','Replace Smoke Detector Batteries — 2025','Safety','Safety','AST-0012','High','Complete','Fall','10/1/2025','10/1/2025','10/5/2025','DIY','Self','REC-019',0.5,15,14,'Replaced 6 units with Duracell 9V.'),
  T('PROP-001','Test Sump Pump — Fall 2025','Plumbing','Basement','AST-0011','High','Complete','Fall','10/1/2025','10/1/2025','10/3/2025','DIY','Self','REC-007',0.5,0,0,'Float switch tested. Battery backup unit tested OK.'),
  T('PROP-001','Winterize Hose Bibs — Fall 2025','Plumbing','Exterior','','High','Complete','Fall','11/1/2025','11/1/2025','11/2/2025','DIY','Self','REC-012',0.5,0,0,'Interior valves closed on 3 spigots. Lines drained.'),
  T('PROP-001','Clean & Inspect Gutters — Fall 2025','Roof & Gutters','Exterior','AST-0017','High','Complete','Fall','11/1/2025','11/1/2025','11/8/2025','DIY','Self','REC-024',2,0,0,'Cleared heavy leaf fall. Downspouts clear. One gutter hanger re-fastened.'),
  T('PROP-001','Test GFCI & AFCI Outlets — Fall 2025','Electrical','Electrical','','High','Complete','Fall','10/1/2025','10/1/2025','10/5/2025','DIY','Self','REC-016',0.5,0,0,'All outlets and breakers tested. All pass.'),
  T('PROP-001','Lubricate Garage Door Springs & Rollers — Fall','Interior','Garage','AST-0018','Medium','Complete','Fall','10/1/2025','10/1/2025','10/6/2025','DIY','Self','REC-032',0.5,5,5,'Silicone spray applied. Both doors operating smoothly.'),
  T('PROP-001','Pest Control — Fall 2025','Pest Prevention','Exterior','','Medium','Complete','Fall','10/15/2025','10/15/2025','10/18/2025','Professional','HomeShield Pest Control','REC-049',1,150,150,'Exterior spray + attic entrance gap sealed.'),
  T('PROP-001','Inspect & Seal Pest Entry Points — 2025','Pest Prevention','Exterior','','Medium','Complete','Fall','10/1/2025','10/1/2025','10/10/2025','DIY','Self','REC-050',1,15,12,'Foam-sealed 2 pipe penetrations and dryer vent gap.'),
  T('PROP-001','Winterize Irrigation System — Fall 2025','Landscaping','Yard / Landscaping','AST-0023','High','Complete','Fall','10/15/2025','10/15/2025','10/16/2025','Professional','GreenRun Irrigation','REC-041',1,80,80,'All 6 zones blown out. Controller set to off.'),
  T('PROP-001','Service Snow Blower — Pre-Season 2025','Landscaping','Yard / Landscaping','AST-0026','Medium','Complete','Fall','11/1/2025','11/1/2025','11/5/2025','DIY','Self','REC-043',1,25,22,'Oil changed, shear pins checked, auger lubed.'),
  T('PROP-001','Dethatch & Overseed Lawn — Fall 2025','Landscaping','Yard / Landscaping','','Medium','Complete','Fall','9/1/2025','9/1/2025','9/6/2025','DIY','Self','REC-044',4,80,92,'Dethatched with rental, overseeded with tall fescue blend.'),
  T('PROP-001','Flush Water Heater — Fall 2025','Plumbing','Plumbing','AST-0003','Medium','Complete','Fall','9/1/2025','9/1/2025','9/4/2025','DIY','Self','REC-008',1,0,0,'Flushed until clear. About 2 gallons of sediment removed.'),
  T('PROP-001','Inspect Attic for Moisture / Pests','Interior','Attic','','Medium','Complete','Summer','7/1/2025','7/1/2025','7/5/2025','DIY','Self','REC-031',0.5,0,0,'No moisture stains, no pest evidence. Soffit vents clear.'),
  T('PROP-001','Inspect Electrical Panel — Visual 2025','Electrical','Electrical','AST-0015','Medium','Complete','Summer','7/1/2025','7/1/2025','7/5/2025','DIY','Self','REC-020',0.25,0,0,'No tripped breakers, no heat marks. All labels current.'),
  T('PROP-001','Touch-Up Exterior Paint — Trim 2025','Exterior','Exterior','','Low','Complete','Summer','7/1/2025','7/1/2025','7/12/2025','DIY','Self','REC-028',3,40,45,'Primed and painted east trim, front porch rail post.'),
  T('PROP-001','Deep Clean Kitchen — Fall 2025','Cleaning','Kitchen','','Low','Complete','Fall','10/1/2025','10/1/2025','10/11/2025','DIY','Self','REC-052',3,0,0,'Hood filter degreased, backsplash scrubbed, cabinet fronts wiped.'),
  T('PROP-001','Annual Home Energy Audit — 2025','Energy Efficiency','Whole Home','','Low','Complete','Spring','3/1/2025','3/1/2025','3/14/2025','Professional','UtilityFirst Energy','REC-053',3,0,0,'Blower door test: ACH50 = 4.2. Recommendations: attic air seal over top plate, add weatherstrip to garage door.'),
  T('PROP-001','Inspect & Re-Caulk Deck — Spring 2025','Exterior','Patio / Deck','','Medium','Complete','Spring','4/15/2025','4/15/2025','4/30/2025','DIY','Self','REC-054',2,20,16,'Recaulked ledger board seam and 2 post bases. Board clips all secure.'),

  // ── Current / In Progress tasks (2026) ───────────────────────────────────────
  T('PROP-001','Replace HVAC Air Filter — May','HVAC','HVAC','AST-0001','High','Scheduled','Spring','5/1/2026','5/15/2026','','DIY','Self','REC-001',0.25,5,0,''),
  T('PROP-001','Central AC Tune-Up — Spring 2026','HVAC','HVAC','AST-0002','High','Scheduled','Spring','5/1/2026','5/15/2026','','Professional','AirPro HVAC LLC','REC-003',2,165,0,'Book appointment in April.'),
  T('PROP-001','Clean & Inspect Gutters — Spring 2026','Roof & Gutters','Exterior','AST-0017','High','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-023',2,0,0,''),
  T('PROP-001','Roof Visual Inspection — Spring 2026','Roof & Gutters','Roof','AST-0016','Medium','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-025',0.5,0,0,''),
  T('PROP-001','Activate Irrigation System — Spring 2026','Landscaping','Yard / Landscaping','AST-0023','Medium','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-040',1,0,0,''),
  T('PROP-001','Apply Pre-Emergent Weed Control — 2026','Landscaping','Yard / Landscaping','','Medium','Not Started','Spring','4/1/2026','4/1/2026','','DIY','Self','REC-046',1,40,0,''),
  T('PROP-001','Apply Lawn Fertilizer — Spring 2026','Landscaping','Yard / Landscaping','','Medium','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-045',1,60,0,''),
  T('PROP-001','Test GFCI & AFCI Outlets — Spring 2026','Electrical','Electrical','','High','Not Started','Spring','4/1/2026','4/1/2026','','DIY','Self','REC-016',0.5,0,0,''),
  T('PROP-001','Test Sump Pump — Spring 2026','Plumbing','Basement','AST-0011','High','Not Started','Spring','4/1/2026','4/1/2026','','DIY','Self','REC-007',0.5,0,0,''),
  T('PROP-001','Inspect All Plumbing Supply Lines — 2026','Plumbing','Whole Home','','Medium','Not Started','Spring','4/1/2026','4/1/2026','','DIY','Self','REC-011',0.5,0,0,''),
  T('PROP-001','Pest Control Perimeter Treatment — Spring 2026','Pest Prevention','Exterior','','Medium','Scheduled','Spring','4/15/2026','4/15/2026','','Professional','HomeShield Pest Control','REC-049',1,155,0,'Appointment booked.'),
  T('PROP-001','Inspect Fire Extinguishers — 2026','Safety','Safety','AST-0014','High','Not Started','Winter','1/15/2026','1/15/2026','','DIY','Self','REC-021',0.25,0,0,'Gauge in green; pin and tamper seal intact.'),
  T('PROP-001','Test Smoke Detectors — January 2026','Safety','Safety','AST-0012','Urgent','Complete','Winter','1/1/2026','1/15/2026','1/6/2026','DIY','Self','REC-017',0.25,0,0,'All 6 units tested OK.'),
  T('PROP-001','Test Carbon Monoxide Detectors — January 2026','Safety','Safety','AST-0013','Urgent','Complete','Winter','1/1/2026','1/15/2026','1/6/2026','DIY','Self','REC-018',0.1,0,0,'3 units tested OK. Note: approaching 7-year life.'),
  T('PROP-001','Replace Carbon Monoxide Detectors','Safety','Safety','AST-0013','High','In Progress','Winter','2/1/2026','3/1/2026','','Professional','Self','',0.5,60,0,'Units approaching 7-year rated life. Ordered replacement 3-pack.'),
  T('PROP-001','Run Water Softener Salt Check — Jan 2026','Water Systems','Plumbing','AST-0027','Medium','Complete','Winter','1/1/2026','1/15/2026','1/8/2026','DIY','Self','REC-015',0.1,20,20,'Added one 40-lb bag Morton Clean & Protect.'),
  T('PROP-001','Run Washing Machine Drum-Clean — Jan 2026','Appliances','Laundry','AST-0008','Low','Complete','Winter','1/1/2026','1/15/2026','1/8/2026','DIY','Self','REC-037',0.5,3,3,'Affresh tablet used.'),
  T('PROP-001','Inspect & Clean Deck — Spring 2026','Exterior','Patio / Deck','','Medium','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-054',2,20,0,''),
  T('PROP-001','Lubricate Garage Door — Spring 2026','Interior','Garage','AST-0018','Medium','Not Started','Spring','4/1/2026','4/1/2026','','DIY','Self','REC-032',0.5,5,0,''),
  T('PROP-001','Test Garage Door Auto-Reverse Safety','Interior','Garage','AST-0018','High','Not Started','Spring','4/1/2026','4/1/2026','','DIY','Self','REC-033',0.25,0,0,''),
  T('PROP-001','Power Wash Siding & Driveway — 2026','Exterior','Exterior','','Low','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-030',3,0,0,''),
  T('PROP-001','Inspect Foundation & Grading — 2026','Exterior','Exterior','','Medium','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-029',1,0,0,'Monitor east wall efflorescence noted in 2025.'),
  T('PROP-001','Inspect & Caulk Window & Door Seals — 2026','Exterior','Exterior','','Medium','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-027',2,20,0,''),
  T('PROP-001','Clean Dishwasher Filter & Spray Arms','Appliances','Kitchen','AST-0005','Low','Scheduled','Spring','4/1/2026','4/1/2026','','DIY','Self','REC-036',0.25,0,0,''),
  T('PROP-001','Change Whole-Home Air Purifier Filters','Cleaning','HVAC','','Medium','Scheduled','Spring','5/1/2026','5/1/2026','','DIY','Self','REC-051',0.25,80,0,'Filters ordered. HEPA + carbon set.'),
  T('PROP-001','Seasonal Walk-Through — Spring 2026','General Inspection','Whole Home','','Medium','Not Started','Spring','4/15/2026','4/15/2026','','DIY','Self','REC-055',2,0,0,'Annual spring systems check.'),
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title row ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 13, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['Maintenance Task Log']] });

  // ── Subtitle row ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Every maintenance task and its completion history — search, filter, and track by status or category']] });

  // ── KPI card labels row (row 2, 0-indexed) ───────────────────────────────────
  const kpiDefs = [
    { label: 'Total Tasks', col: 0, span: 3, color: C.primary },
    { label: 'Completed', col: 4, span: 3, color: C.success },
    { label: 'In Progress', col: 8, span: 3, color: C.info },
    { label: 'Overdue', col: 12, span: 3, color: C.attention },
    { label: 'Total Spent', col: 17, span: 4, color: C.secondary },
    { label: '% Complete', col: 21, span: 3, color: C.wheat },
  ];
  for (const k of kpiDefs) {
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(k.color), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${colL(k.col)}3`, values: [[k.label]] });
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }

  const D1 = 6; // 1-indexed first data row
  const DE = 1000; // last data row for formulas

  const kpiFormulas = [
    { col: 'A', formula: `=COUNTA(A${D1+1}:A${DE})` },
    { col: 'E', formula: `=SUMPRODUCT((H${D1+1}:H${DE}="Complete")*1)` },
    { col: 'I', formula: `=SUMPRODUCT((H${D1+1}:H${DE}="In Progress")*1)+SUMPRODUCT((H${D1+1}:H${DE}="Scheduled")*1)` },
    { col: 'M', formula: `=SUMPRODUCT((T${D1+1}:T${DE}=TRUE)*1)` },
    { col: 'R', formula: `=SUMPRODUCT((ISNUMBER(R${D1+1}:R${DE}))*R${D1+1}:R${DE})` },
    { col: 'V', formula: `=IFERROR(SUMPRODUCT((H${D1+1}:H${DE}="Complete")*1)/COUNTA(A${D1+1}:A${DE}),0)` },
  ];
  for (const { col, formula } of kpiFormulas) {
    vals.push({ range: `${S}!${col}4`, values: [[formula]] });
  }

  // KPI value number formats
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 3), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 4, 7), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 8, 11), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 12, 15), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 17, 21), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 21, 24), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Row heights for top rows
  for (const [ri, px] of [[0,34],[1,24],[2,22],[3,42]]) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  }

  // ── Divider row (row 4, 0-indexed) ───────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ── Header row (row 5, 0-indexed) ────────────────────────────────────────────
  const HEADERS = [
    'Task ID','Property ID','Task Name','Category','Home Area','Asset ID','Priority','Status',
    'Season','Scheduled Date','Due Date','Completed Date','DIY or Pro','Provider','Template ID',
    'Est. Duration (hr)','Est. Cost ($)','Actual Cost ($)','Notes','Overdue?','Days Until Due',
    'Month','Year','Cost Logged?'
  ];
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A6:X6`, values: [HEADERS] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Data rows ────────────────────────────────────────────────────────────────
  const dataRows = [];
  TASKS.forEach((row, ri) => {
    const r = ri + D1 + 1; // 1-indexed row
    // A(0) = Task ID
    row[0] = `=IF(B${r}="","","TSK-"&TEXT(ROW()-6,"0000"))`;
    // T(19) = Overdue: due date in past AND status not Complete/Cancelled/Skipped
    row[19] = `=IF(AND(K${r}<>"",K${r}<TODAY(),H${r}<>"Complete",H${r}<>"Cancelled",H${r}<>"Skipped"),TRUE,FALSE)`;
    // U(20) = Days Until Due
    row[20] = `=IF(K${r}="","",K${r}-TODAY())`;
    // V(21) = Month (for monthly aggregation)
    row[21] = `=IF(L${r}<>"",MONTH(L${r}),IF(K${r}<>"",MONTH(K${r}),""))`;
    // W(22) = Year
    row[22] = `=IF(L${r}<>"",YEAR(L${r}),IF(K${r}<>"",YEAR(K${r}),""))`;
    // X(23) = Cost Logged? (is there an actual cost)
    row[23] = `=IF(R${r}>0,TRUE,FALSE)`;
    dataRows.push(row);
  });

  vals.push({ range: `${S}!A7:X${6+TASKS.length}`, values: dataRows });

  // ── Data formatting ────────────────────────────────────────────────────────────
  const DR = TASKS.length;
  const DS = D1; // 0-indexed data start

  // Alt row banding
  for (let ri = 0; ri < DR; ri++) {
    const rowIdx = DS + ri;
    const bg = ri % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Formula cols: A(0), T(19), U(20), V(21), W(22), X(23)
  for (const ci of [0, 19, 20, 21, 22, 23]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Input bg for user-editable cols (B=1 to S=18)
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 1, 19), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Date number formats: J(9) Scheduled, K(10) Due, L(11) Completed
  for (const ci of [9, 10, 11]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // Currency: Q(16) Est Cost, R(17) Actual Cost
  for (const ci of [16, 17]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // Number: P(15) Est Duration, U(20) Days Until Due
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 15, 16), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 20, 21), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Notes col (S=18) wrap text
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 18, 19), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat.wrapStrategy' } });

  // ── Data validations ─────────────────────────────────────────────────────────
  const valDefs = [
    [3, ['HVAC','Plumbing','Electrical','Roof & Gutters','Exterior','Interior','Appliances','Safety','Landscaping','Cleaning','Pest Prevention','Water Systems','Energy Efficiency','Seasonal Preparation','General Inspection','Other']],
    [4, ['Whole Home','Kitchen','Bathroom','Bedroom','Living Room','Dining Room','Laundry','Basement','Attic','Garage','Exterior','Roof','Yard / Landscaping','Driveway','Patio / Deck','HVAC','Plumbing','Electrical','Safety','Other']],
    [6, ['Low','Medium','High','Urgent']],
    [7, ['Not Started','Scheduled','In Progress','Waiting','Complete','Skipped','Cancelled']],
    [8, ['Winter','Spring','Summer','Fall','Year-Round']],
    [12, ['DIY','Professional','Mixed','Not Decided']],
  ];
  for (const [col, list] of valDefs) {
    fmt.push({ setDataValidation: { range: gridRange(SID, DS, DS+DR, col, col+1), rule: { condition: { type: 'ONE_OF_LIST', values: list.map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: true } } });
  }

  // ── Row heights ───────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DS, endIndex: DS+DR }, properties: { pixelSize: 46 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const colWidths = [75,70,220,120,110,75,70,90,80,90,90,90,80,140,80,80,80,80,220,65,75,55,55,80];
  colWidths.forEach((px, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  // Title row merges full width — no frozenColumnCount
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'task-fmt');
  await valuesBatchUpdate(id, vals, 'task-vals');

  console.log(`✓ Maintenance Task Log complete — ${TASKS.length} tasks written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
