'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Home Repair Log'];
const S = "'Home Repair Log'";
const AST = "'Appliances & Systems'";
const NC = 25; // A-Y

// Col mapping:
// A(0)=Repair ID   B(1)=Property ID   C(2)=Asset ID      D(3)=Repair Name
// E(4)=Category    F(5)=Home Area      G(6)=Type          H(7)=Priority
// I(8)=Status      J(9)=Date Reported  K(10)=Date Started  L(11)=Date Completed
// M(12)=DIY or Pro N(13)=Provider      O(14)=Parts & Materials
// P(15)=Labor Cost  Q(16)=Total Cost(fml)  R(17)=Warranty Claim?
// S(18)=Covered by Warranty?  T(19)=Est Days Active(fml)  U(20)=Year(fml)
// V(21)=Urgency Notes  W(22)=Root Cause    X(23)=Resolution   Y(24)=Prevention Notes

const R = (propId, assetId, name, cat, area, type, pri, status, dateRep, dateStart, dateComp, diyPro, provider, parts, labor, warrantyClm, warrantyCov, urgencyNotes, rootCause, resolution, preventionNotes) => {
  const row = Array(NC).fill('');
  row[1] = propId;
  row[2] = assetId;
  row[3] = name;
  row[4] = cat;
  row[5] = area;
  row[6] = type;
  row[7] = pri;
  row[8] = status;
  row[9] = dateRep;
  row[10] = dateStart;
  row[11] = dateComp;
  row[12] = diyPro;
  row[13] = provider;
  row[14] = parts;
  row[15] = labor;
  // Q(16) = total cost formula — set per-row
  row[17] = warrantyClm;
  row[18] = warrantyCov;
  // T(19) = est days active formula — set per-row
  // U(20) = year formula — set per-row
  row[21] = urgencyNotes;
  row[22] = rootCause;
  row[23] = resolution;
  row[24] = preventionNotes;
  return row;
};

const REPAIRS = [
  // ── Completed repairs ──────────────────────────────────────────────────────────
  R('PROP-001','AST-0003','Water Heater — Replace Thermocouple','Plumbing','Plumbing','Corrective','High','Complete','3/2/2023','3/3/2023','3/3/2023','Professional','ClearFlow Plumbing',45,140,false,false,'Pilot light would not stay lit — no hot water.','Thermocouple failed after ~15 years of use.','Replaced thermocouple. Water heater restored to full operation.','Inspect thermocouple every 5 years as part of annual flush.'),
  R('PROP-001','AST-0001','Furnace — Replace Ignitor','HVAC','HVAC','Corrective','Urgent','Complete','11/1/2025','11/1/2025','11/1/2025','Professional','AirPro HVAC LLC',45,0,false,false,'No heat on cold morning; furnace not igniting. Caught during tune-up.','Hot surface ignitor failed — typical 7-10 year lifespan.','Replaced ignitor during annual tune-up. No additional labor charge.','Document ignitor install date; budget for replacement every 10 years.'),
  R('PROP-001','AST-0023','Irrigation — Replace Zone 3 Head','Landscaping','Yard / Landscaping','Corrective','Medium','Complete','4/16/2025','4/16/2025','4/16/2025','DIY','Self',18,0,false,false,'Zone 3 head cracked over winter — found during spring activation.','Freeze damage to uninsulated head left exposed to frost.','Replaced with matching Rain Bird 5000 series head.','Winterize all heads below grade next fall.'),
  R('PROP-001','','Rear Door — Replace Bottom Sweep','Exterior','Exterior','Corrective','Low','Complete','4/26/2025','4/26/2025','4/26/2025','DIY','Self',14,0,false,false,'Gap at bottom of rear door allowing drafts and insects.','Original sweep worn after 7 years.','Installed new vinyl sweep. Door seals properly.','Check door sweeps annually during spring inspection.'),
  R('PROP-001','AST-0009','Dryer Vent — Clean & Inspect Duct','HVAC','Laundry','Preventive','High','Complete','10/4/2025','10/4/2025','10/4/2025','DIY','Self',0,0,false,false,'Annual dryer vent cleaning — fire prevention.','Lint accumulation over annual cycle (18 feet of duct run).','Removed heavy lint. Duct integrity confirmed. Cable brush kit used.','Consider professional cleaning every other year given long duct run.'),
  R('PROP-001','','Foundation — Seal Pipe Penetration (2 locations)','Plumbing','Exterior','Preventive','Medium','Complete','10/10/2025','10/10/2025','10/10/2025','DIY','Self',12,0,false,false,'Annual pest exclusion inspection revealed 2 unsealed pipe penetrations.','Original construction gaps never sealed.','Foam-sealed both penetrations plus dryer vent gap.','Re-inspect annually each fall as part of pest exclusion.'),
  R('PROP-001','AST-0002','AC — Diagnose & Flush Condensate Drain','HVAC','HVAC','Corrective','Medium','Complete','8/12/2023','8/12/2023','8/12/2023','DIY','Self',0,0,false,false,'Drain pan beginning to fill — minor water on mechanical room floor.','Algae growth partially blocking condensate drain line.','Flushed with bleach solution. Drain running freely.','Add condensate drain flush to spring AC startup checklist.'),
  R('PROP-001','AST-0018','Garage Door — Replace Rear Spring','Interior','Garage','Corrective','Urgent','Complete','2/14/2024','2/15/2024','2/15/2024','Professional','Ridgeline Home Services',95,165,false,false,'Left garage door dropped suddenly — spring broke overnight.','Torsion spring failure after ~12 years (typical 10-15k cycle life).','Both torsion springs replaced as pair. Door balanced and adjusted.','Replace springs as a pair. Budget every 10-12 years.'),
  R('PROP-001','','Basement — Seal Efflorescence / Crack (East Wall)','Exterior','Basement','Corrective','Medium','Complete','6/5/2024','6/10/2024','6/12/2024','DIY','Self',35,0,false,false,'White mineral deposits and hairline crack on east foundation wall.','Minor efflorescence from water wicking through block — grading issue.','Applied hydraulic cement to crack. Regraded soil 6 inches per 10 ft along east wall.','Monitor during spring thaw. Re-inspect grading annually.'),
  R('PROP-001','AST-0003','Water Heater — Anode Rod Inspection','Plumbing','Plumbing','Preventive','Medium','Complete','9/4/2025','9/4/2025','9/4/2025','DIY','Self',0,0,false,false,'Anode rod inspection during annual flush — 6-year interval.','Routine inspection — rod was 60% depleted.','Rod has 2-3 years remaining. Will inspect again at next flush.','Plan anode rod replacement for 2027 annual flush.'),
  R('PROP-001','AST-0004','Refrigerator — Condenser Fan Rattling','Appliances','Kitchen','Corrective','Medium','Complete','7/20/2024','7/22/2024','7/22/2024','Professional','AppliancePro Repair',35,95,false,false,'Intermittent rattling from back of refrigerator — audible in kitchen.','Leaf of debris caught in condenser fan blade.','Cleaned fan area. Tightened fan mount. Rattle resolved.','Clean condenser area annually; do not store items behind fridge.'),
  R('PROP-001','','Bathroom — Caulk Bathtub Surround','Interior','Bathroom','Corrective','Medium','Complete','5/18/2024','5/18/2024','5/18/2024','DIY','Self',8,0,false,false,'Master bath tub surround caulk cracking and pulling away — moisture risk.','Caulk dried out and pulled from tile after ~6 years.','Removed old caulk, cleaned, applied new 100% silicone caulk.','Re-caulk wet areas every 5-7 years or when cracking is visible.'),
  R('PROP-001','AST-0007','Microwave — Door Latch Replacement','Appliances','Kitchen','Corrective','Medium','Complete','3/10/2024','3/11/2024','3/11/2024','Professional','AppliancePro Repair',28,65,false,false,'Microwave door not latching properly — unit would not power on.','Door latch tab broken — plastic fatigue.','Replaced door latch assembly.','Microwave door latches commonly need replacement every 5-8 years.'),
  R('PROP-001','','Deck — Tighten Rail Post Hardware','Exterior','Patio / Deck','Preventive','Medium','Complete','4/30/2025','4/30/2025','4/30/2025','DIY','Self',5,0,false,false,'Two deck rail posts had visible play — safety concern.','Post hardware loosened from seasonal expansion/contraction.','Tightened all through-bolts and replaced 2 carriage bolts.','Inspect all deck hardware annually during spring deck check.'),
  R('PROP-001','AST-0011','Sump Pump — Replace Backup Battery','Plumbing','Basement','Corrective','High','Complete','4/5/2024','4/5/2024','4/5/2024','DIY','Self',95,0,false,false,'Backup battery unit not holding charge — 5 years old.','Lead-acid battery exceeded typical 3-5 year service life.','Replaced AGM backup battery. Tested float. System operating.','Test backup annually; replace battery proactively every 3-4 years.'),
  R('PROP-001','','Bathroom — Replace Toilet Flapper (x2)','Plumbing','Bathroom','Corrective','Low','Complete','11/15/2023','11/15/2023','11/15/2023','DIY','Self',12,0,false,false,'Both main-floor toilets running intermittently — phantom flush.','Flapper rubber degraded after ~7 years; no longer sealing.','Replaced flappers in both toilets.','Replace flappers every 5 years proactively to prevent water waste.'),
  R('PROP-001','AST-0001','Furnace — Clean Blower & Check Belt','HVAC','HVAC','Preventive','Medium','Complete','11/1/2024','11/1/2024','11/1/2024','Professional','AirPro HVAC LLC',0,0,false,false,'Blower cleaning completed during annual furnace tune-up.','Routine annual maintenance — belt wear normal.','Belt in good condition. Blower wheel cleaned. No action required.','Next blower belt replacement estimated 3-5 years.'),
  R('PROP-001','','Windows — Reglaze Failed Seal (3 Units)','Interior','Whole Home','Corrective','Medium','Complete','9/15/2023','9/18/2023','9/22/2023','Professional','Ridgeline Home Services',280,320,false,true,'3 double-pane windows showing foggy condensation between panes.','IGU seal failure — argon gas escaped, moisture entered.','Three window IGUs replaced under manufacturer warranty (still active).','Inspect windows annually for fogging. Remaining windows under 8-year warranty.'),
  R('PROP-001','AST-0025','Lawn Mower — Blade Sharpening & Oil Change','Landscaping','Yard / Landscaping','Preventive','Low','Complete','4/6/2025','4/6/2025','4/6/2025','DIY','Self',38,0,false,false,'Annual spring service — mower stored since October.','Routine annual maintenance.','Oil changed, blade sharpened, air filter replaced.','Annual spring service keeps mower running efficiently.'),
  R('PROP-001','','Attic — Seal Top Plate Air Leaks','Energy Efficiency','Attic','Preventive','Medium','Complete','4/1/2025','4/5/2025','4/8/2025','DIY','Self',35,0,false,false,'Energy audit identified air sealing over top plate as #1 improvement opportunity.','Original construction lacked air sealing between conditioned and attic space.','Foam-sealed all top plate penetrations and can light air gaps in attic.','Re-test with blower door in 2 years to confirm improvement.'),

  // ── Active / Upcoming repairs ─────────────────────────────────────────────────
  R('PROP-001','AST-0013','CO Detectors — Scheduled Replacement (x3)','Safety','Safety','Replacement','High','In Progress','1/15/2026','2/1/2026','','DIY','Self',60,0,false,false,'Units approaching 7-year end-of-life. Replacement ordered.','Kidde CO detectors rated 7 years from manufacture date.','3-pack ordered. Scheduled installation this month.','Set calendar reminder for next replacement in 7 years (2033).'),
  R('PROP-001','AST-0003','Water Heater — Plan Replacement','Plumbing','Plumbing','Replacement','Medium','Diagnosing','2/1/2026','','','Professional','ClearFlow Plumbing',0,0,false,false,'Water heater is 7 years old — past midpoint of 12-year useful life. Budget planning.','Anode rod inspection shows unit has 2-3 years remaining life.','Gathering quotes for hybrid heat pump water heater upgrade.','Consider AO Smith HPTU-50N or Rheem ProTerra. Rebates may be available.'),
  R('PROP-001','','East Wall Efflorescence — Annual Monitor','Exterior','Basement','Inspection','Low','Monitoring','4/1/2026','','','DIY','Self',0,0,false,false,'Hairline crack sealed in 2024. Monitor for recurrence.','Grading corrected; monitor if moisture continues.','Will inspect after spring thaw and heavy rain.','If efflorescence returns, consider interior drain tile system.'),
  R('PROP-001','AST-0002','AC — Schedule Tune-Up Spring 2026','HVAC','HVAC','Preventive','High','Scheduled','5/1/2026','5/1/2026','','Professional','AirPro HVAC LLC',0,165,false,false,'Annual AC tune-up. Appointment to be booked in April.','Routine seasonal maintenance.','',''),
  R('PROP-001','','Master Bath — Replace Exhaust Fan','Interior','Bathroom','Corrective','Medium','Scheduled','6/1/2026','','','DIY','Self',45,0,false,false,'Master bath fan running at reduced CFM — likely worn bearing.','Fan motor bearing wear after 7 years of use.','Researching replacement unit (Broan-Nutone 110 CFM). Planning spring installation.','Replace bathroom exhaust fans every 8-10 years.'),
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title row ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 13, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['Home Repair Log']] });

  // ── Subtitle row ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Every repair — completed, active, and planned — with root cause, resolution, and prevention notes']] });

  // ── KPI cards ─────────────────────────────────────────────────────────────────
  const kpiDefs = [
    { label: 'Total Repairs', col: 0, span: 4, color: C.primary },
    { label: 'Active / Planned', col: 5, span: 4, color: C.warning },
    { label: 'Total Repair Cost', col: 11, span: 5, color: C.secondary },
    { label: 'Warranty Saves', col: 17, span: 4, color: C.success },
    { label: 'DIY Saves', col: 22, span: 3, color: C.olive },
  ];

  const D1 = 6;
  const DE = 500;

  for (const k of kpiDefs) {
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(k.color), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }

  const kpiFormulas = [
    { col: 'A', formula: `=COUNTA(A${D1+1}:A${DE})` },
    { col: 'F', formula: `=SUMPRODUCT((I${D1+1}:I${DE}<>"Complete")*(I${D1+1}:I${DE}<>"Cancelled")*(I${D1+1}:I${DE}<>"")*1)` },
    { col: 'L', formula: `=SUMPRODUCT((ISNUMBER(Q${D1+1}:Q${DE}))*Q${D1+1}:Q${DE})` },
    { col: 'R', formula: `=SUMPRODUCT((S${D1+1}:S${DE}=TRUE)*ISNUMBER(Q${D1+1}:Q${DE})*Q${D1+1}:Q${DE})` },
    { col: 'W', formula: `=SUMPRODUCT((M${D1+1}:M${DE}="DIY")*ISNUMBER(Q${D1+1}:Q${DE})*Q${D1+1}:Q${DE})` },
  ];
  for (const { col, formula } of kpiFormulas) {
    vals.push({ range: `${S}!${col}4`, values: [[formula]] });
  }

  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 4), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 5, 9), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 11, 16), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 17, 21), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 22, 25), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  for (const [ri, px] of [[0,34],[1,24],[2,22],[3,42]]) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  }

  // ── Divider row (row 4) ───────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ── Header row (row 5, 0-indexed) ────────────────────────────────────────────
  const HEADERS = [
    'Repair ID','Property ID','Asset ID','Repair / Issue','Category','Home Area','Type','Priority',
    'Status','Date Reported','Date Started','Date Completed','DIY or Pro','Provider',
    'Parts & Materials ($)','Labor Cost ($)','Total Cost (fml)','Warranty Claim?','Covered by Warranty?',
    'Days Active (fml)','Year (fml)','Urgency Notes','Root Cause','Resolution','Prevention Notes'
  ];
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A6:Y6`, values: [HEADERS] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 42 }, fields: 'pixelSize' } });

  // ── Data rows ────────────────────────────────────────────────────────────────
  const DR = REPAIRS.length;
  const DS = D1; // 0-indexed data start

  const dataRows = [];
  REPAIRS.forEach((row, ri) => {
    const r = ri + D1 + 1;
    // A(0) = Repair ID
    row[0] = `=IF(B${r}="","","REP-"&TEXT(ROW()-6,"0000"))`;
    // Q(16) = Total Cost = Parts + Labor
    row[16] = `=IF(AND(O${r}="",P${r}=""),"",IFERROR(VALUE(O${r}),0)+IFERROR(VALUE(P${r}),0))`;
    // T(19) = Days Active
    row[19] = `=IF(J${r}="","",IF(L${r}<>"",L${r}-J${r},TODAY()-J${r}))`;
    // U(20) = Year
    row[20] = `=IF(J${r}<>"",YEAR(J${r}),"")`;
    dataRows.push(row);
  });

  vals.push({ range: `${S}!A7:Y${6+DR}`, values: dataRows });

  // ── Data formatting ────────────────────────────────────────────────────────────
  for (let ri = 0; ri < DR; ri++) {
    const rowIdx = DS + ri;
    const bg = ri % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Formula cols: A(0), Q(16), T(19), U(20)
  for (const ci of [0, 16, 19, 20]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Input bg
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 1, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 17, 19), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 21, 25), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Date formats: J(9), K(10), L(11)
  for (const ci of [9, 10, 11]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // Currency: O(14), P(15), Q(16) formula
  for (const ci of [14, 15, 16]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // Days Active number
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 19, 20), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Checkboxes: R(17) Warranty Claim, S(18) Covered
  for (const ci of [17, 18]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, ci, ci+1), cell: { dataValidation: { condition: { type: 'BOOLEAN' }, showCustomUi: true } }, fields: 'dataValidation' } });
  }

  // Wrap: V(21), W(22), X(23), Y(24)
  fmt.push({ repeatCell: { range: gridRange(SID, DS, DS+DR, 21, 25), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat.wrapStrategy' } });

  // ── Data validations ─────────────────────────────────────────────────────────
  const valDefs = [
    [4, ['HVAC','Plumbing','Electrical','Roof & Gutters','Exterior','Interior','Appliances','Safety','Landscaping','Cleaning','Pest Prevention','Water Systems','Energy Efficiency','Seasonal Preparation','General Inspection','Other']],
    [5, ['Whole Home','Kitchen','Bathroom','Bedroom','Living Room','Dining Room','Laundry','Basement','Attic','Garage','Exterior','Roof','Yard / Landscaping','Driveway','Patio / Deck','HVAC','Plumbing','Electrical','Safety','Other']],
    [6, ['Preventive','Corrective','Emergency','Cosmetic','Replacement','Inspection','Other']],
    [7, ['Low','Medium','High','Urgent']],
    [8, ['Reported','Diagnosing','Scheduled','In Progress','Waiting for Parts','Complete','Monitoring','Cancelled']],
    [12, ['DIY','Professional','Mixed','Not Decided']],
  ];
  for (const [col, list] of valDefs) {
    fmt.push({ setDataValidation: { range: gridRange(SID, DS, DS+DR, col, col+1), rule: { condition: { type: 'ONE_OF_LIST', values: list.map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: true } } });
  }

  // ── Row heights ───────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DS, endIndex: DS+DR }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const colWidths = [75,70,80,220,110,100,90,70,100,90,90,90,80,140,90,80,85,75,80,75,55,180,180,200,180];
  colWidths.forEach((px, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'repair-fmt');
  await valuesBatchUpdate(id, vals, 'repair-vals');

  console.log(`✓ Home Repair Log complete — ${DR} repairs written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
