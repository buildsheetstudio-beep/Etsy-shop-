'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recurring Maintenance'];
const S = "'Recurring Maintenance'";
const REF = "'Reference Data'";
const NC = 25; // A-Y

// Col mapping:
// A(0)=Template ID  B(1)=Property ID   C(2)=Task Name        D(3)=Category
// E(4)=Home Area    F(5)=Season        G(6)=Priority         H(7)=DIY or Pro
// I(8)=Frequency    J(9)=Custom Interval  K(10)=Custom Unit  L(11)=Est Duration Hr
// M(12)=Est Cost    N(13)=Preferred Month  O(14)=Anchor Date  P(15)=Next Due(fml)
// Q(16)=Last Done   R(17)=Active Checkbox  S(18)=Notes
// T-Y reserved / spare

// colLetter helper for 25-col sheet
function colL(i) {
  if (i < 26) return String.fromCharCode(65 + i);
  return String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26));
}

const R = (propId, name, cat, area, season, pri, diyPro, freq, custInt, custUnit, estHr, estCost, prefMo, anchor, lastDone, active=true, notes='') => {
  const row = Array(NC).fill('');
  row[1] = propId;
  row[2] = name;
  row[3] = cat;
  row[4] = area;
  row[5] = season;
  row[6] = pri;
  row[7] = diyPro;
  row[8] = freq;
  row[9] = custInt;
  row[10] = custUnit;
  row[11] = estHr;
  row[12] = estCost;
  row[13] = prefMo;
  row[14] = anchor;
  // col P(15) = Next Due — formula col, filled separately
  row[16] = lastDone;
  row[17] = active;
  row[18] = notes;
  return row;
};

const TEMPLATES = [
  // ── HVAC ─────────────────────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Replace HVAC Air Filter (1-inch)','HVAC','HVAC','Year-Round','High','DIY','Monthly','','','0.25',5,'','',       '1/15/2026',true,'Check filter monthly; replace when grey. 1-inch filters need monthly attention.'),
  R('PROP-001','HVAC Furnace Tune-Up & Safety Check','HVAC','HVAC','Fall','High','Professional','Annual','','','2',180,'November','11/1','11/1/2025',true,'Book October for November slot. Includes burner inspection, heat exchanger check, flue test.'),
  R('PROP-001','Central AC Tune-Up & Coil Cleaning','HVAC','HVAC','Spring','High','Professional','Annual','','','2',160,'May','5/1','5/15/2025',true,'Book April for May slot. Includes refrigerant check, coil cleaning, drain pan flush.'),
  R('PROP-001','Clean AC Condensate Drain Line','HVAC','HVAC','Spring','Medium','DIY','Annual','','','0.5',0,'May','5/1','5/1/2025',true,'Flush with diluted bleach to prevent algae clog and water overflow.'),
  R('PROP-001','Inspect & Clean Dryer Vent Duct','HVAC','Laundry','Fall','High','DIY','Annual','','','1',0,'October','10/1','10/1/2025',true,'Fire hazard if lint-clogged. Disconnect dryer and use vent brush kit.'),
  R('PROP-001','Check & Lubricate HVAC Blower Motor','HVAC','HVAC','Fall','Low','Professional','Annual','','','1',0,'November','11/1','11/1/2025',true,'Done during annual furnace tune-up if included.'),

  // ── Plumbing ─────────────────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Test & Inspect Sump Pump','Plumbing','Basement','Spring','High','DIY','Semiannual','','','0.5',0,'April,October','4/1','4/1/2025',true,'Pour water in pit to test float switch. Check discharge line is clear.'),
  R('PROP-001','Flush Water Heater (Sediment)','Plumbing','Plumbing','Fall','Medium','DIY','Annual','','','1',0,'September','9/1','9/1/2025',true,'Connect hose to drain valve; flush until clear. Extends tank life.'),
  R('PROP-001','Inspect Water Heater Anode Rod','Plumbing','Plumbing','Fall','Medium','DIY','Every 3 Years','','','2',30,'September','9/1','9/1/2022',true,'Replace if less than 1/2 inch of core wire remains.'),
  R('PROP-001','Test Pressure Relief Valve (Water Heater)','Plumbing','Plumbing','Spring','Medium','DIY','Annual','','','0.25',0,'April','4/1','4/1/2025',true,'Lift lever briefly; water should discharge. If stuck, replace valve.'),
  R('PROP-001','Inspect All Plumbing Supply Lines','Plumbing','Whole Home','Spring','Medium','DIY','Annual','','','0.5',0,'April','4/1','4/1/2025',true,'Check under sinks and behind appliances for cracks or bulges in braided hoses.'),
  R('PROP-001','Winterize Hose Bibs / Shut-Off Interior Valves','Plumbing','Exterior','Fall','High','DIY','Annual','','','0.5',0,'November','11/1','10/15/2025',true,'Drain exterior spigots and close interior shut-off valves before first freeze.'),
  R('PROP-001','Activate & Inspect Hose Bibs (Spring)','Plumbing','Exterior','Spring','Medium','DIY','Annual','','','0.25',0,'April','4/15','4/15/2025',true,'Reopen interior valves, check for winter damage at spigots.'),
  R('PROP-001','Clean Bathroom Drain Stoppers','Plumbing','Bathroom','Summer','Low','DIY','Quarterly','','','0.5',0,'','','10/1/2025',true,'Remove and clean hair/soap buildup from all bathroom drain stoppers.'),
  R('PROP-001','Run Water Softener Salt Check','Water Systems','Plumbing','Year-Round','Medium','DIY','Monthly','','','0.1',20,'','','1/15/2026',true,'Check salt level in brine tank; top off as needed. Use appropriate salt type.'),

  // ── Electrical ────────────────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Test GFCI & AFCI Outlets / Breakers','Electrical','Electrical','Spring','High','DIY','Semiannual','','','0.5',0,'April,October','4/1','4/1/2025',true,'Press test/reset on each GFCI outlet. Test AFCI breakers at panel.'),
  R('PROP-001','Test Smoke Detectors (All Zones)','Safety','Safety','Year-Round','Urgent','DIY','Monthly','','','0.25',0,'','','1/15/2026',true,'Press test button on each unit. Replace batteries if chirping.'),
  R('PROP-001','Test Carbon Monoxide Detectors','Safety','Safety','Year-Round','Urgent','DIY','Monthly','','','0.1',0,'','','1/15/2026',true,'Press test button. Note: CO detector units have 5-7 year rated life.'),
  R('PROP-001','Replace Smoke Detector Batteries','Safety','Safety','Fall','High','DIY','Annual','','','0.5',15,'October','10/1','10/1/2025',true,'Replace all smoke detector batteries at daylight saving time change.'),
  R('PROP-001','Inspect Electrical Panel (Visual)','Electrical','Electrical','Summer','Medium','DIY','Annual','','','0.25',0,'July','7/1','7/1/2025',true,'Check for tripped breakers, signs of heat/scorching, or unusual odors.'),
  R('PROP-001','Inspect Fire Extinguishers','Safety','Safety','Winter','High','DIY','Annual','','','0.25',0,'January','1/15','1/15/2026',true,'Check pressure gauge is in green zone. Inspect pin and tamper seal.'),
  R('PROP-001','Replace Fire Extinguisher (if expired)','Safety','Safety','Winter','High','DIY','Every 5 Years','','','0.25',60,'January','1/15','1/15/2024',true,'Replace or professionally recharge if past rated date.'),

  // ── Roof & Exterior ──────────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Clean & Inspect Gutters — Spring','Roof & Gutters','Exterior','Spring','High','DIY','Annual','','','2',0,'April','4/15','4/15/2025',true,'Remove winter debris, flush downspouts, check for sags or loose hangers.'),
  R('PROP-001','Clean & Inspect Gutters — Fall','Roof & Gutters','Exterior','Fall','High','DIY','Annual','','','2',0,'November','11/1','11/1/2025',true,'Clear fallen leaves before winter. Check downspout extensions are diverting water away.'),
  R('PROP-001','Roof Visual Inspection (Binoculars)','Roof & Gutters','Roof','Spring','Medium','DIY','Annual','','','0.5',0,'April','4/15','4/15/2025',true,'Check for missing/lifted shingles, granule loss in gutters, flashing condition.'),
  R('PROP-001','Professional Roof Inspection','Roof & Gutters','Roof','Spring','Medium','Professional','Every 3 Years','','','2',250,'April','4/15','4/15/2024',true,'Recommended every 3-5 years or after major storms.'),
  R('PROP-001','Inspect & Caulk Window & Door Seals','Exterior','Exterior','Spring','Medium','DIY','Annual','','','2',20,'April','4/15','4/15/2025',true,'Check caulking around windows, doors, and penetrations. Recaulk any gaps.'),
  R('PROP-001','Touch-Up Exterior Paint / Trim','Exterior','Exterior','Summer','Low','DIY','Annual','','','3',40,'July','7/1','7/1/2025',true,'Touch up peeling or chipped areas to prevent wood rot.'),
  R('PROP-001','Inspect Foundation & Grading','Exterior','Exterior','Spring','Medium','DIY','Annual','','','1',0,'April','4/15','4/15/2025',true,'Check for cracks in foundation. Ensure soil slopes away from house at least 6 inches per 10 feet.'),
  R('PROP-001','Power Wash Siding / Driveway','Exterior','Exterior','Spring','Low','DIY','Annual','','','3',0,'April','4/15','4/15/2025',true,'Rinse winter salt and grime off siding and concrete. Inspect siding condition.'),

  // ── Interior ────────────────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Inspect Attic for Moisture / Pests','Interior','Attic','Summer','Medium','DIY','Annual','','','0.5',0,'July','7/1','7/1/2025',true,'Check insulation for moisture stains, animal intrusion, and soffit vent blockage.'),
  R('PROP-001','Lubricate Garage Door Springs & Rollers','Interior','Garage','Fall','Medium','DIY','Semiannual','','','0.5',5,'April,October','4/1','10/1/2025',true,'Use silicone or lithium spray — not WD-40. Lubricate springs, rollers, hinges, and tracks.'),
  R('PROP-001','Test Garage Door Auto-Reverse Safety','Interior','Garage','Spring','High','DIY','Annual','','','0.25',0,'April','4/1','4/1/2025',true,'Place 2x4 flat on floor under door; door should auto-reverse. Also test wall button stop.'),
  R('PROP-001','Inspect & Clean Bathroom Exhaust Fans','Interior','Bathroom','Spring','Medium','DIY','Annual','','','1',0,'April','4/15','4/15/2025',true,'Vacuum dust from fan grills. Check CFM by holding tissue to confirm airflow.'),
  R('PROP-001','Clean Refrigerator Condenser Coils','Appliances','Kitchen','Spring','Medium','DIY','Annual','','','0.5',0,'April','4/15','4/15/2025',true,'Unplug fridge, pull from wall, vacuum coils. Improves efficiency and longevity.'),
  R('PROP-001','Clean Oven / Range (Self-Clean Cycle)','Appliances','Kitchen','Fall','Low','DIY','Semiannual','','','0.5',0,'','','10/1/2025',true,'Run self-clean cycle. Remove racks first; ventilate kitchen well.'),
  R('PROP-001','Clean Dishwasher Filter & Spray Arms','Appliances','Kitchen','Spring','Low','DIY','Quarterly','','','0.25',0,'','','10/1/2025',true,'Remove and rinse filter monthly. Check spray arm holes for debris quarterly.'),
  R('PROP-001','Run Washing Machine Drum-Clean Cycle','Appliances','Laundry','Year-Round','Low','DIY','Monthly','','','0.5',3,'','','1/15/2026',true,'Use drum-clean tablet or tub-clean cycle. Leave door open between uses to prevent mold.'),
  R('PROP-001','Inspect Washer Hoses for Bulging/Cracks','Appliances','Laundry','Spring','High','DIY','Annual','','','0.25',0,'April','4/1','4/1/2025',true,'Braided stainless hoses recommended. Replace if any bulging, cracks, or leaks.'),

  // ── Landscaping / Seasonal ────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Activate Irrigation System (Spring)','Landscaping','Yard / Landscaping','Spring','Medium','DIY','Annual','','','1',0,'April','4/15','4/15/2025',true,'Open main valve slowly, test each zone, check for broken heads or leaks.'),
  R('PROP-001','Winterize Irrigation System (Fall)','Landscaping','Yard / Landscaping','Fall','High','Professional','Annual','','','1',80,'October','10/15','10/15/2025',true,'Blow out lines with compressed air before first hard freeze.'),
  R('PROP-001','Service Lawn Mower (Tune-Up)','Landscaping','Yard / Landscaping','Spring','Medium','DIY','Annual','','','1.5',30,'April','4/1','4/1/2025',true,'Change oil, replace spark plug, sharpen blade, clean air filter.'),
  R('PROP-001','Service Snow Blower (Pre-Season)','Landscaping','Yard / Landscaping','Fall','Medium','DIY','Annual','','','1',25,'November','11/1','11/1/2025',true,'Check oil, replace spark plug, lubricate auger, verify shear pins.'),
  R('PROP-001','Dethatch & Overseed Lawn','Landscaping','Yard / Landscaping','Fall','Medium','DIY','Annual','','','4',80,'September','9/1','9/1/2025',true,'Fall seeding gives best germination for cool-season grass.'),
  R('PROP-001','Apply Lawn Fertilizer (Spring)','Landscaping','Yard / Landscaping','Spring','Medium','DIY','Annual','','','1',60,'April','4/15','4/15/2025',true,'Use slow-release balanced fertilizer after soil temp reaches 55°F.'),
  R('PROP-001','Apply Pre-Emergent Weed Control','Landscaping','Yard / Landscaping','Spring','Medium','DIY','Annual','','','1',40,'April','4/1','4/1/2025',true,'Apply before soil temps reach 55°F to prevent crabgrass germination.'),
  R('PROP-001','Mulch Garden Beds','Landscaping','Yard / Landscaping','Spring','Low','DIY','Annual','','','3',60,'May','5/1','5/1/2025',true,'2-3 inches of mulch suppresses weeds and retains moisture. Keep away from tree trunks.'),

  // ── Pest Prevention ───────────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Pest Control Perimeter Treatment','Pest Prevention','Exterior','Spring','Medium','Professional','Semiannual','','','1',150,'April,October','4/15','4/15/2025',true,'Exterior spray for ants, spiders, and wasps. Book before summer peak.'),
  R('PROP-001','Inspect & Seal Pest Entry Points','Pest Prevention','Exterior','Fall','Medium','DIY','Annual','','','1',15,'October','10/1','10/15/2025',true,'Seal gaps around pipes, foundation cracks, and utility penetrations with foam or caulk.'),

  // ── General ─────────────────────────────────────────────────────────────────────────────────────────────────
  R('PROP-001','Change Whole-Home Air Purifier Filters','Cleaning','HVAC','Year-Round','Medium','DIY','Every 4 Months','','','0.25',80,'','','9/1/2025',true,'HEPA + carbon filters per manufacturer schedule. Check pre-filter monthly.'),
  R('PROP-001','Deep Clean Kitchen (Hood, Vents, Backsplash)','Cleaning','Kitchen','Fall','Low','DIY','Semiannual','','','3',0,'','','10/1/2025',true,'Degrease range hood filter, clean backsplash grout, wipe cabinet fronts.'),
  R('PROP-001','Annual Home Energy Audit','Energy Efficiency','Whole Home','Spring','Low','Professional','Annual','','','3',0,'March','3/1','3/1/2025',true,'Utility company often offers free or subsidized audit. Identifies air sealing and insulation opportunities.'),
  R('PROP-001','Inspect & Re-Caulk Deck Boards & Rails','Exterior','Patio / Deck','Spring','Medium','DIY','Annual','','','2',20,'April','4/15','4/15/2025',true,'Composite deck: check board clips, railing post hardware, and ledger board attachment.'),
  R('PROP-001','Whole-Home Seasonal Walk-Through Inspection','General Inspection','Whole Home','Seasonal','Medium','DIY','Seasonal','','','2',0,'','','10/15/2025',true,'Comprehensive check of all systems each season. See Seasonal Planner tab.'),
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Expand columns if needed (25 cols = Y; default 26 so no expansion needed) ──

  // ── Title row ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 13, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['Recurring Maintenance Templates']] });

  // ── Subtitle row ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { italic: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Define recurring tasks once — generate scheduled tasks from these templates']] });

  // ── KPI card labels row (row 2, 0-indexed) ───────────────────────────────────
  // 4 KPI cards: Active Templates | Monthly Cost | Annual Cost | % DIY
  const kpiDefs = [
    { label: 'Active Templates', col: 0, span: 3, color: C.primary },
    { label: 'Est. Monthly Cost', col: 6, span: 4, color: C.secondary },
    { label: 'Est. Annual Cost', col: 13, span: 4, color: C.wheat },
    { label: '% DIY Tasks', col: 20, span: 5, color: C.olive },
  ];

  for (const k of kpiDefs) {
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(k.color), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }
  vals.push({ range: `${S}!A3:Y3`, values: [kpiDefs.map(() => '').reduce((a, _, i) => { const k = kpiDefs.find(k2 => k2.col === i); if (k) a[i] = k.label; return a; }, Array(NC).fill(''))] });

  // ── KPI values row (row 3, 0-indexed) ────────────────────────────────────────
  const DATA_START = 6; // 1-indexed row where data starts (row 7)
  const DATA_END = 300;

  for (const k of kpiDefs) {
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }

  // KPI formulas — all based on column R (index 17 = R, 1-indexed col 18)
  // Active: count rows where R (Active) = TRUE
  // Monthly: SUMPRODUCT where freq=Monthly and active
  // Annual: SUM of estimated yearly cost per row (Monthly*12, Quarterly*4, etc.)
  // DIY%: count DIY / total active
  const kpiFormulas = [
    `=SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*1)`,   // Active Templates (col A)
    `=SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*(ISNUMBER(M${DATA_START+1}:M${DATA_END}))*M${DATA_START+1}:M${DATA_END}*(I${DATA_START+1}:I${DATA_END}="Monthly")*1)+SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*(ISNUMBER(M${DATA_START+1}:M${DATA_END}))*M${DATA_START+1}:M${DATA_END}*(I${DATA_START+1}:I${DATA_END}="Weekly")*4.33)`,   // Monthly cost (col G)
    `=SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*(ISNUMBER(M${DATA_START+1}:M${DATA_END}))*M${DATA_START+1}:M${DATA_END}*(I${DATA_START+1}:I${DATA_END}="Annual")*1)+SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*(ISNUMBER(M${DATA_START+1}:M${DATA_END}))*M${DATA_START+1}:M${DATA_END}*(I${DATA_START+1}:I${DATA_END}="Monthly")*12)+SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*(ISNUMBER(M${DATA_START+1}:M${DATA_END}))*M${DATA_START+1}:M${DATA_END}*(I${DATA_START+1}:I${DATA_END}="Quarterly")*4)+SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*(ISNUMBER(M${DATA_START+1}:M${DATA_END}))*M${DATA_START+1}:M${DATA_END}*(I${DATA_START+1}:I${DATA_END}="Semiannual")*2)`,   // Annual cost (col N)
    `=IFERROR(SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*(H${DATA_START+1}:H${DATA_END}="DIY"))/SUMPRODUCT((R${DATA_START+1}:R${DATA_END}=TRUE)*1),0)`,   // DIY % (col U)
  ];
  vals.push({ range: `${S}!A4`, values: [[kpiFormulas[0]]] });
  vals.push({ range: `${S}!G4`, values: [[kpiFormulas[1]]] });
  vals.push({ range: `${S}!N4`, values: [[kpiFormulas[2]]] });
  vals.push({ range: `${S}!U4`, values: [[kpiFormulas[3]]] });

  // Number formats for KPI values
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 3), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 6, 10), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 13, 17), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 20, 25), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Row height for title/subtitle/KPI rows
  for (const [ri, px] of [[0,34],[1,24],[2,22],[3,42]]) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  }

  // ── Divider row (row 4, 0-indexed) ───────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.wheat) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ── Header row (row 5, 0-indexed) ────────────────────────────────────────────
  const HEADERS = [
    'Template ID','Property ID','Task Name','Category','Home Area','Season','Priority',
    'DIY or Pro','Frequency','Custom Interval','Custom Unit','Est. Duration (hr)',
    'Est. Cost ($)','Preferred Month','Anchor Date','Next Due','Last Done',
    'Active','Notes','','','','','',''
  ];
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A6:Y6`, values: [HEADERS] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Data rows ────────────────────────────────────────────────────────────────
  const dataRows = [];
  TEMPLATES.forEach((row, ri) => {
    const r = ri + DATA_START + 1; // 1-indexed spreadsheet row
    // A(0) = Template ID (formula)
    row[0] = `=IF(B${r}="","","REC-"&TEXT(ROW()-6,"000"))`;
    // P(15) = Next Due (formula based on anchor date and last done)
    // Simple logic: if Last Done is empty, use Anchor; else use Last Done + frequency offset
    // For simplicity: if Q (Last Done) = "", return Anchor Date, else return ""
    // (Users update manually or let task log drive scheduling)
    row[15] = `=IF(Q${r}="",IF(O${r}="","",DATEVALUE(O${r})),IF(I${r}="Monthly",EDATE(Q${r},1),IF(I${r}="Annual",EDATE(Q${r},12),IF(I${r}="Quarterly",EDATE(Q${r},3),IF(I${r}="Semiannual",EDATE(Q${r},6),IF(I${r}="Every 2 Years",EDATE(Q${r},24),IF(I${r}="Every 3 Years",EDATE(Q${r},36),IF(I${r}="Every 5 Years",EDATE(Q${r},60),IF(I${r}="Weekly",Q${r}+7,IF(I${r}="Every 2 Months",EDATE(Q${r},2),IF(I${r}="Every 4 Months",EDATE(Q${r},4),IF(I${r}="Every 18 Months",EDATE(Q${r},18),""))))))))))))`;
    dataRows.push(row);
  });

  vals.push({ range: `${S}!A7:Y${6+TEMPLATES.length}`, values: dataRows });

  // ── Data formatting ────────────────────────────────────────────────────────────
  const DR = TEMPLATES.length;
  // Alt row banding
  for (let ri = 0; ri < DR; ri++) {
    const rowIdx = DATA_START + ri;
    const bg = ri % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // Formula cols: A(0) = Template ID, P(15) = Next Due — formula blue
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, 15, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.formula), numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } }, fields: 'userEnteredFormat(backgroundColor,numberFormat)' } });

  // Date cols: Q(16) Last Done
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, 16, 17), cell: { userEnteredFormat: { backgroundColor: hex(C.input), numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } }, fields: 'userEnteredFormat(backgroundColor,numberFormat)' } });

  // Currency: M(12) Est Cost
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, 12, 13), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Number: L(11) Est Duration, J(9) Custom Interval
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, 11, 12), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Input bg for editable cols (B-N, Q)
  for (const [c1, c2] of [[1, 15], [16, 17]]) {
    fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, c1, c2), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // Active checkbox (R = col 17)
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, 17, 18), cell: { dataValidation: { condition: { type: 'BOOLEAN' }, showCustomUi: true } }, fields: 'dataValidation' } });

  // Notes col (S = col 18) wrap text
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+DR, 18, 19), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat.wrapStrategy' } });

  // ── Data validations ─────────────────────────────────────────────────────────
  const REF_SID = sheetMap['Reference Data'];
  const valDefs = [
    // [col, listSource (array) or range string]
    [3, ['HVAC','Plumbing','Electrical','Roof & Gutters','Exterior','Interior','Appliances','Safety','Landscaping','Cleaning','Pest Prevention','Water Systems','Energy Efficiency','Seasonal Preparation','General Inspection','Other']],
    [4, ['Whole Home','Kitchen','Bathroom','Bedroom','Living Room','Dining Room','Laundry','Basement','Attic','Garage','Exterior','Roof','Yard / Landscaping','Driveway','Patio / Deck','HVAC','Plumbing','Electrical','Safety','Other']],
    [5, ['Winter','Spring','Summer','Fall','Year-Round','Seasonal']],
    [6, ['Low','Medium','High','Urgent']],
    [7, ['DIY','Professional','Mixed','Not Decided']],
    [8, ['Weekly','Monthly','Every 2 Months','Quarterly','Every 4 Months','Semiannual','Annual','Every 18 Months','Every 2 Years','Every 3 Years','Every 5 Years','Custom Months','Custom Years','Seasonal','Usage-Based','One-Time']],
    [10, ['Hours','Months','Miles','Cycles','Filters','Gallons','Other']],
  ];
  for (const [col, list] of valDefs) {
    fmt.push({ setDataValidation: { range: gridRange(SID, DATA_START, DATA_START+DR, col, col+1), rule: { condition: { type: 'ONE_OF_LIST', values: list.map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: true } } });
  }

  // ── Row heights ───────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DATA_START, endIndex: DATA_START+DR }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const colWidths = [80,70,220,130,120,80,70,90,100,70,70,80,80,100,90,90,90,55,220,50,50,50,50,50,50];
  colWidths.forEach((px, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // ── Freeze rows ───────────────────────────────────────────────────────────────
  // Title row (0) merges full width — no frozenColumnCount
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'rec-fmt');
  await valuesBatchUpdate(id, vals, 'rec-vals');

  console.log(`✓ Recurring Maintenance complete — ${TEMPLATES.length} templates written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
