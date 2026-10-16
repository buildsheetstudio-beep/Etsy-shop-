'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Appliances & Systems'];
const S = "'Appliances & Systems'";
const REP = "'Home Repair Log'";
const CST = "'Maintenance Costs'";
const NC = 25; // A-Y

// A(0)=Asset ID B(1)=Property ID C(2)=Name D(3)=Type E(4)=Area F(5)=Manufacturer
// G(6)=Model H(7)=Serial# I(8)=Purchase Date J(9)=Purchase Cost K(10)=Warranty Months
// L(11)=Warranty Exp M(12)=Warranty Status N(13)=Est Useful Life Yrs O(14)=Est Replacement Year
// P(15)=Condition Q(16)=Last Service R(17)=Next Service S(18)=Provider
// T(19)=Total Repair Cost U(20)=Total Maint Cost V(21)=Replacement Cost Est
// W(22)=Replacement Planning X(23)=Active Y(24)=Notes

const A = (b,c,d,e,f,g,h,i,j,k,n,p,q,r,s,v,w=false,x=true,y='') => {
  const row = Array(25).fill('');
  row[1]=b; row[2]=c; row[3]=d; row[4]=e; row[5]=f; row[6]=g; row[7]=h;
  row[8]=i; row[9]=j; row[10]=k; row[13]=n||''; row[15]=p; row[16]=q; row[17]=r;
  row[18]=s; row[21]=v; row[22]=w; row[23]=x; row[24]=y;
  return row;
};

// (propId, name, type, area, mfr, model, serial, purchaseDate, purchaseCost, warrantyMonths, usefulLife, condition, lastSvc, nextSvc, provider, replaceCost, replacePlanning, active, notes)
const ASSETS = [
  A('PROP-001','Gas Furnace','Furnace','HVAC','Carrier','Performance 96','CR96-7842','10/15/2019',2800,120,20,'Good','11/1/2025','11/1/2026','AirPro HVAC LLC',3500,false,true,'Annual tune-up done; filter replaced monthly. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Central Air Conditioner','Air Conditioner','HVAC','Carrier','24ACC636A','CA-5519-22B','10/15/2019',3200,120,15,'Good','5/15/2025','5/1/2026','AirPro HVAC LLC',4000,false,true,'R-410A system; coils cleaned spring 2025. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Gas Water Heater','Water Heater','Plumbing','Rheem','XG50T12HE40U0','RH-8821','3/20/2019',780,72,12,'Fair','3/20/2025','3/20/2026','ClearFlow Plumbing',1100,false,true,'Anode rod due for inspection. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Refrigerator','Refrigerator','Kitchen','Samsung','RF28R7351SR','SG-1143-7A','4/10/2019',1350,24,15,'Good','','','',1600,false,true,'French door, ice maker working. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Dishwasher','Dishwasher','Kitchen','Bosch','SHPM88Z75N','BH-3321','4/10/2019',980,24,12,'Good','','','',1100,false,true,'Stainless interior. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Gas Range / Oven','Oven / Range','Kitchen','GE','JGBS66REKSS','GE-9930','4/10/2019',720,12,20,'Good','','','',850,false,true,'5-burner gas. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Microwave','Microwave','Kitchen','Whirlpool','WMH31017HS','WP-5512','4/10/2019',280,12,10,'Good','','','',350,false,true,'Over-range. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Washing Machine','Washer','Laundry','LG','WM3900HWA','LG-8844','4/10/2019',920,24,12,'Good','','','',1050,false,true,'Front load. Drum clean cycle monthly. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Gas Dryer','Dryer','Laundry','LG','DLGX3901W','LG-8845','4/10/2019',820,24,13,'Good','','','',950,false,true,'Vent cleaned annually. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Garbage Disposal','Garbage Disposal','Kitchen','InSinkErator','Badger 5','IS-3391','4/10/2019',190,24,12,'Good','','','',250,false,true,'1/2 HP. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Sump Pump','Sump Pump','Basement','Zoeller','M53','ZL-5521','6/1/2021',380,36,10,'Good','4/1/2025','4/1/2026','ClearFlow Plumbing',480,false,true,'Primary pump; backup battery unit also installed. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Smoke Detectors (x6)','Smoke Detector','Safety','Kidde','i12010S-6P','','3/15/2019',90,120,10,'Good','1/1/2026','1/1/2027','Self',90,false,true,'Battery backup. Test monthly. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Carbon Monoxide Detectors (x3)','Carbon Monoxide Detector','Safety','Kidde','KN-COSM-BA','','3/15/2019',60,60,7,'Monitor','1/1/2026','1/1/2027','Self',60,true,true,'Expiring Soon — units approaching 7-year rated life. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Fire Extinguishers (x2)','Fire Extinguisher','Safety','Kidde','466112','','3/15/2019',50,0,12,'Good','1/1/2026','1/1/2032','Self',60,false,true,'Annual inspection required. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Electrical Panel','Electrical Panel','Electrical','Square D','QO130L200PG','SQD-7741','10/1/2019',1200,0,40,'Good','10/1/2021','10/1/2031','Volt & Wire Electric',2500,false,true,'200A service. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Asphalt Shingle Roof','Roof','Exterior','GAF','Timberline HDZ','','6/15/2021',14500,120,25,'Excellent','6/15/2021','6/1/2031','Summit Ridge Roofing',16000,false,true,'New 2021; dimensional shingle. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Gutters & Downspouts','Gutters','Exterior','LeafGuard','Seamless Aluminum','','6/15/2021',2800,24,25,'Good','10/15/2025','4/15/2026','Self / Ridgeline Home Services',3200,false,true,'Gutter guards installed. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Garage Door (x2 doors)','Garage Door','Garage','Clopay','Gallery Series 9x7','','10/5/2019',2400,12,25,'Good','10/1/2025','10/1/2026','Self',3000,false,true,'Two 9x7 insulated panels. Springs lubed annually. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Garage Door Openers (x2)','Garage Door','Garage','LiftMaster','8550W','','10/5/2019',680,24,15,'Good','','','',800,false,true,'Battery backup. Force sensors checked. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Attic Insulation','Interior','Attic','Owens Corning','Pink Batts R-38','','3/1/2019',2200,0,40,'Good','','','',2500,false,true,'R-38 blown-in. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Windows (x18)','Windows','Whole Home','Andersen','400 Series Double-Hung','','10/1/2019',12000,120,30,'Good','4/1/2025','4/1/2030','Self',14000,false,true,'Low-E glass, argon fill. Seals checked annually. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Exterior Doors (x3)','Exterior Doors','Exterior','Therma-Tru','Fiber-Classic 36in','','10/1/2019',2700,24,30,'Good','','','',3000,false,true,'Front, side, and rear. Weatherstripping checked annually. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Irrigation System','Irrigation System','Yard / Landscaping','Rain Bird','ESP-Me 6-zone','','5/10/2020',1400,24,20,'Good','10/15/2025','4/15/2026','Self',1600,false,true,'Winterized each fall, activated each spring. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Lawn Mower (Self-Propelled)','Lawn Equipment','Yard / Landscaping','Honda','HRX217VKA','HN-3312','4/1/2021',550,24,12,'Good','4/1/2025','4/1/2026','Self',600,false,true,'Annual service: blade sharpen, oil, air filter. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Snow Blower','Lawn Equipment','Garage','Ariens','Classic 24','AR-8812','11/15/2020',700,24,15,'Good','11/1/2025','11/1/2026','Self',800,false,true,'2-stage, 24-in clearing width. Pre-season tune-up. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Water Softener','Water Softener','Plumbing','Kinetico','Premier Series','KN-5541','3/20/2019',1800,48,20,'Good','3/20/2025','3/20/2026','ClearFlow Plumbing',2200,false,true,'Salt added monthly. Resin cleaned annually. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Basement Dehumidifier','Other','Basement','Frigidaire','FFAD7033R1','FG-4412','7/1/2022',280,12,10,'Good','','','',320,false,true,'Drains to floor drain. Set to 55% RH. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Bathroom Exhaust Fans (x3)','Other','Bathroom','Broan','AE80B','','3/15/2019',180,12,15,'Good','1/1/2025','1/1/2026','Self',200,false,true,'CFM checked annually. Clean grilles twice/year. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Whole-Home Air Purifier','HVAC System','HVAC','AprilAire','2410 Media Air Cleaner','AP-2410','10/15/2019',450,24,20,'Good','11/1/2025','11/1/2026','AirPro HVAC LLC',500,false,true,'MERV-11 filter; change with furnace filter. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Programmable Thermostat','HVAC System','HVAC','Ecobee','SmartThermostat Premium','EB-9921','10/15/2019',250,36,12,'Good','','','',280,false,true,'Calibrated annually at HVAC tune-up. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Dryer Vent System','Other','Laundry','Deflect-O','DPF4W 4-in Metal','','4/10/2019',80,0,20,'Good','4/1/2025','4/1/2026','Self',90,false,true,'Clean lint annually or more if usage is heavy. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Outdoor Hose Bibs (x4)','Plumbing Fixture','Exterior','Various','Frost-Free 3/4in','','10/1/2019',120,0,20,'Good','10/15/2025','4/15/2026','Self',130,false,true,'Winterize before freeze. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Plumbing Shut-Off Valves','Plumbing Fixture','Plumbing','Various','Ball Valve 3/4in','','10/1/2019',300,0,25,'Fair','3/1/2022','3/1/2027','ClearFlow Plumbing',350,false,true,'Main and branch shutoffs. Partial corrosion on one valve — monitor. Example planning estimate — verify for your equipment.'),
  A('PROP-001','Patio / Deck (Composite)','Other','Patio / Deck','Trex','Enhance Naturals','','7/15/2021',8500,0,25,'Excellent','7/15/2025','7/15/2026','Self',9500,false,true,'Composite decking; clean and inspect annually. Example planning estimate — verify for your equipment.'),
];

(async () => {
  const fmt = [];
  const vals = [];

  // Expand to 25 columns if needed
  // Default is 26 so 25 is fine

  // ── Title ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['APPLIANCES & SYSTEMS']] });

  // ── Subtitle ───────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Track all home appliances, systems, warranties, service history, and lifecycle status. Asset IDs link to all other tabs.']] });

  // ── Disclaimer ─────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.clay), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['Useful-life estimates are planning aids only. Follow manufacturer instructions, inspection findings, and professional guidance for your specific equipment.']] });

  // ── Summary cards (rows 3-6, 0-indexed) ──────────────────────────────────
  const CARD_LABELS = ['Active Assets','Warranty Active','Warranty Expiring Soon','Assets to Monitor','Replacement Planning','Repairs This Year','Total Repair Cost','Repl. Exposure Est.'];
  const CARD_FMLS = [
    `=COUNTIF($X$8:$X$507,TRUE)`,
    `=COUNTIFS($M$8:$M$507,"Active",$X$8:$X$507,TRUE)`,
    `=COUNTIFS($M$8:$M$507,"Expiring Soon",$X$8:$X$507,TRUE)`,
    `=COUNTIFS($P$8:$P$507,"Monitor",$X$8:$X$507,TRUE)`,
    `=COUNTIFS($W$8:$W$507,TRUE,$X$8:$X$507,TRUE)`,
    `=IFERROR(SUMPRODUCT((YEAR(IF(${REP}!$J$8:$J$2007="",DATE(1900,1,1),${REP}!$J$8:$J$2007))=YEAR(TODAY()))*(${REP}!$G$8:$G$2007<>"")),0)`,
    `=IFERROR(SUM($T$8:$T$507),0)`,
    `=IFERROR(SUMIFS($V$8:$V$507,$W$8:$W$507,TRUE),0)`,
  ];
  const CARD_COLORS = [C.primary,C.success,C.warning,C.warning,C.attention,C.secondary,C.attention,C.secondary];

  CARD_LABELS.forEach((lbl, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, i*3, i*3+3<=NC ? Math.min(i*3+3,NC) : NC), cell: { userEnteredFormat: { backgroundColor: hex(CARD_COLORS[i]), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, i*3, Math.min(i*3+3,NC)), cell: { userEnteredFormat: { backgroundColor: hex(CARD_COLORS[i]), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    // Only place values for i where col index < NC (avoid overflow)
    if (i*3 < NC) {
      vals.push({ range: `${S}!${colL(i*3)}4`, values: [[lbl]] });
      vals.push({ range: `${S}!${colL(i*3)}5`, values: [[CARD_FMLS[i]]] });
    }
  });
  // Currency format on repair cost and exposure
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 18, 21), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Spacer row 5 ──────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column headers (row 6, 0-indexed) ────────────────────────────────────
  const COLS = ['Asset ID','Property ID','Asset / System Name','Asset Type','Home Area','Manufacturer','Model','Serial # Mask','Purchase Date','Purchase Cost','Warranty (Mo.)','Warranty Exp.','Warranty Status','Est. Useful Life Yrs','Est. Repl. Year','Condition','Last Service','Next Service','Provider','Total Repair Cost','Total Maint. Cost','Repl. Cost Est.','Repl. Planning?','Active?','Notes'];
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy,verticalAlignment)' } });
  vals.push({ range: `${S}!A7`, values: [COLS] });

  // ── Data rows (rows 7-506, 0-indexed = 8-507) ─────────────────────────────
  const NDATA = 500;

  // Write asset data FIRST so formula column writes below override the blanks
  ASSETS.forEach((row, i) => {
    const r1 = 8 + i;
    vals.push({ range: `${S}!B${r1}:Y${r1}`, values: [row.slice(1)] });
  });

  // Asset ID formula
  const astIdFmls = Array.from({ length: NDATA }, (_, i) => [`=IF(C${8+i}="","","AST-"&TEXT(ROW()-7,"0000"))`]);
  vals.push({ range: `${S}!A8:A507`, values: astIdFmls });

  // Warranty Expiration formula
  const wexpFmls = Array.from({ length: NDATA }, (_, i) => [`=IF(OR(I${8+i}="",K${8+i}=""),"",EDATE(I${8+i},K${8+i}))`]);
  vals.push({ range: `${S}!L8:L507`, values: wexpFmls });

  // Warranty Status formula
  const wstaFmls = Array.from({ length: NDATA }, (_, i) => [`=IF(L${8+i}="","Unknown",IF(L${8+i}<TODAY(),"Expired",IF(L${8+i}-TODAY()<=90,"Expiring Soon","Active")))`]);
  vals.push({ range: `${S}!M8:M507`, values: wstaFmls });

  // Estimated Replacement Year
  const replFmls = Array.from({ length: NDATA }, (_, i) => [`=IF(OR(I${8+i}="",N${8+i}=""),"",YEAR(I${8+i})+N${8+i})`]);
  vals.push({ range: `${S}!O8:O507`, values: replFmls });

  // Total Repair Cost (from Repair Log by Asset ID)
  const repCostFmls = Array.from({ length: NDATA }, (_, i) => [`=IFERROR(SUMIF(${REP}!$C$8:$C$2007,A${8+i},${REP}!$Q$8:$Q$2007),0)`]);
  vals.push({ range: `${S}!T8:T507`, values: repCostFmls });

  // Total Maintenance Cost (from Costs by Asset ID)
  const maintCostFmls = Array.from({ length: NDATA }, (_, i) => [`=IFERROR(SUMIF(${CST}!$G$8:$G$3007,A${8+i},${CST}!$M$8:$M$3007),0)`]);
  vals.push({ range: `${S}!U8:U507`, values: maintCostFmls });

  // ── Row styling ────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  for (let r = 0; r < NDATA; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, 7+r, 8+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // Formula columns: A(0), L(11), M(12), O(14), T(19), U(20)
  [0, 11, 12, 14, 19, 20].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Input cells: others
  [1,2,3,4,5,6,7,8,9,10,13,15,16,17,18,21,22,23,24].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Number formats
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA,  8,  9), cell: { userEnteredFormat: { numberFormat: { type: 'DATE',     pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Purchase Date
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA,  9, 10), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Purchase Cost
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 11, 12), cell: { userEnteredFormat: { numberFormat: { type: 'DATE',     pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Warranty Exp
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 16, 18), cell: { userEnteredFormat: { numberFormat: { type: 'DATE',     pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Last/Next Service
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 19, 22), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Costs

  // BooleanCondition for Replacement Planning (W=22) and Active (X=23)
  [22, 23].forEach(ci => {
    fmt.push({ setDataValidation: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), rule: { condition: { type: 'BOOLEAN' }, strict: true } } });
  });

  // Dropdowns
  const dvL = (ci, items) => ({ setDataValidation: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), rule: { condition: { type: 'ONE_OF_LIST', values: items.map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });
  fmt.push(dvL(3, ['HVAC System','Furnace','Boiler','Air Conditioner','Heat Pump','Water Heater','Refrigerator','Freezer','Oven / Range','Dishwasher','Microwave','Washer','Dryer','Garbage Disposal','Sump Pump','Water Softener','Well Pump','Septic System','Roof','Gutters','Windows','Exterior Doors','Garage Door','Smoke Detector','Carbon Monoxide Detector','Fire Extinguisher','Electrical Panel','Plumbing Fixture','Irrigation System','Lawn Equipment','Generator','Other']));
  fmt.push(dvL(4, ['Whole Home','Kitchen','Bathroom','Bedroom','Living Room','Dining Room','Laundry','Basement','Attic','Garage','Exterior','Roof','Yard / Landscaping','Driveway','Patio / Deck','HVAC','Plumbing','Electrical','Safety','Other']));
  fmt.push(dvL(15, ['Excellent','Good','Fair','Monitor','Service Recommended','Replacement Planning','Out of Service']));

  // ── Freeze rows 1-7 and cols A-C ─────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row heights ────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 42 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 7+NDATA }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // ── Column widths ──────────────────────────────────────────────────────────
  const COL_WIDTHS = [70,70,180,120,100,90,90,90,90,80,70,90,90,60,60,100,90,90,140,80,80,80,70,60,200];
  COL_WIDTHS.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, fmt, 'app-fmt');
  await valuesBatchUpdate(id, vals, 'app-vals');
  console.log(`✓ Appliances & Systems complete — ${ASSETS.length} assets written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });

function colL(idx) {
  if (idx < 26) return String.fromCharCode(65+idx);
  return String.fromCharCode(64+Math.floor(idx/26)) + String.fromCharCode(65+(idx%26));
}
