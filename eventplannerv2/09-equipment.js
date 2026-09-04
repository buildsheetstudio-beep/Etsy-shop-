'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 7;

// 21 columns A:U
const HEADERS = [
  'Item ID','Event ID','Item Name','Category','Owner/Dept','Qty Needed','Qty Available',
  'Qty Ordered','Status','Unit Cost','Total Cost','Vendor/Source',
  'Order Date','Expected Delivery','Actual Delivery','Return Required?',
  'Returned?','Storage Location','Condition','Shortage Flag','Notes',
];

const COL_WIDTHS = [75,80,200,120,120,80,80,80,110,80,90,140,100,110,110,90,75,130,100,90,200];

// 60 equipment records
const EQUIPMENT = [
  // AV Equipment
  ['EQ-001','EVT-001','Main Stage PA System','AV/Tech','AV Team',1,0,1,'Ordered',3500,'','VDR-003','03/01/2026','04/10/2026','','No','','ACC Storage','New','','Conference main stage'],
  ['EQ-002','EVT-001','LED Stage Lighting Rig','AV/Tech','AV Team',1,0,1,'Ordered',5000,'','VDR-003','03/01/2026','04/10/2026','','No','','ACC Storage','New','',''],
  ['EQ-003','EVT-001','4x Wireless Lapel Mics','AV/Tech','AV Team',4,0,4,'Ordered',800,'','VDR-003','03/01/2026','04/10/2026','','No','','AV Cart','New','',''],
  ['EQ-004','EVT-001','Projection Screens (Main Hall)','AV/Tech','AV Team',3,0,3,'Ordered',2400,'','VDR-003','03/01/2026','04/10/2026','','No','','ACC Storage','New','',''],
  ['EQ-005','EVT-001','Laptop & Presenter Clicker','AV/Tech','AV Team',2,2,0,'Available',400,'','Foundation HQ','','','','No','','Staff Bag','Good','','Foundation owned'],
  ['EQ-006','EVT-001','Live Streaming Encoder','AV/Tech','AV Team',1,0,1,'Ordered',1500,'','VDR-003','03/01/2026','04/10/2026','','No','','AV Cart','New','','Zoom + YouTube'],
  ['EQ-007','EVT-002','Wireless Handheld Mics (x2)','AV/Tech','AV Team',2,0,2,'Ordered',600,'','VDR-003','','04/25/2026','','Yes','','Driskill AV Room','New','','Return after event'],
  ['EQ-008','EVT-002','Uplighting Package (Ballroom)','AV/Tech','AV Team',20,0,20,'Ordered',2000,'','VDR-003','','04/25/2026','','Yes','','Ballroom','New','',''],
  ['EQ-009','EVT-003','Portable PA Speaker (Outdoor)','AV/Tech','AV Team',4,0,4,'Ordered',1200,'','VDR-004','04/01/2026','05/10/2026','','Yes','','Event Trailer','New','','Generator-powered'],
  ['EQ-010','EVT-003','Generator (10kW)','AV/Tech','AV Team',2,0,2,'Ordered',1800,'','VDR-004','04/01/2026','05/10/2026','','Yes','','Event Trailer','New','',''],
  // Signage & Printing
  ['EQ-011','EVT-001','Entry Arch Banner','Signage','Ops Team',2,0,2,'Ordered',400,'','VDR-009','03/10/2026','04/05/2026','','No','','Event Trailer','New','',''],
  ['EQ-012','EVT-001','Directional Signage Set (12 signs)','Signage','Ops Team',12,0,12,'Ordered',600,'','VDR-009','03/10/2026','04/05/2026','','No','','Ops Box','New','',''],
  ['EQ-013','EVT-001','Stage Backdrop — Custom','Signage','Ops Team',1,0,1,'Ordered',1200,'','VDR-009','03/10/2026','04/05/2026','','No','','ACC Storage','New','',''],
  ['EQ-014','EVT-001','Name Badge Holders (x600)','Supplies','Registration',600,600,0,'Available',90,'','Office Supply','','','','No','','Registration Box','Good','','Reusing from last year'],
  ['EQ-015','EVT-001','Lanyards (x600)','Supplies','Registration',600,0,600,'Ordered',120,'','VDR-009','03/15/2026','04/05/2026','','No','','Registration Box','New','',''],
  ['EQ-016','EVT-001','Printed Programs (x500)','Printed','Registration',500,0,500,'Ordered',750,'','VDR-009','03/25/2026','04/10/2026','','No','','Registration Box','New','',''],
  ['EQ-017','EVT-001','Speaker Gift Bags (x10)','Supplies','Protocol',10,0,10,'Ordered',500,'','VDR-009','04/01/2026','04/12/2026','','No','','Staff Office','New','','Custom items'],
  ['EQ-018','EVT-002','Auction Bid Paddles (x300)','Supplies','Ops Team',300,0,0,'Not Started',150,'','Local Print','','','','No','','','','','Order pending'],
  ['EQ-019','EVT-002','Auction Item Bidsheet Set','Printed','Ops Team',60,0,0,'Not Started',60,'','Local Print','','','','No','','','','','60 items expected'],
  // Furniture & Décor
  ['EQ-020','EVT-001','Round Banquet Tables (x50)','Furniture','Venue',50,50,0,'Available',0,'','VEN-001','','','','No','','ACC Warehouse','Good','','Included in venue rental'],
  ['EQ-021','EVT-001','Banquet Chairs (x600)','Furniture','Venue',600,600,0,'Available',0,'','VEN-001','','','','No','','ACC Warehouse','Good','','Included in venue rental'],
  ['EQ-022','EVT-001','Cocktail Tables (x20)','Furniture','Venue',20,20,0,'Available',0,'','VEN-001','','','','No','','ACC Foyer','Good','','Included in venue'],
  ['EQ-023','EVT-002','Floral Centerpieces (x30)','Décor','Décor Team',30,0,30,'Ordered',2550,'','VDR-006','04/01/2026','05/01/2026','','No','','Ballroom','New','','Gold & blush'],
  ['EQ-024','EVT-002','Pipe & Drape — Ballroom','Décor','Décor Team',1,0,1,'Ordered',3000,'','VDR-006','04/01/2026','05/01/2026','','No','','Ballroom','New','',''],
  ['EQ-025','EVT-002','Candle & Votive Package','Décor','Décor Team',150,0,150,'Ordered',450,'','VDR-006','','05/01/2026','','No','','Ballroom','New','',''],
  // Medical & Safety
  ['EQ-026','EVT-001','First Aid Kit (x5)','Medical','Safety',5,5,0,'Available',200,'','Foundation HQ','','','','No','','Registration Area','Good','','AED also on-site'],
  ['EQ-027','EVT-001','AED Defibrillator','Medical','Safety',1,1,0,'Available',0,'','ACC','','','','No','','Main Entrance','Good','','Venue-provided'],
  ['EQ-028','EVT-003','Blood Pressure Monitor (x10)','Medical','Health Partners',10,0,10,'Ordered',800,'','VDR-016','04/15/2026','05/10/2026','','Yes','','Health Booth','New','','Return to partners'],
  ['EQ-029','EVT-003','Glucose Test Kits (x200)','Medical','Health Partners',200,0,200,'Ordered',300,'','VDR-016','04/15/2026','05/10/2026','','No','','Health Supplies Box','New','','Consumable'],
  ['EQ-030','EVT-003','Nurse Station Canopy (x10)','Furniture','Health Partners',10,0,10,'Ordered',2000,'','VDR-017','04/15/2026','05/10/2026','','Yes','','Event Trailer','New','',''],
  // Technology
  ['EQ-031','EVT-001','Check-In iPad Stations (x6)','Technology','Registration',6,4,2,'Ordered',0,'','Foundation HQ','','04/12/2026','','No','','Registration Table','Good','','2 on order'],
  ['EQ-032','EVT-001','QR Code Scanner (x4)','Technology','Registration',4,4,0,'Available',0,'','Foundation HQ','','','','No','','Registration Table','Good','',''],
  ['EQ-033','EVT-007','Virtual Event Platform License','Technology','IT Team',1,1,0,'Available',1500,'','VDR-019','','','','No','','N/A — Cloud','Active','','Annual license'],
  ['EQ-034','EVT-007','Host Laptop — Video Production','Technology','IT Team',1,1,0,'Available',0,'','Foundation HQ','','','','No','','Virtual Studio','Good','',''],
  ['EQ-035','EVT-013','Streaming Server Rack','Technology','AV Team',1,0,1,'Ordered',8000,'','VDR-003','','07/01/2026','','Yes','','Palmer AV Room','New','',''],
  // Sports Equipment
  ['EQ-036','EVT-006','Soccer Goals (x8)','Sports','Sports Team',8,0,8,'Ordered',1600,'','Sports Supply Co','04/01/2026','06/01/2026','','Yes','','Field Storage','New','',''],
  ['EQ-037','EVT-006','Soccer Balls (x24)','Sports','Sports Team',24,0,24,'Ordered',480,'','Sports Supply Co','04/01/2026','06/01/2026','','No','','Equipment Bag','New','',''],
  ['EQ-038','EVT-006','Corner Flags (x32)','Sports','Sports Team',32,0,32,'Ordered',160,'','Sports Supply Co','04/01/2026','06/01/2026','','Yes','','Field Storage','New','',''],
  ['EQ-039','EVT-006','Referee Whistles & Cards','Sports','Sports Team',8,8,0,'Available',40,'','Foundation HQ','','','','No','','Equipment Bag','Good','',''],
  ['EQ-040','EVT-006','Score Boards (x4)','Sports','Sports Team',4,0,4,'Ordered',800,'','Sports Supply Co','04/01/2026','06/01/2026','','Yes','','Field Storage','New','','Return after event'],
  // Catering & Kitchen
  ['EQ-041','EVT-001','Portable Coffee Station','Catering','Catering',3,0,3,'Ordered',900,'','VDR-001','04/01/2026','04/12/2026','','Yes','','Break Room','New','',''],
  ['EQ-042','EVT-001','Water Station (x8)','Catering','Catering',8,0,8,'Ordered',400,'','VDR-001','04/01/2026','04/12/2026','','Yes','','Hallways','New','',''],
  ['EQ-043','EVT-001','Buffet Warmers (x12)','Catering','Catering',12,0,12,'Ordered',600,'','VDR-001','04/01/2026','04/12/2026','','Yes','','Kitchen','New','',''],
  ['EQ-044','EVT-002','Silver Serving Trays (x40)','Catering','Catering',40,0,40,'Ordered',1200,'','VDR-002','','04/25/2026','','Yes','','Ballroom Kitchen','New','',''],
  ['EQ-045','EVT-003','Disposable Cups & Plates (x2000)','Catering','Catering',2000,0,2000,'Ordered',200,'','Restaurant Supply','04/15/2026','05/05/2026','','No','','Supply Box','New','',''],
  // Tent & Staging
  ['EQ-046','EVT-003','10x10 Vendor Tents (x30)','Tent','Ops Team',30,0,30,'Ordered',4500,'','VDR-017','04/15/2026','05/12/2026','','Yes','','Park Storage','New','',''],
  ['EQ-047','EVT-003','Main Stage Platform (20x30)','Staging','Ops Team',1,0,1,'Ordered',3000,'','VDR-017','04/15/2026','05/12/2026','','Yes','','Park Center','New','',''],
  ['EQ-048','EVT-011','Pop-Up Canopies (x10)','Tent','Ops Team',10,0,0,'Not Started',800,'','VDR-017','','','','Yes','','','','','Order pending'],
  ['EQ-049','EVT-019','Cultural Display Booths (x30)','Staging','Ops Team',30,0,30,'Ordered',3000,'','VDR-024','','08/10/2026','','Yes','','Republic Square','New','',''],
  ['EQ-050','EVT-019','Main Stage (40x60)','Staging','Ops Team',1,0,1,'Ordered',6000,'','VDR-024','','08/10/2026','','Yes','','Republic Square','New','',''],
  // Miscellaneous
  ['EQ-051','EVT-001','Extension Cords & Power Strips','Supplies','Ops Team',20,15,5,'In Use',150,'','Foundation HQ','','04/10/2026','','No','','Ops Box','Good','',''],
  ['EQ-052','EVT-001','Tape, Scissors, Markers Kit','Supplies','Ops Team',1,1,0,'Available',50,'','Foundation HQ','','','','No','','Ops Box','Good','',''],
  ['EQ-053','EVT-001','Event Radios (x10)','Communication','Staff',10,10,0,'Available',500,'','Foundation HQ','','','','No','','Charging Station','Good','',''],
  ['EQ-054','EVT-001','Volunteer Vests (x30)','Supplies','Staff',30,25,5,'In Use',300,'','Online Order','03/01/2026','04/01/2026','','No','','Storage Closet','Good','',''],
  ['EQ-055','EVT-002','Valet Ticket Books','Supplies','Ops Team',5,0,5,'Ordered',25,'','VDR-015','','04/25/2026','','No','','Valet Booth','New','',''],
  ['EQ-056','EVT-004','Participant Workbooks (x45)','Printed','Training',45,0,45,'Ordered',225,'','VDR-009','04/01/2026','05/15/2026','','No','','Training Room','New','',''],
  ['EQ-057','EVT-004','Participant Certificates (x40)','Printed','Training',40,0,0,'Not Started',80,'','Local Print','','','','No','','','','','Template pending'],
  ['EQ-058','EVT-006','Team Award Trophies (x6)','Awards','Ops Team',6,0,6,'Ordered',600,'','VDR-018','04/01/2026','06/01/2026','','No','','Event Trailer','New','',''],
  ['EQ-059','EVT-006','Individual Medals (x48)','Awards','Ops Team',48,0,48,'Ordered',480,'','VDR-018','04/01/2026','06/01/2026','','No','','Event Trailer','New','',''],
  ['EQ-060','EVT-022','Sponsor Golf Cart Wraps (x10)','Marketing','Golf Team',10,0,0,'Not Started',500,'','Print Shop','','','','No','','','','','Design pending'],
];

(async () => {
  const values = [];
  const fmt = [];

  // Title
  values.push({ range:"'Equipment & Supplies'!A1", values:[['EQUIPMENT & SUPPLIES']] });
  values.push({ range:"'Equipment & Supplies'!A2", values:[['Track all equipment, supplies, décor, technology, and printed materials. Shortage flags and cost totals are auto-calculated.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,21), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,21), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // Stats row 3
  values.push({ range:"'Equipment & Supplies'!A3", values:[[
    `=IFERROR("Total Items: "&COUNTA(C6:C500)&" | Available: "&COUNTIF(I6:I500,"Available")&" | Ordered: "&COUNTIF(I6:I500,"Ordered")&" | Shortages: "&COUNTIF(T6:T500,"SHORTAGE"),"Loading...")`
  ]]});
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,21), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#CBDCC5'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#2A5C35')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  // Info row 4
  values.push({ range:"'Equipment & Supplies'!A4", values:[['ℹ Total Cost (K), Shortage Flag (T), Return Required (P), and Returned (Q) are formula or checkbox columns.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,3,4,0,21), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,3,4,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#EEF4F8'), textFormat:{italic:true,fontSize:9,foregroundColor:hex('#4A5056')},
    verticalAlignment:'MIDDLE', padding:{left:10},
  }}, fields:'userEnteredFormat' }});

  // Header row 5
  values.push({ range:"'Equipment & Supplies'!A5", values:[HEADERS] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5},
    properties:{pixelSize:32}, fields:'pixelSize' }});

  COL_WIDTHS.forEach((w,i)=>{
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1},
      properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:5}},
    fields:'gridProperties.frozenRowCount' }});

  // Data rows with formulas
  const dataRows = EQUIPMENT.map((eq,i)=>{
    const r = i+6;
    const row = [...eq];
    // K=10: Total Cost = Qty Needed * Unit Cost
    row[10] = `=IFERROR(F${r}*J${r},0)`;
    // P=15: Return Required — keep as-is (Yes/No string → checkbox)
    row[15] = row[15] === 'Yes';
    // Q=16: Returned — checkbox
    row[16] = row[16] === 'Yes';
    // T=19: Shortage Flag
    row[19] = `=IFERROR(IF(AND(F${r}>0,F${r}>(G${r}+H${r})),"SHORTAGE",""),"")`;
    return row;
  });
  values.push({ range:"'Equipment & Supplies'!A6", values: dataRows });

  // Alt row styling
  EQUIPMENT.forEach((_,i)=>{
    const ri = i+5;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,21), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#FDFCF9':'#F2F1EE'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9}, verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Formula cols: K=10, T=19
  EQUIPMENT.forEach((_,i)=>{
    const ri = i+5;
    [10,19].forEach(c=>{
      fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,c,c+1), cell:{userEnteredFormat:{
        backgroundColor:hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374'),fontSize:9},
        horizontalAlignment:'RIGHT',
      }}, fields:'userEnteredFormat' }});
    });
  });

  // Checkboxes: P=15, Q=16
  [15,16].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,5,500,c,c+1), cell:{
      userEnteredFormat:{ horizontalAlignment:'CENTER' },
      dataValidation:{ condition:{ type:'BOOLEAN' }, showCustomUi:true },
    }, fields:'userEnteredFormat.horizontalAlignment,dataValidation' }});
  });

  // Currency: J=9, K=10
  [9,10].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,5,500,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0.00' }
    }}, fields:'userEnteredFormat.numberFormat' }});
  });

  // Validations
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,3,4), rule:{
    condition:{ type:'ONE_OF_LIST', values:['AV/Tech','Signage','Printed','Supplies','Furniture',
      'Décor','Medical','Technology','Sports','Catering','Tent','Staging','Communication',
      'Awards','Marketing','Other'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,8,9), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Available','Ordered','In Use','Not Started','Returned','Damaged','Missing'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});

  // CF: Status (I=col8)
  const statusCF = [
    { v:'Available',   bg:'#CBDCC5', text:'#2A5C35' },
    { v:'Ordered',     bg:'#D5E4F2', text:'#1E3A5F' },
    { v:'In Use',      bg:'#E9DBA7', text:'#5C4A00' },
    { v:'Not Started', bg:'#F2F1EE', text:'#4A5056' },
    { v:'Returned',    bg:'#DDD0DF', text:'#3A1A5C' },
    { v:'Damaged',     bg:'#E8C0B8', text:'#5C1A1A' },
    { v:'Missing',     bg:'#BC7770', text:'#FFFFFF'  },
  ];
  statusCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,500,8,9)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // CF: SHORTAGE flag (T=col19)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,0,21)],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:'=$T6="SHORTAGE"'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});

  // CF: Condition = Damaged/Missing (S=col18)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,18,19)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Damaged'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,18,19)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Missing'}] },
      format:{ backgroundColor:hex('#BC7770'), textFormat:{foregroundColor:hex('#FFFFFF'),bold:true} } }
  }, index:0 }});

  // Filter
  fmt.push({ setBasicFilter:{ filter:{ range: gridRange(SID,4,500,0,21) }}});

  await valuesBatchUpdate(id, values, '09-equipment values');
  await batchUpdate(id, fmt, '09-equipment format');
  console.log('✅ Equipment & Supplies done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
