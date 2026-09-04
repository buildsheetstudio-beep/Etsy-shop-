'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ACC = sheetMap['🏠 Accommodation & Travel'];

// [name, type, capacity, guestsAssigned, costPerNight, nights, bookingStatus, confirmNum, notes]
const ACCOMMODATION = [
  ['Lake Eildon Cabin — Master Room','Cabin Room',2,'James & Sarah Henderson',0,2,'Fully Booked','CAB-001','Included in venue hire'],
  ['Lake Eildon Cabin — Room 2','Cabin Room',4,'Margaret + Claire + Liam + Emma Henderson',0,2,'Fully Booked','CAB-001','Included in venue hire'],
  ['Lake Eildon Cabin — Room 3','Cabin Room',4,'Diana + Robert Thompson + Lucy Thompson',0,2,'Fully Booked','CAB-001','Included in venue hire'],
  ['Lake Eildon Cabin — Room 4','Cabin Room',3,'Wilson Family (Peter, Anna, Chloe, Noah)',0,2,'Fully Booked','CAB-001','Included in venue hire'],
  ['Overflow — Eildon Hotel Room 1','Hotel Room',2,'Tom & Kate Henderson',180,2,'Deposit Paid','EIL-045','Flies in Aug 15'],
  ['Overflow — Eildon Hotel Room 2','Hotel Room',4,'David + Rachel + Isla Smith / George + Betty Henderson',180,2,'Deposit Paid','EIL-046','Mixed room'],
  ['Local Caravan Site','Caravan',4,'Jack + Mia Thompson + Teen cousins',60,2,'Fully Booked','CAR-012',''],
  ['Self-arranged — AirBnB (local)','Other',6,'Brian Collins + Family Friends group',0,2,'Fully Booked','AIR-789','Own cost'],
];

// [family, fromCity, arrivalDate, arrivalTime, depDate, depTime, transport, pickupNeeded, pickupOrganised, notes]
const TRAVEL = [
  ['Henderson (James & Sarah)','Melbourne, VIC','2025-08-15','2:00 PM','2025-08-16','5:00 PM','Driving',false,'','Organiser — arrives early to set up'],
  ['Henderson (Tom & Kate)','Sydney, NSW','2025-08-15','12:30 PM','2025-08-16','4:00 PM','Flying',true,'James Henderson','MEL flight VA123 — pickup needed'],
  ['Henderson (Margaret & Claire)','Geelong, VIC','2025-08-15','3:00 PM','2025-08-16','3:30 PM','Driving',false,'',''],
  ['Thompson (Diana & Robert)','Brisbane, QLD','2025-08-15','1:45 PM','2025-08-17','9:30 AM','Flying',true,'Peter Wilson','BNE–MEL flight — pickup needed'],
  ['Thompson (Lucy + Jack + Mia)','Brisbane, QLD','2025-08-15','1:45 PM','2025-08-16','4:00 PM','Flying',false,'','Same flight as Diana & Robert'],
  ['Wilson (Peter, Anna, Chloe, Noah)','Melbourne, VIC','2025-08-15','2:30 PM','2025-08-17','10:00 AM','Driving',false,'',''],
  ['Smith (David, Rachel, Isla)','Adelaide, SA','2025-08-15','10:00 AM','2025-08-16','3:00 PM','Driving',false,'','Long drive — arrive early'],
  ['Henderson (George & Betty)','Perth, WA','2025-08-15','3:15 PM','2025-08-17','11:00 AM','Flying',true,'Tom Henderson','PER–MEL flight — pickup needed'],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🏠 Accommodation & Travel'";
  const NCOLS = 10; // A-J

  // Column widths: A-J
  [200,130,80,220,110,70,110,130,130,200].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: ACC, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner
  reqs.push({ mergeCells: { range: gridRange(ACC,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ACC,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  🏠  Accommodation & Travel — Henderson Family Reunion 2025']] });

  // Row 1: Accommodation section header
  reqs.push({ mergeCells: { range: gridRange(ACC,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ACC,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['  🛏️  ACCOMMODATION TRACKER']] });

  // Row 2: Accommodation col headers (GSheets row 3)
  const accHdrs = ['ACCOMMODATION','TYPE','CAPACITY','GUESTS ASSIGNED','COST / NIGHT','NIGHTS','TOTAL COST','BOOKING STATUS','CONFIRMATION #','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(ACC,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [accHdrs] });

  // Accommodation data rows (0-indexed 3-10 = GSheets 4-11)
  for (let i = 0; i < ACCOMMODATION.length; i++) {
    const ri  = 3 + i;
    const row = 4 + i;
    const bg  = i % 2 === 0 ? C.white : C.warmCream;
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    // Currency format cols E, G (indices 4, 6)
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 6, 7),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
    // Center cols C, F (capacity, nights)
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 2, 3),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 5, 6),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    const [name, type, capacity, guests, costPerNight, nights, bookingStatus, confNum, notes] = ACCOMMODATION[i];
    const totalFormula = `=IFERROR(E${row}*F${row},"")`;
    vals.push({ range: `${TAB}!A${row}`, values: [[name, type, capacity, guests, costPerNight, nights, totalFormula, bookingStatus, confNum, notes]] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: 3, endIndex: 3+ACCOMMODATION.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Borders accommodation
  reqs.push({ updateBorders: {
    range: gridRange(ACC,2,3+ACCOMMODATION.length,0,NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // ── TRAVEL TRACKER SECTION ──
  const trvDivRi = 3 + ACCOMMODATION.length; // 0-indexed = 11

  // Amber divider
  reqs.push({ repeatCell: {
    range: gridRange(ACC,trvDivRi,trvDivRi+1,0,NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmAmber) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: trvDivRi, endIndex: trvDivRi+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // Travel section header
  const trvSubRi = trvDivRi + 1;
  reqs.push({ mergeCells: { range: gridRange(ACC,trvSubRi,trvSubRi+1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ACC,trvSubRi,trvSubRi+1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: trvSubRi, endIndex: trvSubRi+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A${trvSubRi+1}`, values: [['  ✈️  TRAVEL TRACKER']] });

  // Travel col headers
  const trvColHdrRi = trvSubRi + 1;
  const trvHdrs = ['FAMILY GROUP','FROM (CITY)','ARRIVAL DATE','ARRIVAL TIME','DEPARTURE DATE','DEPARTURE TIME','TRANSPORT','PICKUP NEEDED?','PICKUP ORGANISED','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(ACC,trvColHdrRi,trvColHdrRi+1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: trvColHdrRi, endIndex: trvColHdrRi+1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A${trvColHdrRi+1}`, values: [trvHdrs] });

  // Travel data rows
  const trvDataStartRi = trvColHdrRi + 1;
  for (let i = 0; i < TRAVEL.length; i++) {
    const ri  = trvDataStartRi + i;
    const row = ri + 1;
    const bg  = i % 2 === 0 ? C.white : C.warmCream;
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    // Date format cols C, E (indices 2, 4)
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 2, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }, backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }, backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    // Time format cols D, F (indices 3, 5)
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 3, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'TIME', pattern: 'h:mm AM/PM' }, horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 5, 6),
      cell: { userEnteredFormat: { numberFormat: { type: 'TIME', pattern: 'h:mm AM/PM' }, horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
    // Center col H (pickup needed)
    reqs.push({ repeatCell: {
      range: gridRange(ACC, ri, ri+1, 7, 8),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(horizontalAlignment,backgroundColor)',
    }});
    const [family, fromCity, arrDate, arrTime, depDate, depTime, transport, pickupNeeded, pickupOrg, notes] = TRAVEL[i];
    vals.push({ range: `${TAB}!A${row}`, values: [[family, fromCity, arrDate, arrTime, depDate, depTime, transport, pickupNeeded ? 'TRUE' : 'FALSE', pickupOrg, notes]] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACC, dimension: 'ROWS', startIndex: trvDataStartRi, endIndex: trvDataStartRi+TRAVEL.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Checkbox for travel col H (pickup needed, index 7)
  reqs.push({ setDataValidation: {
    range: gridRange(ACC, trvDataStartRi, trvDataStartRi+TRAVEL.length, 7, 8),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
  }});

  // Borders travel
  reqs.push({ updateBorders: {
    range: gridRange(ACC,trvColHdrRi,trvDataStartRi+TRAVEL.length,0,NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'accommodation-format');
  await valuesBatchUpdate(id, vals, 'accommodation-values');
  console.log('Accommodation & Travel complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
