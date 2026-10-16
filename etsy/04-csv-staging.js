'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, IMPORT_STATUSES, INCOME_SOURCES } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Etsy CSV Staging'];
const S  = "'Etsy CSV Staging'";

// Etsy CSV columns (standard Etsy order export):
// A=Order ID, B=Order Date, C=Buyer, D=Item Title, E=SKU, F=Quantity, G=Item Price,
// H=Shipping Price, I=Coupon Code, J=Coupon Discount, K=Order Value,
// L=Transaction Fee, M=Payment Processing Fee, N=Refund, O=Net Revenue, P=Import Status,
// Q=Notes, R=Year(auto), S=Month(auto), T=Quarter(auto)

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 2, endIndex: 60 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Column widths
  const colW = [110,100,130,240,100,70,100,90,110,110,100,110,120,90,110,120,160,60,90,60];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:T1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Header row
  reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, 0, 20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});

  // Data rows alternating
  for (let r = 2; r < 60; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 20), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Formula/auto cols R,S,T (17,18,19) — formula bg
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 60, 17, 20), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' }});

  // Import Status col P (15) — input highlight
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 60, 15, 16), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' }});

  // Date format B
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 60, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Currency G,H,J,K,L,M,N,O (6,7,9,10,11,12,13,14)
  [6,7,9,10,11,12,13,14].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 60, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Center: A,B,F,R,S,T
  [0,1,5,17,18,19].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 1, 60, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 1, 60, 0, 20),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze rows only (no column freeze — merged title row spans all cols, can't mix)
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'staging-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['📥 Etsy CSV Staging — Import, Review & Map Your Etsy Order Exports']] });
  data.push({ range: `${S}!A2:T2`, values: [['Order ID','Order Date','Buyer','Item Title','SKU','Qty','Item Price','Shipping','Coupon Code','Coupon Disc.','Order Value','Etsy Trans. Fee','Pay Process. Fee','Refund','Net Revenue','Import Status','Notes','Year','Month','Quarter']] });

  // 50 sample staging rows — a realistic Etsy CSV import
  const TITLES = [
    '2026 Printable Annual Planner','Botanical Wedding Invitation Suite','Minimalist Daily Planner Pad',
    'Watercolor Floral Wall Art Bundle','Custom Business Card Template','Self-Care Journal Printable',
    'Kids Birthday Party Printable Pack','Gradient Abstract Art Print','Recipe Card Template Set',
    'Yoga & Wellness Planner','Budget Tracker Spreadsheet','Mushroom Forest Watercolor Print',
    'Baby Shower Games Bundle','Custom Wedding Day Itinerary','90-Day Goal Setting Workbook',
    'Tropical Botanical Print Trio','Weekly Meal Planning Bundle','Boho Nursery Wall Art Set',
    'Teacher Planner Bundle 2026','Gratitude Journal 365 Prompts',
  ];
  const BUYERS = ['jennifer_m','craft_lover_uk','sarah_k22','printlover99','designfan2','art_studio_nyc','happymom88','teacherlife','plannerqueen','bohovibes'];
  const STATUSES = ['Imported','Imported','Imported','Review Needed','Imported'];

  const sampleRows = [];
  let orderId = 3200100;
  let dateBase = new Date('2026-01-05');

  for (let i = 0; i < 50; i++) {
    const daysAhead = Math.floor(i * 3.5);
    const dt = new Date(dateBase.getTime() + daysAhead * 86400000);
    const dateStr = dt.toISOString().slice(0,10);
    const title = TITLES[i % TITLES.length];
    const isPOD = title.includes('Daily Planner Pad') || title.includes('Abstract Art') || title.includes('Tropical');
    const price = isPOD ? [14.99,16.99,19.99][i%3] : [4.99,7.99,9.99,12.99,14.99][i%5];
    const qty = i % 8 === 0 ? 2 : 1;
    const shipping = (i % 4 === 0) ? 0 : 0;
    const couponDisc = (i % 7 === 0) ? +(price * 0.1).toFixed(2) : 0;
    const orderVal = +(price * qty - couponDisc).toFixed(2);
    const transFee = +(orderVal * 0.065).toFixed(2);
    const processFee = +(orderVal * 0.03 + 0.25).toFixed(2);
    const refund = (i === 12 || i === 31) ? -price : 0;
    const net = +(orderVal - transFee - processFee + refund).toFixed(2);
    const status = STATUSES[i % STATUSES.length];
    const mo = dt.getMonth() + 1;
    const q = mo <= 3 ? 'Q1' : mo <= 6 ? 'Q2' : mo <= 9 ? 'Q3' : 'Q4';

    sampleRows.push([
      `EP${orderId+i}`, dateStr, BUYERS[i%BUYERS.length], title, `WPS-${1001+(i%20)}`,
      qty, price, shipping, (i%7===0?'SAVE10':''), couponDisc, orderVal,
      transFee, processFee, refund, net, status, '',
      2026, MONTHS_SHORT(mo), q,
    ]);
  }

  function MONTHS_SHORT(m) {
    return ['January','February','March','April','May','June','July','August','September','October','November','December'][m-1];
  }

  data.push({ range: `${S}!A3:T52`, values: sampleRows });

  await valuesBatchUpdate(id, data, 'staging-values');

  // Data validation for Import Status
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 2, 60, 15, 16),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$H$2:$H${1+IMPORT_STATUSES.length}` }] }, showCustomUi: true, strict: true },
  }});
  await batchUpdate(id, dvReqs, 'staging-validation');

  console.log('Etsy CSV Staging complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
