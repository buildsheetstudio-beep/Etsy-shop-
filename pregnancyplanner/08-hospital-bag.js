'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const HB = sheetMap['🎒 Hospital Bag Checklist'];

// 6 cols, 75 rows
// Row 0: Title, Row 1: Summary (progress bar)
// Row 2: Summary values
// Row 3: Col headers
// Row 4+: items by section

(async () => {
  const reqs = [];

  // Title
  reqs.push({ mergeCells: { range: gridRange(HB, 0, 1, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(HB, 0, 1, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: HB, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Summary rows 1-2
  reqs.push({ repeatCell: {
    range: gridRange(HB, 1, 2, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(HB, 2, 3, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: HB, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Col headers row 3
  reqs.push({ repeatCell: {
    range: gridRange(HB, 3, 4, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: HB, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Alt rows 4-74
  for (let r = 4; r < 75; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(HB, r, r+1, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Wrap
  reqs.push({ repeatCell: { range: gridRange(HB, 4, 75, 0, 6), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});

  // Column widths
  const colW = [160, 90, 90, 180, 100, 180];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: HB, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(HB, 3, 75, 0, 6), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});

  await batchUpdate(id, reqs, 'hb-format');

  const vdata = [];
  vdata.push({ range: `'🎒 Hospital Bag Checklist'!A1`, values: [['🎒 Hospital Bag Checklist']] });

  vdata.push({ range: `'🎒 Hospital Bag Checklist'!A2:F2`, values: [['Total Items','Packed','Still Needed','% Packed','Progress','Notes']] });
  vdata.push({ range: `'🎒 Hospital Bag Checklist'!A3:F3`, values: [[
    '=COUNTA(A5:A74)',
    '=COUNTIF(C5:C74,TRUE)',
    '=A3-B3',
    '=IFERROR(TEXT(B3/A3,"0%"),"—")',
    `=IFERROR(SPARKLINE({B3,A3-B3},{"charttype","bar";"color1","#6B5B8B";"color2","#DDD5E8"}),"")`,
    '',
  ]]});

  vdata.push({ range: `'🎒 Hospital Bag Checklist'!A4:F4`, values: [['Item','Section','Packed','Notes','Qty','Brand / Detail']] });

  // Items grouped by section (A=Item, B=Section, C=Packed checkbox formula placeholder, D=Notes, E=Qty, F=Brand)
  const items = [
    // Labour & Delivery
    ['Birth plan copies','Labour & Delivery',true,'3-4 copies','3',''],
    ['Insurance card & ID','Labour & Delivery',true,'','2',''],
    ['Hospital pre-registration forms','Labour & Delivery',true,'Completed online','1',''],
    ['Phone + charger (long cable)','Labour & Delivery',false,'','1',''],
    ['Portable battery pack','Labour & Delivery',false,'','1',''],
    ['Bluetooth speaker','Labour & Delivery',false,'For birthing playlist','1',''],
    ['Massage tools / tennis balls','Labour & Delivery',false,'For back labour','1',''],
    ['Lip balm','Labour & Delivery',false,'Lips dry during labour','2',''],
    ['Snacks for partner','Labour & Delivery',false,'Granola bars, nuts, etc.','Bag',''],
    ['Water bottle (large)','Labour & Delivery',false,'Straw top preferred','1',''],
    ['Fan / misting spray','Labour & Delivery',false,'Labour gets hot!','1',''],
    ['Hair ties / clips','Labour & Delivery',false,'','Several',''],
    // Postpartum (Mum)
    ['Loose PJs or nightgown (nursing-friendly)','Postpartum (Mum)',false,'','2',''],
    ['Robe','Postpartum (Mum)',false,'For skin-to-skin','1',''],
    ['Nursing bra x2','Postpartum (Mum)',false,'','2',''],
    ['Disposable underwear x10','Postpartum (Mum)',false,'Or high-waist cotton','10','Frida Mom'],
    ['Peri bottle','Postpartum (Mum)',false,'','1','Frida Mom'],
    ['Witch hazel pads / Tucks','Postpartum (Mum)',false,'','1 pack',''],
    ['Nipple cream','Postpartum (Mum)',false,'','1','Lansinoh'],
    ['Breast pads','Postpartum (Mum)',false,'','1 pack',''],
    ['Socks — warm, non-slip','Postpartum (Mum)',false,'Labour floors are cold!','3',''],
    ['Flip-flops for shower','Postpartum (Mum)',false,'','1',''],
    ['Toiletries (travel size)','Postpartum (Mum)',false,'Shampoo, conditioner, body wash','Bag',''],
    ['Dry shampoo','Postpartum (Mum)',false,'','1',''],
    ['Comfy going-home outfit','Postpartum (Mum)',false,'Loose, bump-friendly','1',''],
    ['Glasses / contacts supplies','Postpartum (Mum)',false,'','1',''],
    ['Pillow from home','Postpartum (Mum)',false,'Hospital pillows are flat!','1',''],
    // Postpartum (Baby)
    ['Newborn outfit x2','Postpartum (Baby)',false,'','2',''],
    ['Onesies + footie PJs x2','Postpartum (Baby)',false,'','2',''],
    ['Baby hat (soft)','Postpartum (Baby)',false,'','1',''],
    ['Baby mittens x2','Postpartum (Baby)',false,'Prevent scratches','2',''],
    ['Swaddle blankets x2','Postpartum (Baby)',false,'','2','HALO'],
    ['Muslin receiving blanket','Postpartum (Baby)',false,'','2',''],
    ['Going-home outfit (special!)','Postpartum (Baby)',false,'','1',''],
    ['Car seat (installed & inspected)','Postpartum (Baby)',false,'Required to leave hospital','1','Britax'],
    // Partner Bag
    ['Change of clothes x2','Partner Bag',false,'','2',''],
    ['Toiletries','Partner Bag',false,'','1',''],
    ['Snacks & drinks','Partner Bag',false,'','Bag',''],
    ['Pillow + blanket','Partner Bag',false,'','1',''],
    ['Camera + extra battery','Partner Bag',false,'','1',''],
    ['Headphones','Partner Bag',false,'','1',''],
    ['Cash + parking money','Partner Bag',false,'','$50',''],
    // Documents
    ['Photo ID','Documents',true,'','2',''],
    ['Insurance card','Documents',true,'','1',''],
    ['Paediatrician info','Documents',false,'Name + number','1',''],
    ['Cord blood kit (if ordered)','Documents',false,'','1',''],
    ['Hospital pre-admission paperwork','Documents',true,'','1',''],
    // Going Home
    ['Car seat base (installed)','Going Home',false,'Get checked at fire station','1',''],
    ['Coming-home outfit (mum)','Going Home',false,'Loose & comfy','1',''],
    ['Baby\'s coming-home outfit','Going Home',false,'Weather appropriate','1',''],
    ['Warm blanket for baby','Going Home',false,'Season dependent','1',''],
    ['Thank you cards for nurses','Going Home',false,'Optional but thoughtful','5',''],
  ];
  vdata.push({ range: `'🎒 Hospital Bag Checklist'!A5:F${4+items.length}`, values: items });

  await valuesBatchUpdate(id, vdata, 'hb-values');
  console.log('Hospital Bag complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
