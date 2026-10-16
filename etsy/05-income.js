'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C,
        INCOME_SOURCES, INCOME_STATUSES, PRODUCT_TYPES, LISTINGS, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Income & Orders'];
const S  = "'Income & Orders'";

// Columns:
// A=Order ID  B=Date  C=Year(f)  D=Month(f)  E=Quarter(f)  F=Source
// G=Item Title  H=Gross Revenue  I=Etsy Trans Fee(f)  J=Listing Fee
// K=Payment Processing Fee(f)  L=Offsite Ads Fee  M=Refund  N=Net Revenue(f)
// O=Product Type  P=Status  Q=Customer  R=Shipping Collected  S=Notes

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 5 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 5, endIndex: 500 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Column widths
  const colW = [100,100,50,90,55,130,230,100,100,90,120,100,90,100,130,100,120,100,180];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:S1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 19), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 19), cell: { userEnteredFormat: {
    backgroundColor: hex(C.success),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 17 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Summary cards row 2-4 (indices 1-3)
  // Left card: Total Revenue | Right card: Net Revenue | Middle: YTD summary
  [[0,4],[5,10],[11,19]].forEach(([c0,c1], idx) => {
    const colors = [C.success, C.primary, C.secondary];
    reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(colors[idx]),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 2, 3, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { foregroundColor: hex(colors[idx]), bold: true, fontSize: 20 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 3, 4, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 3, 4, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });

  // Header row (row 5, index 4)
  reqs.push({ repeatCell: { range: gridRange(SH, 4, 5, 0, 19), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});

  // Data rows alternating
  for (let r = 5; r < 500; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 19), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Formula columns C,D,E,I,K,N (2,3,4,8,10,13) — formula bg
  [2,3,4,8,10,13].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 500, c, c+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Date format B
  reqs.push({ repeatCell: { range: gridRange(SH, 5, 500, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Currency H,I,J,K,L,M,N,R (7,8,9,10,11,12,13,17)
  [7,8,9,10,11,12,13,17].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 500, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Summary card currency format
  [[2,3,0,4],[2,3,5,10],[2,3,11,19]].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ repeatCell: { range: gridRange(SH, r1, r2, c1, c2), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Center: A,B,C,D,E,O,P (0,1,2,3,4,14,15)
  [0,1,2,3,4,14,15].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 4, 500, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 4, 500, 0, 19),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze rows 5
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'income-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['💰 Income & Orders — All Revenue Transactions']] });

  // Summary cards
  data.push({ range: `${S}!A2`, values: [['  Gross Revenue (2026)']] });
  data.push({ range: `${S}!A3`, values: [['=SUMPRODUCT((YEAR(B$6:B$500)=2026)*H$6:H$500)']] });
  data.push({ range: `${S}!A4`, values: [['All income sources · fiscal year to date']] });

  data.push({ range: `${S}!F2`, values: [['  Net Revenue (2026)']] });
  data.push({ range: `${S}!F3`, values: [['=SUMPRODUCT((YEAR(B$6:B$500)=2026)*N$6:N$500)']] });
  data.push({ range: `${S}!F4`, values: [['After Etsy fees · refunds deducted']] });

  data.push({ range: `${S}!L2`, values: [['  Total Orders (2026)']] });
  data.push({ range: `${S}!L3`, values: [['=COUNTIFS(B$6:B$500,">="&DATE(2026,1,1),B$6:B$500,"<="&DATE(2026,12,31),P$6:P$500,"<>Refunded")']] });
  data.push({ range: `${S}!L4`, values: [['Etsy orders + off-Etsy combined']] });

  // Column headers (row 5)
  data.push({ range: `${S}!A5:S5`, values: [['Order ID','Date','Year','Month','Quarter','Source','Item Title','Gross Revenue','Trans. Fee (6.5%)','Listing Fee','Pay Processing Fee','Offsite Ads Fee','Refund','Net Revenue','Product Type','Status','Customer','Shipping Collected','Notes']] });

  // 45 sample income rows across Jan-Jul 2026
  const rows = [];
  const BUYERS = ['jennifer_m','craft_lover_uk','sarah_k22','printlover99','designfan2','art_studio_nyc','happymom88','teacherlife','plannerqueen','bohovibes','nature_lover','studioflo','sunnydesigns','paperco_au'];
  let oid = 100001;

  for (let i = 0; i < 45; i++) {
    const mo = Math.floor(i / 6.5) + 1; // spread across Jan-Jul
    const day = (i % 28) + 1;
    const dateStr = `2026-${String(mo).padStart(2,'0')}-${String(Math.min(day,28)).padStart(2,'0')}`;
    const listing = LISTINGS[i % LISTINGS.length];
    const gross = listing.type === 'Custom Order' ? [35,45,55][i%3] :
                  listing.pod > 0 ? listing.pod * [2.5,2.8,3.2][i%3] + 3.99 :
                  [4.99,7.99,9.99,12.99,14.99,6.99][i%6];
    const grossR = +gross.toFixed(2);
    const transFee = +(grossR * 0.065).toFixed(2);
    const listFee  = (i % 15 === 0) ? 0.20 : 0;
    const procFee  = +(grossR * 0.03 + 0.25).toFixed(2);
    const offsiteAds = (i % 9 === 0) ? +(grossR * 0.15).toFixed(2) : 0;
    const refund   = (i === 11 || i === 28) ? -grossR : 0;
    const net      = +(grossR - transFee - listFee - procFee - offsiteAds + refund).toFixed(2);
    const status   = refund !== 0 ? 'Refunded' : 'Received';
    const source   = 'Etsy Order';
    const q        = mo <= 3 ? 'Q1' : mo <= 6 ? 'Q2' : 'Q3';

    rows.push([
      `EP${oid+i}`, dateStr, 2026, MONTHS[mo-1], q, source, listing.title,
      grossR, transFee, listFee, procFee, offsiteAds, refund, net,
      listing.type, status, BUYERS[i%BUYERS.length], 0, '',
    ]);
  }
  data.push({ range: `${S}!A6:S50`, values: rows });

  // Formula columns for all rows 6-500
  for (let r = 6; r <= 55; r++) {
    // C=Year, D=Month, E=Quarter (auto-calc from B)
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(YEAR(B${r}),"")`]] });
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(TEXT(B${r},"mmmm"),"")`]] });
    data.push({ range: `${S}!E${r}`, values: [[`=IFERROR(IF(MONTH(B${r})<=3,"Q1",IF(MONTH(B${r})<=6,"Q2",IF(MONTH(B${r})<=9,"Q3","Q4"))),"")`]] });
    // I=Transaction Fee (6.5%)
    data.push({ range: `${S}!I${r}`, values: [[`=IFERROR(ROUND(H${r}*0.065,2),"")`]] });
    // K=Payment Processing Fee (3% + $0.25)
    data.push({ range: `${S}!K${r}`, values: [[`=IFERROR(ROUND(H${r}*0.03+0.25,2),"")`]] });
    // N=Net Revenue
    data.push({ range: `${S}!N${r}`, values: [[`=IFERROR(H${r}-I${r}-J${r}-K${r}-L${r}+M${r},"")`]] });
  }

  await valuesBatchUpdate(id, data, 'income-values');

  // Data validations
  const dvReqs = [];
  // F=Source
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 5, 500, 5, 6),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$B$2:$B${1+INCOME_SOURCES.length}` }] }, showCustomUi: true, strict: false },
  }});
  // O=Product Type
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 5, 500, 14, 15),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$E$2:$E${1+PRODUCT_TYPES.length}` }] }, showCustomUi: true, strict: false },
  }});
  // P=Status
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 5, 500, 15, 16),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$F$2:$F${1+INCOME_STATUSES.length}` }] }, showCustomUi: true, strict: true },
  }});
  await batchUpdate(id, dvReqs, 'income-validation');

  console.log('Income & Orders complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
