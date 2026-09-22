'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TCO = sheetMap['Total Cost of Ownership'];
const S   = "'Total Cost of Ownership'";
const VQ  = "'Vehicle Details & Quotes'";
const BS  = "'Buyer Setup'";

const VEHICLES = ['V-001','V-002','V-003','V-004','V-005','V-006','V-007','V-008'];
const colLetters = ['D','E','F','G','H','I','J','K'];

(async () => {
  const data = [];

  data.push({ range: `${S}!A1`, values: [['TOTAL COST OF OWNERSHIP']] });
  data.push({ range: `${S}!A2`, values: [['Project the true cost of owning each vehicle over 1, 3, 5, and 7 years including financing, depreciation, fuel, insurance, and maintenance.']] });
  data.push({ range: `${S}!A3`, values: [['Resale value subtracted once at end of selected horizon. Replace defaults with your actual quotes and local data.']] });

  // Controls
  data.push({ range: `${S}!A5`, values: [['TCO CONTROLS']] });
  data.push({ range: `${S}!A6:B6`, values: [['Ownership Horizon (years)', 5]] });
  data.push({ range: `${S}!A7:B7`, values: [['Annual Miles', `=IFERROR('Buyer Setup'!B21,12000)`]] });
  data.push({ range: `${S}!A8:B8`, values: [['Fuel Price ($/gal)', 3.50]] });
  data.push({ range: `${S}!A9:B9`, values: [['Electricity Price ($/kWh)', 0.16]] });
  data.push({ range: `${S}!A10:B10`, values: [['EV Efficiency (miles/kWh)', 3.5]] });
  data.push({ range: `${S}!A11:B11`, values: [['Depr. Rate — New Yr 1', 0.20]] });
  data.push({ range: `${S}!A12:B12`, values: [['Depr. Rate — New Yr 2+', 0.12]] });
  data.push({ range: `${S}!A13:B13`, values: [['Depr. Rate — Used/CPO', 0.10]] });

  // Vehicle ID row
  data.push({ range: `${S}!A15`, values: [['Vehicle ID']] });
  VEHICLES.forEach((vid, i) => {
    data.push({ range: `${S}!${colLetters[i]}15`, values: [[vid]] });
  });

  data.push({ range: `${S}!A16`, values: [['Make / Model']] });
  VEHICLES.forEach((vid, i) => {
    const col = colLetters[i];
    data.push({ range: `${S}!${col}16`, values: [[
      `=IFERROR(INDEX(${VQ}!$C$6:$C$105,MATCH(${col}15,${VQ}!$A$6:$A$105,0))&" "&INDEX(${VQ}!$D$6:$D$105,MATCH(${col}15,${VQ}!$A$6:$A$105,0)),"")`
    ]] });
  });

  const costLines = [
    'Purchase Price / Cap Cost',
    'Financing Interest',
    'Depreciation (over horizon)',
    'Insurance (total)',
    'Fuel / Electricity (total)',
    'Maintenance (total)',
    'Repairs Estimate (total)',
    'Registration / Taxes (total)',
    'Fees',
    'Warranty / Service Plans',
    'Parking & Tolls (total)',
    'Est. Resale Value',
    'TOTAL Ownership Cost',
    'Average Annual Cost',
    'Average Monthly Cost',
    'Cost per Mile',
  ];

  costLines.forEach((label, i) => {
    data.push({ range: `${S}!A${i+18}`, values: [[label]] });
  });

  VEHICLES.forEach((vid, vi) => {
    const col = colLetters[vi];
    const mBase = `MATCH(${col}15,${VQ}!$A$6:$A$105,0)`;
    const yrs  = `$B$6`;
    const miles = `$B$7`;
    const fp   = `$B$8`;
    const ep   = `$B$9`;
    const eff  = `$B$10`;

    const otd   = `IFERROR(INDEX(${VQ}!$W$6:$W$105,${mBase}),0)`;
    const apr   = `IFERROR(INDEX(${VQ}!$X$6:$X$105,${mBase}),0)`;
    const term  = `IFERROR(INDEX(${VQ}!$Y$6:$Y$105,${mBase}),60)`;
    const dp    = `IFERROR(INDEX(${VQ}!$V$6:$V$105,${mBase}),0)`;
    const tradeV= `IFERROR('Buyer Setup'!$B$19,0)`;
    const tradeL= `IFERROR('Buyer Setup'!$B$20,0)`;
    const amtFin= `MAX(0,${otd}-${dp}-${tradeV}+${tradeL})`;
    const mpg   = `IFERROR(INDEX(${VQ}!$Z$6:$Z$105,${mBase}),28)`;
    const fuel  = `IFERROR(INDEX(${VQ}!$I$6:$I$105,${mBase}),"Gasoline")`;
    const cond  = `IFERROR(INDEX(${VQ}!$G$6:$G$105,${mBase}),"Used")`;
    const insM  = `IFERROR(INDEX(${VQ}!$AA$6:$AA$105,${mBase}),120)`;
    const maintM= `IFERROR(INDEX(${VQ}!$AB$6:$AB$105,${mBase}),80)`;

    // Depreciation rate — new vs used
    const deprRate = `IF(${cond}="New",IF(${yrs}=1,$B$11,$B$12),$B$13)`;
    // Resale value
    const resale = `${otd}*(1-${deprRate})^${yrs}`;
    // Annual fuel
    const annFuel = `IF(${fuel}="Electric",(${miles}/${eff})*${ep},IF(${fuel}="Plug-In Hybrid",(${miles}/${eff})*${ep}/2+(${miles}/${mpg})*${fp}/2,(${miles}/${mpg})*${fp}))`;
    // Total interest
    const totInt = `IF(${apr}=0,0,IFERROR(-PMT(${apr}/12,${term},${amtFin})*${term}-${amtFin},0))`;

    data.push({ range: `${S}!${col}18`, values: [[`=IFERROR(${otd},0)`]] });
    data.push({ range: `${S}!${col}19`, values: [[`=IFERROR(${totInt},0)`]] });
    data.push({ range: `${S}!${col}20`, values: [[`=IFERROR(${otd}-${resale},0)`]] });
    data.push({ range: `${S}!${col}21`, values: [[`=IFERROR(${insM}*12*${yrs},0)`]] });
    data.push({ range: `${S}!${col}22`, values: [[`=IFERROR(${annFuel}*${yrs},0)`]] });
    data.push({ range: `${S}!${col}23`, values: [[`=IFERROR(${maintM}*12*${yrs},0)`]] });
    data.push({ range: `${S}!${col}24`, values: [[`=IFERROR(IF(${cond}="New",0,500)*${yrs},0)`]] }); // repairs
    data.push({ range: `${S}!${col}25`, values: [[`=IFERROR(300*${yrs},0)`]] }); // reg/taxes
    data.push({ range: `${S}!${col}26`, values: [[0]] }); // fees (editable)
    data.push({ range: `${S}!${col}27`, values: [[0]] }); // warranty (editable)
    data.push({ range: `${S}!${col}28`, values: [[0]] }); // parking/tolls (editable)
    data.push({ range: `${S}!${col}29`, values: [[`=IFERROR(${resale},0)`]] }); // resale
    data.push({ range: `${S}!${col}30`, values: [[`=IFERROR(${col}18+${col}19+${col}21+${col}22+${col}23+${col}24+${col}25+${col}26+${col}27+${col}28-${col}29,0)`]] });
    data.push({ range: `${S}!${col}31`, values: [[`=IFERROR(${col}30/${yrs},0)`]] });
    data.push({ range: `${S}!${col}32`, values: [[`=IFERROR(${col}31/12,0)`]] });
    data.push({ range: `${S}!${col}33`, values: [[`=IFERROR(${col}30/(${miles}*${yrs}),0)`]] });
  });

  // Result cards (rows 36-42)
  data.push({ range: `${S}!A36`, values: [['RESULT CARDS']] });
  const resultCards = [
    ['Lowest 1-Yr Cost',   `=IFERROR(INDEX(D15:K15,MATCH(MIN(IF($B$6=1,D30:K30,IF($B$6=3,D30:K30,IF($B$6=5,D30:K30,D30:K30)))),D30:K30,0)),"")`,  `=IFERROR(MIN(D30:K30),0)`],
    ['Lowest 5-Yr Cost',   `=IFERROR(INDEX(D15:K15,MATCH(MIN(D30:K30),D30:K30,0)),"")`,   `=IFERROR(MIN(D30:K30),0)`],
    ['Best Cost per Mile', `=IFERROR(INDEX(D15:K15,MATCH(MIN(D33:K33),D33:K33,0)),"")`,   `=IFERROR(MIN(D33:K33),0)`],
    ['Highest Resale',     `=IFERROR(INDEX(D15:K15,MATCH(MAX(D29:K29),D29:K29,0)),"")`,   `=IFERROR(MAX(D29:K29),0)`],
  ];

  resultCards.forEach(([label, vehicle, value], i) => {
    data.push({ range: `${S}!A${i+37}:C${i+37}`, values: [[label, vehicle, value]] });
  });

  await valuesBatchUpdate(id, data, 'tco-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(TCO,0,500,0,12), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Title
  reqs.push({ mergeCells: { range: gridRange(TCO,0,1,0,11), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(TCO,0,1,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:15 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:TCO, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:44 }, fields:'pixelSize' }});

  [1,2].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(TCO,ri,ri+1,0,11), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: { range: gridRange(TCO,ri,ri+1,0,11), cell: { userEnteredFormat: {
      backgroundColor: ri===1 ? hex(C.secondary) : hex(C.mutedBlue),
      textFormat: { italic:true, foregroundColor: hex(C.white), fontSize:9 },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  });

  // Controls section
  reqs.push({ repeatCell: { range: gridRange(TCO,4,5,0,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(TCO,5,13,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Number format controls
  reqs.push({ repeatCell: { range: gridRange(TCO,10,13,1,2), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.0%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Vehicle ID / header rows
  reqs.push({ repeatCell: { range: gridRange(TCO,14,15,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white) }, horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
  reqs.push({ repeatCell: { range: gridRange(TCO,15,16,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue), textFormat: { bold:true, foregroundColor: hex(C.mainText) }, horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // Cost line labels (A18:A33)
  reqs.push({ repeatCell: { range: gridRange(TCO,17,33,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,wrapStrategy)' }});

  // Total row (index 29)
  reqs.push({ repeatCell: { range: gridRange(TCO,29,30,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Resale row (index 28) — highlight
  reqs.push({ repeatCell: { range: gridRange(TCO,28,29,3,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.success), textFormat: { bold:true, foregroundColor: hex('#1A3A1E') }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Formula cells (blue)
  [17,18,19,20,21,22,23,24,29,30,31,32].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(TCO,ri,ri+1,3,11), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});
  });

  // Input cells (26,27 = fees, warranty) — pale yellow
  [25,26,27].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(TCO,ri,ri+1,3,11), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  });

  // Currency format
  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00' };
  reqs.push({ repeatCell: { range: gridRange(TCO,17,33,3,11), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});

  // Cost per mile row (index 32) — override with number format
  reqs.push({ repeatCell: { range: gridRange(TCO,32,33,3,11), cell: { userEnteredFormat: { numberFormat: { type:'NUMBER', pattern:'$0.000' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Result cards
  reqs.push({ mergeCells: { range: gridRange(TCO,35,36,0,8), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(TCO,35,36,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(TCO,36,40,0,3), cell: { userEnteredFormat: {
    backgroundColor: hex(C.success), textFormat: { bold:true, foregroundColor: hex(C.mainText) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(TCO,36,40,2,3), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId:TCO, dimension:'COLUMNS', startIndex:0, endIndex:1 }, properties: { pixelSize:250 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:TCO, dimension:'COLUMNS', startIndex:1, endIndex:2 }, properties: { pixelSize:100 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:TCO, dimension:'COLUMNS', startIndex:3, endIndex:11 }, properties: { pixelSize:130 }, fields:'pixelSize' }});

  reqs.push({ updateSheetProperties: { properties: { sheetId:TCO, gridProperties: { frozenRowCount:6 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'tco-format');
  console.log('✓ Total Cost of Ownership');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
