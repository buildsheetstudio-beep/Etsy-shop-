'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CRD = sheetMap['❤️ Cardio & Heart Rate Zone Tracker'];
const REF = sheetMap['📋 Reference Data'];

// Columns: A=Date, B=Activity, C=Distance, D=Duration (min), E=Avg HR, F=Zone (formula), G=Calories, H=Notes
const HEADERS = ['Date', 'Activity', 'Distance', 'Duration (min)', 'Avg Heart Rate (bpm)', 'Zone', 'Calories Burned', 'Notes'];

// 12 rows spanning all 5 zones (max HR assumed 185 in dashboard input)
// Zone thresholds at 185 max HR: Z1<111, Z2 111-129, Z3 130-148, Z4 149-166, Z5 167+
const ROWS = [
  ['=DATE(2025,6,3)',  'Running',  5.0, 30, 138, '', 290, '5km steady state — Zone 3'],
  ['=DATE(2025,6,7)',  'Walking',  4.5, 45, 95,  '', 180, 'Active recovery walk — Zone 1'],
  ['=DATE(2025,6,9)',  'Cycling',  '',  25, 168, '', 310, 'HIIT bike intervals — Zone 5'],
  ['=DATE(2025,6,13)', 'Cycling',  20,  55, 118, '', 420, '20km easy Z2 ride'],
  ['=DATE(2025,6,15)', 'Running',  '',  20, 152, '', 240, 'Treadmill hill intervals — Zone 4'],
  ['=DATE(2025,6,18)', 'Running',  8.0, 45, 142, '', 380, '8km tempo — Zone 3, new 8k PB'],
  ['=DATE(2025,6,20)', 'Walking',  3.0, 35, 102, '', 150, 'Evening stroll — Zone 1'],
  ['=DATE(2025,6,22)', 'Cycling',  '',  22, 175, '', 340, 'Sprint intervals — Zone 5 peaks'],
  ['=DATE(2025,6,24)', 'Rowing',   '',  30, 126, '', 280, '5000m row — Zone 2 / 3 border'],
  ['=DATE(2025,6,26)', 'Swimming', 1.5, 40, 133, '', 320, 'Laps in pool — Zone 3'],
  ['=DATE(2025,6,28)', 'Running',  3.0, 28, 158, '', 270, 'Fartlek run — Zone 4'],
  ['=DATE(2025,6,30)', 'Running',  4.0, 25, 110, '', 200, 'Recovery jog — Zone 2 border'],
];

const DSH_NAME = "'📊 Dashboard'";
const REF_NAME = "'📋 Reference Data'";

(async () => {
  // Build rows with zone formula in col F (index 5)
  const dataRows = ROWS.map((row, i) => {
    const r = i + 2;
    // Zone formula: E2 = Avg HR, Dashboard!$C$3 = Max HR
    // MATCH HR fraction against {0,0.5,0.6,0.7,0.8,0.9} returns 1-6, INDEX into Ref!E2:E6 (+1 for header)
    const zoneFormula = `=IFERROR(INDEX(${REF_NAME}!$E$2:$E$6,MATCH(E${r}/${DSH_NAME}!$C$3,{0,0.5,0.6,0.7,0.8,0.9},1)),"")`;
    return [row[0], row[1], row[2], row[3], row[4], zoneFormula, row[6], row[7]];
  });

  await valuesBatchUpdate(id, [
    { range: "'❤️ Cardio & Heart Rate Zone Tracker'!A1:H1", values: [HEADERS] },
    { range: "'❤️ Cardio & Heart Rate Zone Tracker'!A2:H13", values: dataRows },
  ], 'crd-values');

  const reqs = [];

  // Base
  reqs.push({ repeatCell: { range: gridRange(CRD,0,25,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Header
  reqs.push({ repeatCell: { range: gridRange(CRD,0,1,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.brick), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Alternating rows
  for (let i = 1; i < 13; i += 2) {
    reqs.push({ repeatCell: { range: gridRange(CRD,i,i+1,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  }

  // Date col — date format
  reqs.push({ repeatCell: { range: gridRange(CRD,1,13,0,1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Activity, Distance, Duration, Avg HR — center
  [1,2,3,4].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(CRD,1,13,ci,ci+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  });

  // Zone col F (idx 5) — formula tint, centered, bold
  reqs.push({ repeatCell: { range: gridRange(CRD,1,13,5,6), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Calories — right align
  reqs.push({ repeatCell: { range: gridRange(CRD,1,13,6,7), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Input bg
  reqs.push({ repeatCell: { range: gridRange(CRD,1,13,0,5), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(CRD,1,13,6,8), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: CRD, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CRD, dimension: 'ROWS', startIndex: 1, endIndex: 13 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,110],[1,100],[2,90],[3,120],[4,150],[5,220],[6,120],[7,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CRD, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'crd-format');
  console.log('Cardio & HR Zone Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
