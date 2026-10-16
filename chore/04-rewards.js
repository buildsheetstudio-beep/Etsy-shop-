'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RWD = sheetMap['Reward & Prize Redemption'];

const HEADERS = ['Family Member','Reward / Prize','Cost','Date Redeemed','Remaining Balance','Notes'];

const ROWS = [
  ['Emma','Small Toy',5.00,'2025-06-07','LEGO mini figure'],
  ['Liam','Screen Time',3.00,'2025-06-07','1 extra hour of tablet'],
  ['Sofia','Cash Out',4.00,'2025-06-08','Piggy bank deposit'],
  ['Noah','Special Outing',8.00,'2025-06-10','Ice cream date with Dad'],
  ['Olivia','Small Toy',6.00,'2025-06-12','Sticker pack + colored pencils'],
  ['Ethan','Cash Out',8.00,'2025-06-14','Saving for video game'],
  ['Emma','Screen Time',2.00,'2025-06-15','Movie night pick'],
  ['Liam','Special Outing',5.00,'2025-06-18','Bowling with Dad'],
  ['Mia','Small Toy',3.50,'2025-06-20','Hair accessories set'],
  ['Jackson','Cash Out',2.00,'2025-06-21','Added to savings jar'],
  ['Sofia','Screen Time',3.00,'2025-06-22','Extra game time Saturday'],
  ['Ava','Small Toy',2.00,'2025-06-24','Fun bookmarks'],
  ['Lucas','Special Outing',6.00,'2025-06-26','Roller skating with friend'],
  ['Noah','Cash Out',1.00,'2025-06-28','Pocket money'],
  ['Olivia','Screen Time',2.00,'2025-06-30','Tablet time before bed'],
];

(async () => {
  if (ROWS.length !== 15) throw new Error(`Expected 15 rows, got ${ROWS.length}`);
  const data = [];
  data.push({ range: "'Reward & Prize Redemption'!A1:F1", values: [HEADERS] });

  ROWS.forEach(([member, reward, cost, date, notes], i) => {
    const row = i + 2;
    // Remaining Balance: total earned for this member minus cumulative redeemed up to this row
    const balFormula = `=IFERROR(SUMIFS('Chore & Allowance Tracker'!$E$25:$E$1000,'Chore & Allowance Tracker'!$B$25:$B$1000,A${row})-SUMIFS($C$2:C${row},$A$2:A${row},A${row}),0)`;
    data.push({
      range: `'Reward & Prize Redemption'!A${row}:F${row}`,
      values: [[member, reward, cost, date, balFormula, notes]]
    });
  });

  await valuesBatchUpdate(id, data, 'rwd-values');

  const reqs = [];
  // Header row
  reqs.push({ repeatCell: { range: gridRange(RWD,0,1,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.turquoise), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: RWD, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Data rows
  for (let r = 1; r <= 15; r++) {
    const bg = hex(r%2===1 ? C.altRow : C.white);
    reqs.push({ repeatCell: { range: gridRange(RWD,r,r+1,0,2), cell: { userEnteredFormat: { backgroundColor: bg, textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(RWD,r,r+1,2,3), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(RWD,r,r+1,3,4), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(RWD,r,r+1,4,5), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 }, numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,horizontalAlignment)' } });
    reqs.push({ repeatCell: { range: gridRange(RWD,r,r+1,5,6), cell: { userEnteredFormat: { backgroundColor: bg, textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' } });
  }
  reqs.push({ updateDimensionProperties: { range: { sheetId: RWD, dimension: 'ROWS', startIndex: 1, endIndex: 16 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  [[0,140],[1,130],[2,90],[3,110],[4,140],[5,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: RWD, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'rwd-format');
  console.log('Reward & Prize Redemption complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
