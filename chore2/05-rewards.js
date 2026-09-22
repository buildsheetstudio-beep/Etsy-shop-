'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RWD = sheetMap['Reward & Prize Redemption'];
const TAB = "'Chore & Allowance Tracker'";
const RWD_NAME = "'Reward & Prize Redemption'";

// Running balance formula: total weekly allowance earned - sum of all prior redemptions for same member
// Column layout:
//  A=Family Member  B=Date  C=Reward Redeemed  D=Cost  E=Remaining Balance
function balanceF(row) {
  return `=IFERROR(IFERROR(INDEX('Chore & Allowance Tracker'!$K$3:$K$12,MATCH(A${row},'Chore & Allowance Tracker'!$I$3:$I$12,0)),0)-SUMIFS($D$2:D${row},$A$2:A${row},A${row}),0)`;
}

(async () => {
  const data = [];

  // Header row
  data.push({
    range: `${RWD_NAME}!A1:E1`,
    values: [['Family Member', 'Date', 'Reward / Prize Redeemed', 'Cost', 'Remaining Balance']]
  });

  // Sample reward redemption rows: Anna×5, Peter×5, John×5
  const rows = [
    // Anna
    ['Anna', '2025-06-03', 'Extra screen time (30 min)', 2.00],
    ['Anna', '2025-06-05', 'Choose dinner', 3.50],
    ['Anna', '2025-06-07', 'Stay up 1 hour late', 5.00],
    ['Anna', '2025-06-09', 'Candy bar of choice', 1.50],
    ['Anna', '2025-06-11', 'Movie night pick', 4.00],
    // Peter
    ['Peter', '2025-06-03', 'Video game time (1 hr)', 3.00],
    ['Peter', '2025-06-06', 'Skip one chore', 5.00],
    ['Peter', '2025-06-07', 'Extra screen time (1 hr)', 3.50],
    ['Peter', '2025-06-10', 'Ice cream trip', 4.50],
    ['Peter', '2025-06-12', 'Choose takeout restaurant', 6.00],
    // John
    ['John', '2025-06-04', 'Extra TV time (45 min)', 2.50],
    ['John', '2025-06-06', 'Choose weekend activity', 5.00],
    ['John', '2025-06-08', 'Sleepover permission', 4.00],
    ['John', '2025-06-10', 'New book of choice', 3.00],
    ['John', '2025-06-12', 'Arcade trip tokens', 5.00],
  ];

  for (let i = 0; i < rows.length; i++) {
    const row = i + 2; // 1-indexed, starting at row 2
    const [member, date, reward, cost] = rows[i];
    data.push({
      range: `${RWD_NAME}!A${row}:E${row}`,
      values: [[member, date, reward, cost, balanceF(row)]]
    });
  }

  await valuesBatchUpdate(id, data, 'rewards-values');

  // Formatting
  const reqs = [];

  // Header row
  reqs.push({ repeatCell: {
    range: gridRange(RWD, 0, 1, 0, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.charcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat'
  } });

  // Data rows — alternating chrome/white, currency on D and E
  const memberColors = { Anna: C.slot[0], Peter: C.slot[1], John: C.slot[2] };
  const memberAlt    = { Anna: C.slotAlt[0], Peter: C.slotAlt[1], John: C.slotAlt[2] };

  for (let i = 0; i < rows.length; i++) {
    const r = i + 1; // 0-indexed
    const member = rows[i][0];
    const bg = i % 2 === 0 ? memberColors[member] : memberAlt[member];
    reqs.push({ repeatCell: {
      range: gridRange(RWD, r, r + 1, 0, 5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat'
    } });
    // Currency for D and E cols (3 and 4)
    for (const ci of [3, 4]) {
      reqs.push({ repeatCell: {
        range: gridRange(RWD, r, r + 1, ci, ci + 1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex(C.charcoal), fontSize: 9 },
          numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat'
      } });
    }
  }

  // Column widths: A=130, B=100, C=220, D=80, E=130
  const colWidths = [[0,130],[1,100],[2,220],[3,80],[4,130]];
  for (const [ci, w] of colWidths) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: RWD, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    } });
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RWD, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize'
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RWD, dimension: 'ROWS', startIndex: 1, endIndex: 16 },
    properties: { pixelSize: 22 }, fields: 'pixelSize'
  } });

  // Date format for B2:B16
  reqs.push({ repeatCell: {
    range: gridRange(RWD, 1, 16, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat'
  } });

  // Borders
  const thin = { style: 'SOLID', color: hex('#BBBBBB') };
  const thick = { style: 'SOLID_MEDIUM', color: hex(C.charcoal) };
  reqs.push({ updateBorders: {
    range: gridRange(RWD, 0, 16, 0, 5),
    top: thick, bottom: thick, left: thick, right: thick,
    innerHorizontal: thin, innerVertical: thin
  } });

  await batchUpdate(id, reqs, 'rewards-format');
  console.log('Rewards complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
