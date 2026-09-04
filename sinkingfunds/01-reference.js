'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

(async () => {
  const fmt = [];
  const vals = [];

  // Helper: write a labeled list section starting at (row, col)
  const section = (row, col, label, items) => {
    vals.push({ range: `${S}!${colL(col)}${row}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, row-1, row, col, col+1), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    items.forEach((item, i) => {
      vals.push({ range: `${S}!${colL(col)}${row+1+i}`, values: [[item]] });
    });
  };

  // Column A: Owners
  section(1, 0, 'Owners', ['Self','Person 1','Person 2','Joint / Household']);

  // Column A: Fund Categories (row 7+)
  section(7, 0, 'Fund Categories', [
    'Vacation','Travel','Holidays','Christmas','Birthdays','Gifts',
    'Car Maintenance','New Vehicle','Home Maintenance','Home Renovation',
    'Furniture','Appliances','Medical','Dental','Insurance','Taxes',
    'Education','School Expenses','Wedding','Baby','Pets','Technology',
    'Annual Subscriptions','Memberships','Emergency Buffer','Moving',
    'Business','Personal','Other',
  ]);

  // Column B: Funding Methods
  section(1, 1, 'Funding Methods', [
    'Goal-Date Based','Fixed Monthly Amount','Fixed Paycheck Amount','Manual Contribution',
  ]);

  // Column B: Fund Priorities (row 7+)
  section(7, 1, 'Fund Priorities', ['Low','Medium','High','Critical']);

  // Column B: Fund Statuses (row 13+)
  section(13, 1, 'Fund Statuses', [
    'Not Started','Active','Ahead of Plan','On Track','Behind Plan',
    'Paused','Goal Reached','Closed','Archived',
  ]);

  // Column C: Transaction Types
  section(1, 2, 'Transaction Types', [
    'Contribution','Withdrawal','Transfer In','Transfer Out',
    'Adjustment Increase','Adjustment Decrease','Refund','Other',
  ]);

  // Column C: Contribution Sources (row 11+)
  section(11, 2, 'Contribution Sources', [
    'Paycheck','Monthly Budget','Bonus','Side Income',
    'Tax Refund','Gift','Cash Windfall','Transfer','Other',
  ]);

  // Column D: Frequencies
  section(1, 3, 'Frequencies', [
    'Weekly','Biweekly','Semi-Monthly','Monthly','Quarterly','Annual','Manual',
  ]);

  // Column D: Milestone Types (row 10+)
  section(10, 3, 'Milestone Types', [
    '10% Funded','25% Funded','50% Funded','75% Funded','90% Funded',
    '100% Goal Reached','Custom Amount','Custom Percentage','Target Date','Other',
  ]);

  // Column D: Milestone Statuses (row 22+)
  section(22, 3, 'Milestone Statuses', [
    'Not Started','In Progress','Achieved','Delayed','Reassess',
  ]);

  // Column E: Yes/No
  section(1, 4, 'Yes / No', ['Yes','No']);

  // Column E: Funding Order Status (row 5+)
  section(5, 4, 'Funding Order Status', [
    'Fund First','High Planning Priority','Standard Planning Priority',
    'Lower Planning Priority','Goal Reached','Paused','Insufficient Data',
  ]);

  // Column F: Multi-Pastel Card Colors
  section(1, 5, 'Card Palette', [
    'Seafoam #AFCFC7','Dusty Blue #B7C9DF','Soft Lilac #CFC2DE',
    'Muted Rose #D9B9BB','Pale Olive #C8CAA7','Soft Peach #E6C3AD',
    'Powder Aqua #BDD8D9','Misty Periwinkle #C3C8E3','Warm Blush #DEC4CC','Soft Sage #C3D4BC',
  ]);

  // Color the palette swatches
  const pastels = [C.seafoam,C.dustyBlue,C.softLilac,C.mutedRose,C.paleOlive,
                   C.softPeach,C.powderAqua,C.mistyPeri,C.warmBlush,C.softSage];
  pastels.forEach((col, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 1+i, 2+i, 5, 6), cell: { userEnteredFormat: { backgroundColor: hex(col) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Sheet title
  vals.push({ range: `${S}!A1`, values: [['REFERENCE DATA — Do Not Edit']] });

  await batchUpdate(id, fmt, 'ref-fmt');
  await valuesBatchUpdate(id, vals, 'ref-vals');

  // Hide the sheet
  await batchUpdate(id, [{
    updateSheetProperties: {
      properties: { sheetId: SID, hidden: true },
      fields: 'hidden',
    },
  }], 'ref-hide');

  console.log('✓ Reference Data complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
