'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DA = sheetMap['💻 Digital Assets & Online Accounts'];

// 15 digital accounts
// Cols: A=# B=Service/Platform C=Type D=Username/Email E=Status F=Action Required G=Assigned To H=Date Actioned I=Notes
const ACCOUNTS = [
  [1, 'Gmail (margaret.e.williams@gmail.com)', 'Email', 'margaret.e.williams@gmail.com', 'Active', 'Download data, then close', 'Robert Williams', '', 'Use Google Takeout to download data first'],
  [2, 'Outlook / Hotmail (mewilliams1950@hotmail.com)', 'Email', 'mewilliams1950@hotmail.com', 'Active', 'Download data, then close', 'Robert Williams', '', ''],
  [3, 'Facebook — Personal Profile', 'Social Media', 'margaret.e.williams@gmail.com', 'Active', 'Memorialise or close', 'Catherine Williams', '', 'Family vote: memorialise'],
  [4, 'LinkedIn — Professional Profile', 'Social Media', 'margaret.e.williams@gmail.com', 'Active', 'To Be Closed', 'Catherine Williams', '', ''],
  [5, 'Netflix', 'Subscription', 'margaret.e.williams@gmail.com', 'Closed', 'Closed', 'Robert Williams', '2025-05-01', 'Cancelled and refund obtained'],
  [6, 'Spotify', 'Subscription', 'margaret.e.williams@gmail.com', 'Closed', 'Closed', 'Robert Williams', '2025-05-01', 'Cancelled'],
  [7, 'Amazon Prime / Amazon.com.au', 'Shopping', 'margaret.e.williams@gmail.com', 'To Be Closed', 'Cancel, check for pending orders', 'Robert Williams', '', 'Gift card balance: $45.00'],
  [8, 'eBay Australia', 'Shopping', 'marg.williams.seller@ebay.com.au', 'To Be Closed', 'Close seller account', 'Robert Williams', '', 'No active listings'],
  [9, 'iCloud / Apple ID', 'Cloud Storage', 'mewilliams@icloud.com', 'Active', 'Download photos, close account', 'Thomas Williams', '', 'Photos library ~18GB — download via iCloud.com'],
  [10, 'Google Drive / Google One', 'Cloud Storage', 'margaret.e.williams@gmail.com', 'Active', 'Download and archive', 'Thomas Williams', '', '~4GB of documents'],
  [11, 'NAB Online Banking', 'Banking/Finance', 'Account via branch', 'Active', 'Notify bank of death (done)', 'Robert Williams', '2025-04-24', 'Account frozen pending estate'],
  [12, 'CommBank NetBank', 'Banking/Finance', 'Account via branch', 'Active', 'Notify bank of death (done)', 'Robert Williams', '2025-04-24', 'Transferred to estate account'],
  [13, 'MyGov / ATO Online', 'Government/Legal', 'margaret.e.williams@gmail.com', 'Active', 'Notify ATO (done)', 'Robert Williams', '2025-04-30', 'TFN cancellation initiated'],
  [14, 'AustralianSuper Member Portal', 'Investment', 'mewilliams1950@hotmail.com', 'Closed', 'Death benefit claimed (done)', 'Robert Williams', '2025-05-28', 'Benefit paid to estate'],
  [15, 'My Health Record', 'Healthcare', 'margaret.e.williams@gmail.com', 'To Be Closed', 'Submit cancellation request', 'Robert Williams', '', 'Optional — inform Medicare'],
];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DA, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  reqs.push({ updateDimensionProperties: { range: { sheetId: DA, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DA, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DA, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DA, dimension: 'ROWS', startIndex: 3, endIndex: 30 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Col widths: A=40 B=260 C=140 D=240 E=110 F=200 G=160 H=110 I=250
  const colW = [40, 260, 140, 240, 110, 200, 160, 110, 250];
  colW.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: DA, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // Title
  reqs.push({ mergeCells: { range: gridRange(DA, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DA, 0, 1, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Disclaimer bar
  reqs.push({ mergeCells: { range: gridRange(DA, 1, 2, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DA, 1, 2, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.disclaimerBg),
      textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    },
  }, fields: 'userEnteredFormat' }});

  // Headers
  reqs.push({ repeatCell: { range: gridRange(DA, 2, 3, 0, 9), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      borders: { bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 } },
    },
  }, fields: 'userEnteredFormat' }});

  // Data rows
  for (let r = 0; r < 15; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(DA, r+3, r+4, 0, 9), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Center cols A, C, E, G, H
  [0, 2, 4, 6, 7].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(DA, 3, 18, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Borders
  reqs.push({ updateBorders: { range: gridRange(DA, 2, 18, 0, 9),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  await batchUpdate(id, reqs, 'digital-fmt');

  const vals = [];
  vals.push({ range: `'💻 Digital Assets & Online Accounts'!A1`, values: [['💻 DIGITAL ASSETS & ONLINE ACCOUNTS — Estate of Margaret Eleanor Williams']] });
  vals.push({ range: `'💻 Digital Assets & Online Accounts'!A2`, values: [['IMPORTANT DISCLAIMER: Handle all digital accounts with care. Review terms of service before accessing accounts. Some actions may require death certificates. Seek legal advice before accessing accounts belonging to the deceased.']] });
  vals.push({ range: `'💻 Digital Assets & Online Accounts'!A3:I3`, values: [['#', 'Service / Platform', 'Account Type', 'Username / Email', 'Status', 'Action Required', 'Assigned To', 'Date Actioned', 'Notes']] });
  vals.push({ range: `'💻 Digital Assets & Online Accounts'!A4`, values: ACCOUNTS });

  await valuesBatchUpdate(id, vals, 'digital-vals');
  console.log('Digital Assets complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
