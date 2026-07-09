'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SPO = sheetMap['💰 Sponsorship Tracker'];

// A=Sponsor Name | B=Episode # | C=Deal Type | D=Amount | E=Payment Status
// F=Invoice Date | G=Payment Due Date | H=Notes

const HEADERS = ['Sponsor Name','Episode #','Deal Type','Amount','Payment Status','Invoice Date','Payment Due Date','Notes'];
const WIDTHS  = [180, 90, 110, 100, 120, 110, 120, 280];

// 12 deals — all 4 payment statuses present
// Invoiced, Pending, Paid, Overdue
const sampleData = [
  ['ConvertKit','1','Mid-roll',350,'Paid','2024-01-02','2024-01-31','Recurring monthly — activated with Ep 1 launch'],
  ['Notion','2','Pre-roll',250,'Paid','2024-01-10','2024-02-10','Single episode deal — show notes link required'],
  ['Riverside.fm','5','Full Episode',800,'Paid','2024-02-01','2024-03-01','Full episode sponsorship including promo segment and CTA'],
  ['Canva','8','Post-roll',200,'Paid','2024-02-20','2024-03-20','3-month deal; Ep 8, 12, 16 — invoice 1 of 3'],
  ['ConvertKit','10','Mid-roll',350,'Paid','2024-03-01','2024-03-31','Recurring monthly — March edition'],
  ['Kajabi','13','Pre-roll',400,'Paid','2024-03-15','2024-04-15','New platform partner — usage rights for promo clip'],
  ['Canva','16','Post-roll',200,'Paid','2024-04-01','2024-04-30','Ep 8/12/16 deal — invoice 2 of 3'],
  ['Teachable','19','Mid-roll',450,'Invoiced','2024-05-01','2024-05-31','Invoice sent; awaiting PO approval from their team'],
  ['Epidemic Sound','21','Affiliate',180,'Pending','2024-05-10','2024-06-10','Affiliate commission from listener sign-ups — Q1 payout pending'],
  ['Canva','22','Post-roll',200,'Pending','2024-05-15','2024-06-15','Ep 8/12/16/22 deal — invoice 3 of 3; net-30 pending'],
  ['Descript','25','Full Episode',900,'Invoiced','2024-06-01','2024-07-01','Sponsoring Launch Day episode — premium placement'],
  ['ConvertKit','27','Mid-roll',350,'Overdue','2024-04-01','2024-05-01','June recurring — payment 15 days past due; follow up ASAP'],
];

(async () => {
  const reqs = [];

  // Row 0: Title (col A standalone for freeze compat)
  reqs.push({ mergeCells: { range: gridRange(SPO, 0, 1, 1, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SPO, 0, 1, 0, 1),
      cell: {
        userEnteredValue: { stringValue: '💰' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SPO, 0, 1, 1, 8),
      cell: {
        userEnteredValue: { stringValue: 'Sponsorship & Monetization Tracker' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SPO, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(SPO, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.amber),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SPO, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows
  for (let r = 2; r < 14; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(SPO, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SPO, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(SPO, 14, 500, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SPO, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(SPO, 2, 500, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  for (const col of [5, 6]) {
    reqs.push({ repeatCell: {
      range: gridRange(SPO, 2, 500, col, col+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(SPO, 2, 500, 1, 2),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SPO, 2, 500, 2, 3),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SPO, 2, 500, 4, 5),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'spo-format');

  // Values
  const data = [];
  data.push({ range: `'💰 Sponsorship Tracker'!A2:H2`, values: [HEADERS] });
  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    data.push({ range: `'💰 Sponsorship Tracker'!A${row}:H${row}`, values: [sampleData[i]] });
  }

  await valuesBatchUpdate(id, data, 'spo-values');
  console.log('Sponsorship Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
