'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// Col A: Content Pillar (10)
// Col B: Episode Type (10)
// Col C: Target Audience (8)
// Col D: Episode Status (6)
// Col E: Guest Outreach Status (6)
// Col F: Sponsorship Payment Status (4)
// Col G: Social Platform (6)
// Col H: Equipment Category (8)

const HEADERS = ['Content Pillar','Episode Type','Target Audience','Episode Status','Guest Outreach Status','Sponsorship Payment','Social Platform','Equipment Category'];

const DATA = {
  A: ['Personal Growth','Content Creation','Productivity','Freelancing Tips','Online Business','Mindset & Motivation','Tools & Workflows','Audience Building','Work-Life Balance','Monetization'],
  B: ['Solo Talk','Guest Interview','How-To','Behind-the-Scenes','Q&A','Storytime','Mini Episode','Lessons Learned','Launch Day','Monthly Recap'],
  C: ['New Freelancers','Solo Creators','Digital Product Sellers','Aspiring Entrepreneurs','Remote Workers','Self-employed Millennials','Content Creators','Overwhelmed Beginners'],
  D: ['Published','Scheduled','Recording','Editing','On Hold','Ready to Record'],
  E: ['Not Contacted','Contacted','Negotiating','Confirmed','Declined','Recorded'],
  F: ['Invoiced','Pending','Paid','Overdue'],
  G: ['Instagram','TikTok','YouTube Shorts','Twitter/X','LinkedIn','Newsletter'],
  H: ['Microphone','Audio Interface','Camera','Lighting','Software','Backup/Storage','Headphones','Pop Filter'],
};

(async () => {
  const reqs = [];

  // Title row
  reqs.push({ mergeCells: { range: gridRange(REF, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '📋 Reference Data (Do Not Edit)' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Header row
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.burntOrange),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Data rows bg
  for (let r = 2; r < 12; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(REF, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    }});
  }

  // Col widths
  for (let i = 0; i < 8; i++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 165 }, fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'ref-format');

  // Values
  const vals = [];
  vals.push({ range: `'📋 Reference Data'!A2:H2`, values: [HEADERS] });

  const maxRows = Math.max(...Object.values(DATA).map(a => a.length));
  for (let i = 0; i < maxRows; i++) {
    const row = ['A','B','C','D','E','F','G','H'].map(k => DATA[k][i] || '');
    vals.push({ range: `'📋 Reference Data'!A${i+3}:H${i+3}`, values: [row] });
  }

  await valuesBatchUpdate(id, vals, 'ref-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
