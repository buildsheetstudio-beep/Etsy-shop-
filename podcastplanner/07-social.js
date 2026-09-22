'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SOC = sheetMap['📱 Social Promotion'];

// A=Episode # | B=Platform | C=Content Type | D=Scheduled Date | E=Posted (checkbox) | F=Notes

const HEADERS = ['Episode #','Platform','Content Type','Scheduled Date','Posted?','Notes'];
const WIDTHS  = [90, 130, 120, 120, 80, 320];

// 25 rows — multiple platforms per episode, all platforms covered
// Posted = true for past dates, false for future/upcoming
const sampleData = [
  [1,'Instagram','Clip','2024-01-08',true,'Launch clip: "From 9-to-5 to Freelance" — hook: the day I quit'],
  [1,'TikTok','Clip','2024-01-09',true,'Same clip repurposed vertically — 60 sec cut'],
  [1,'Newsletter','Full Post','2024-01-08',true,'Episode recap + top 3 takeaways for subscribers'],
  [2,'Instagram','Quote Graphic','2024-01-15',true,'Sarah\'s best quote: "Consistency beats perfection every time"'],
  [2,'TikTok','Clip','2024-01-16',true,'B-roll from interview — Sarah\'s content batching reveal'],
  [2,'LinkedIn','Full Post','2024-01-15',true,'Guest spotlight post — professional audience angle'],
  [3,'Instagram','Story','2024-01-22',true,'Countdown to new episode + poll: Do you time-block?'],
  [3,'YouTube Shorts','Clip','2024-01-23',true,'90-second how-to: my morning time-block routine'],
  [4,'Instagram','Quote Graphic','2024-01-29',true,'Marcus quote: "Your first client is the hardest — after that it compounds"'],
  [4,'Twitter/X','Full Post','2024-01-29',true,'Thread: 5 things I learned from Marcus about landing $5K clients'],
  [5,'Instagram','Clip','2024-02-05',true,'Episode clip: passive income myth-busting moment'],
  [5,'Newsletter','Full Post','2024-02-05',true,'The real math on passive income — plus link to Etsy shop'],
  [8,'Instagram','Story','2024-02-26',true,'Guest reveal: Priya Nair + teaser for Pinterest episode'],
  [8,'LinkedIn','Full Post','2024-02-26',true,'B2B angle: email list building lessons from Pinterest algorithm'],
  [13,'TikTok','Clip','2024-04-01',true,'Jordan Lee ADHD system reveal — viral-bait moment'],
  [13,'YouTube Shorts','Clip','2024-04-02',true,'30-sec system breakdown — Jordan\'s calendar method'],
  [20,'Instagram','Quote Graphic','2024-05-20',true,'Emily Park: "Teaching gave me skills; Etsy gave me freedom"'],
  [20,'TikTok','Clip','2024-05-21',true,'Emily showing her Etsy dashboard — $15K month reveal'],
  [22,'Instagram','Clip','2024-06-03',false,'Schedule Mon launch morning — "3-Part Framework" hook ready'],
  [22,'TikTok','Clip','2024-06-04',false,'Repurpose same clip — add trending audio'],
  [22,'Newsletter','Full Post','2024-06-03',false,'Framework breakdown email — link to ep + free template'],
  [23,'Instagram','Story','2024-06-10',false,'Guest reveal story for Alex Kim episode'],
  [23,'LinkedIn','Full Post','2024-06-10',false,'Deep work angle — strong fit for LinkedIn professional audience'],
  [25,'Instagram','Clip','2024-06-24',false,'Launch Day episode promo — toolkit reveal clip'],
  [25,'Twitter/X','Full Post','2024-06-24',false,'Launch Day thread: everything inside the toolkit'],
];

(async () => {
  const reqs = [];

  // Row 0: Title (full-width merge)
  reqs.push({ mergeCells: { range: gridRange(SOC, 0, 1, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SOC, 0, 1, 0, 6),
      cell: {
        userEnteredValue: { stringValue: '📱 Social Media Promotion Checklist' },
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
    range: { sheetId: SOC, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(SOC, 1, 2, 0, 6),
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
    range: { sheetId: SOC, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows
  for (let r = 2; r < 27; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(SOC, r, r+1, 0, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SOC, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(SOC, 27, 500, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SOC, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Formats
  reqs.push({ repeatCell: {
    range: gridRange(SOC, 2, 500, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SOC, 2, 500, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SOC, 2, 500, 1, 3),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SOC, 2, 500, 4, 5),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'soc-format');

  // Values
  const data = [];
  data.push({ range: `'📱 Social Promotion'!A2:F2`, values: [HEADERS] });
  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    data.push({ range: `'📱 Social Promotion'!A${row}:F${row}`, values: [sampleData[i]] });
  }

  await valuesBatchUpdate(id, data, 'soc-values');
  console.log('Social Promotion complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
