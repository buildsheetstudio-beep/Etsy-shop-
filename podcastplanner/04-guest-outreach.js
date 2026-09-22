'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GST = sheetMap['👥 Guest Outreach'];

// A=Guest Name | B=Contact Info | C=Topic/Expertise | D=Outreach Date | E=Follow-Up Date
// F=Status | G=Linked Episode # (formula) | H=Notes

const HEADERS = ['Guest Name','Contact Info','Topic / Expertise','Outreach Date','Follow-Up Date','Status','Linked Episode #','Notes'];
const WIDTHS  = [160, 180, 200, 100, 100, 110, 110, 260];

// 15 guests — every status present at least once
// STATUS: Not Contacted, Contacted, Negotiating, Confirmed, Declined, Recorded
// Guests that are "Confirmed" or "Recorded" must appear in Content Calendar col H for the INDEX/MATCH to resolve
const sampleData = [
  ['Sarah Chen','sarah@contentco.com','Part-time content creation & audience growth','2024-01-05','2024-01-10','Recorded','','Already recorded — Ep 2 published Jan 15'],
  ['Marcus Webb','marcus@freelanceflow.io','$10K/month freelance formula & client acquisition','2023-12-15','2023-12-22','Recorded','','Recorded before calendar start — Ep 4 published Jan 29'],
  ['Priya Nair','priya@pinterestpro.com','Pinterest strategy for email list building','2024-01-20','2024-01-27','Recorded','','Recorded for Ep 8 — published Feb 26'],
  ['Jordan Lee','hello@jordanlee.co','ADHD entrepreneurship & productivity systems','2024-02-10','2024-02-17','Recorded','','Recorded for Ep 13 — published Apr 1'],
  ['Tomás Rivera','tomas@risingco.mx','Realistic 0-to-$100K business journey','2024-02-28','2024-03-06','Recorded','','Recorded for Ep 16 — published Apr 22'],
  ['Emily Park','emily@etsyempire.com','Etsy digital products — $15K/month','2024-03-15','2024-03-22','Recorded','','Recorded for Ep 20 — published May 20'],
  ['Dr. Aisha Brown','aisha@coachbrown.com','Mindset coaching for solopreneurs','2024-03-01','2024-03-08','Recorded','','Recorded for Ep 11 — published Mar 18'],
  ['Alex Kim','alex@remoterockstar.com','Deep work & remote team productivity','2024-04-15','2024-04-22','Confirmed','','Scheduled for Ep 23 — June 10'],
  ['Dr. Maya Singh','maya@mindfulbiz.com','Imposter syndrome at high income levels','2024-04-20','2024-04-27','Confirmed','','Scheduled for Ep 26 — July 1'],
  ['Lisa Tran','lisa@affiliatequeen.com','Affiliate marketing for digital product creators','2024-05-01','2024-05-08','Confirmed','','Scheduled for Ep 30 — July 29'],
  ['Noah Carter','noah@seotools.io','SEO for podcast show notes & discoverability','2024-05-10','2024-05-17','Negotiating','','Discussing episode scope; fee TBD'],
  ['Camille Dubois','camille@creativeboss.fr','Running a creative business across time zones','2024-05-20','2024-05-27','Negotiating','','Timezone logistics to figure out; follow up again June 1'],
  ['Ryan Nguyen','ryan@videofirst.com','Short-form video repurposing from podcasts','2024-06-01','2024-06-08','Contacted','','Initial DM sent; awaiting response'],
  ['Danielle Moss','danielle@brandbuilder.co','Personal branding for introverts','2024-06-05','2024-06-12','Not Contacted','','On wish list; found via LinkedIn'],
  ['James Okafor','james@scalupco.com','Scaling from freelancer to agency','2024-05-25','','Declined','','Politely declined — not taking podcast appearances this quarter'],
];

(async () => {
  const reqs = [];

  // Row 0: Title (col A standalone for freeze compat)
  reqs.push({ mergeCells: { range: gridRange(GST, 0, 1, 1, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(GST, 0, 1, 0, 1),
      cell: {
        userEnteredValue: { stringValue: '👥' },
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
      range: gridRange(GST, 0, 1, 1, 8),
      cell: {
        userEnteredValue: { stringValue: 'Guest Outreach & Booking Tracker' },
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
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(GST, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows
  for (let r = 2; r < 17; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(GST, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GST, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(GST, 17, 500, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col G (index 6): formula bg
  reqs.push({ repeatCell: {
    range: gridRange(GST, 2, 500, 6, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GST, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Date formats
  for (const col of [3, 4]) {
    reqs.push({ repeatCell: {
      range: gridRange(GST, 2, 500, col, col+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(GST, 2, 500, 5, 6),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(GST, 2, 500, 6, 7),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'guest-format');

  // Values
  const data = [];
  data.push({ range: `'👥 Guest Outreach'!A2:H2`, values: [HEADERS] });

  const CAL_SHEET = "'🎙 Content Calendar'";
  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    const rowData = [...sampleData[i]];
    // Col G (index 6): INDEX/MATCH to find episode # where guest name matches col H of Content Calendar
    rowData[6] = `=IFERROR(INDEX(${CAL_SHEET}!$A:$A,MATCH(A${row},${CAL_SHEET}!$H:$H,0)),"Unbooked")`;
    data.push({ range: `'👥 Guest Outreach'!A${row}:H${row}`, values: [rowData] });
  }

  // Empty row formulas
  const emptyRows = [];
  for (let row = 18; row <= 200; row++) {
    const r = ['','','','','','',`=IFERROR(INDEX(${CAL_SHEET}!$A:$A,MATCH(A${row},${CAL_SHEET}!$H:$H,0)),"")`,'' ];
    emptyRows.push(r);
  }
  data.push({ range: `'👥 Guest Outreach'!A18`, values: emptyRows });

  await valuesBatchUpdate(id, data, 'guest-values');
  console.log('Guest Outreach complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
