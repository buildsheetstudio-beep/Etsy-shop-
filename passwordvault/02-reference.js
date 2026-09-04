'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];

(async () => {
  // ── Values ───────────────────────────────────────────────────────────────
  // Col A: Category | B: Password Strength | C: Account Status | D: MFA Type
  // E: Subscription Cycle | F: Subscription Status | G: Sub Category
  const data = [
    { range: "'📋 Reference Data'!A1", values: [['Category']] },
    { range: "'📋 Reference Data'!A2", values: [
      ['Email & Communication'],
      ['Social Media'],
      ['Banking & Finance'],
      ['Shopping & Retail'],
      ['Streaming & Entertainment'],
      ['Work & Productivity'],
      ['Health & Wellness'],
      ['Government & Legal'],
      ['Travel & Transport'],
      ['Gaming'],
      ['Cloud Storage'],
      ['Security & Password'],
      ['Utilities & Home'],
      ['Other'],
    ]},
    { range: "'📋 Reference Data'!B1", values: [['Password Strength']] },
    { range: "'📋 Reference Data'!B2", values: [
      ['Strong'],
      ['Medium'],
      ['Weak'],
      ['Critical'],
    ]},
    { range: "'📋 Reference Data'!C1", values: [['Account Status']] },
    { range: "'📋 Reference Data'!C2", values: [
      ['Active'],
      ['Inactive'],
      ['Suspended'],
      ['Closed'],
      ['Under Review'],
    ]},
    { range: "'📋 Reference Data'!D1", values: [['MFA Type']] },
    { range: "'📋 Reference Data'!D2", values: [
      ['Authenticator App'],
      ['SMS'],
      ['Email'],
      ['Hardware Key'],
      ['Biometric'],
      ['None'],
    ]},
    { range: "'📋 Reference Data'!E1", values: [['Billing Cycle']] },
    { range: "'📋 Reference Data'!E2", values: [
      ['Monthly'],
      ['Annual'],
      ['Bi-Annual'],
      ['Quarterly'],
      ['One-Time'],
      ['Free'],
    ]},
    { range: "'📋 Reference Data'!F1", values: [['Subscription Status']] },
    { range: "'📋 Reference Data'!F2", values: [
      ['Active'],
      ['Paused'],
      ['Cancelled'],
      ['Trial'],
      ['Expired'],
    ]},
    { range: "'📋 Reference Data'!G1", values: [['Sub Category']] },
    { range: "'📋 Reference Data'!G2", values: [
      ['Streaming Video'],
      ['Streaming Music'],
      ['Cloud Storage'],
      ['Software / SaaS'],
      ['News & Media'],
      ['Gaming'],
      ['Health & Fitness'],
      ['Shopping Membership'],
      ['Financial Tool'],
      ['Communication'],
      ['Security'],
      ['Other'],
    ]},
  ];
  await valuesBatchUpdate(id, data, 'reference-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];
  // Header row bold, indigo bg, white text
  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 0, 7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepIndigo),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Data cells
  reqs.push({ repeatCell: {
    range: gridRange(REF, 1, 20, 0, 7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleIndigo),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  // Col widths
  [0,1,2,3,4,5,6].forEach(c => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 },
      properties: { pixelSize: 160 },
      fields: 'pixelSize',
    }});
  });
  // Hide the tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true },
    fields: 'hidden',
  }});

  await batchUpdate(id, reqs, 'reference-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
