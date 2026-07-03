'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const COA = sheetMap['📋 Change of Address'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'📋 Change of Address'!A1", values: [['📋 All-in-One Moving Planner — Change of Address']] });

  // ── Summary block ─────────────────────────────────────────────────────────
  values.push({ range: "'📋 Change of Address'!A2", values: [
    ['Total Entries', 'Notified', 'Pending', '% Done'],
  ]});
  values.push({ range: "'📋 Change of Address'!A3", values: [[
    '=COUNTA(A6:A65)',
    '=COUNTIF(F6:F65,"Online")+COUNTIF(F6:F65,"Phone")+COUNTIF(F6:F65,"Letter")+COUNTIF(F6:F65,"In Person")+COUNTIF(F6:F65,"Email")',
    '=A3-B3',
    '=IFERROR(B3/A3,0)',
  ]]});

  // ── Column headers (row 5) ────────────────────────────────────────────────
  values.push({ range: "'📋 Change of Address'!A5", values: [[
    '#', 'Category', 'Organisation / Contact', 'Phone / Website', 'Done?', 'Notify Method', 'Date Notified', 'Notes',
  ]]});

  // ── 45 pre-populated entries ──────────────────────────────────────────────
  const entries = [
    // Government & Official
    [1,  'Government & Official', 'Electoral Roll',                   'aec.gov.au',              false, '', '', 'Online update available'],
    [2,  'Government & Official', 'Australian Taxation Office (ATO)', 'ato.gov.au / myGov',      false, '', '', 'Update via myGov'],
    [3,  'Government & Official', 'Medicare / myGov',                 'my.gov.au',               false, '', '', ''],
    [4,  'Government & Official', 'Centrelink',                       'servicesaustralia.gov.au', false, '', '', 'If applicable'],
    [5,  'Government & Official', 'VicRoads / Transport Dept.',       'vicroads.vic.gov.au',     false, '', '', 'Driver licence + rego'],
    [6,  'Government & Official', 'Passport Office',                  'passports.gov.au',        false, '', '', 'If applicable'],
    [7,  'Government & Official', 'Australia Post — Mail Redirect',   'auspost.com.au',          false, '', '', 'Set up mail redirect'],
    // Financial
    [8,  'Financial', 'ANZ Bank',                       'anz.com.au',            false, '', '', 'All accounts'],
    [9,  'Financial', 'Commonwealth Bank (CBA)',         'commbank.com.au',       false, '', '', ''],
    [10, 'Financial', 'Westpac',                         'westpac.com.au',        false, '', '', 'If applicable'],
    [11, 'Financial', 'Superannuation Fund',             'Check fund website',    false, '', '', 'Update member details'],
    [12, 'Financial', 'PayPal',                          'paypal.com',            false, '', '', ''],
    [13, 'Financial', 'Accountant / Tax Agent',          'Direct contact',        false, '', '', ''],
    [14, 'Financial', 'Credit Card Provider',            'Card back / online',    false, '', '', ''],
    // Insurance
    [15, 'Insurance', 'Home & Contents Insurance',       'Provider website',      false, '', '', 'Update from move date'],
    [16, 'Insurance', 'Car Insurance',                   'Provider website',      false, '', '', ''],
    [17, 'Insurance', 'Health Insurance',                'Provider website',      false, '', '', ''],
    [18, 'Insurance', 'Life Insurance',                  'Provider website',      false, '', '', 'If applicable'],
    [19, 'Insurance', 'Income Protection Insurance',     'Provider website',      false, '', '', 'If applicable'],
    // Subscriptions & Delivery
    [20, 'Subscriptions & Delivery', 'Amazon',                    'amazon.com.au',         false, '', '', ''],
    [21, 'Subscriptions & Delivery', 'eBay',                      'ebay.com.au',           false, '', '', ''],
    [22, 'Subscriptions & Delivery', 'Netflix',                   'netflix.com',           false, '', '', ''],
    [23, 'Subscriptions & Delivery', 'Spotify / Apple Music',     'spotify.com / apple',   false, '', '', ''],
    [24, 'Subscriptions & Delivery', 'Magazine Subscriptions',    'Direct to publisher',   false, '', '', ''],
    [25, 'Subscriptions & Delivery', 'Meal Kit Delivery (Marley)', 'marleyspoon.com.au',   false, '', '', 'If applicable'],
    [26, 'Subscriptions & Delivery', 'Gym Membership',            'Gym reception/online',  false, '', '', ''],
    [27, 'Subscriptions & Delivery', 'Online Shopping accounts',  'Respective sites',      false, '', '', 'e.g. Catch, Kmart'],
    // Healthcare
    [28, 'Healthcare', 'General Practitioner (GP)',      'Clinic direct',         false, '', '', 'Request records transfer'],
    [29, 'Healthcare', 'Dentist',                        'Clinic direct',         false, '', '', 'Request records'],
    [30, 'Healthcare', 'Optometrist',                    'Clinic direct',         false, '', '', ''],
    [31, 'Healthcare', 'Specialist / Allied Health',     'Direct contact',        false, '', '', 'If applicable'],
    [32, 'Healthcare', 'Pharmacy',                       'Pharmacy direct',       false, '', '', 'Transfer scripts if needed'],
    [33, 'Healthcare', 'Veterinarian',                   'Vet direct',            false, '', '', 'Pet records transfer'],
    // Work & Professional
    [34, 'Work & Professional', 'Employer (HR Dept.)',             'Internal HR system',    false, '', '', ''],
    [35, 'Work & Professional', 'Professional Association',        'Member portal',         false, '', '', 'If applicable'],
    [36, 'Work & Professional', 'LinkedIn',                        'linkedin.com',          false, '', '', 'Update profile address'],
    // Personal Contacts
    [37, 'Personal Contacts', 'Family Members',                 'Phone / text',          false, '', '', ''],
    [38, 'Personal Contacts', 'Close Friends',                   'Phone / text',          false, '', '', ''],
    [39, 'Personal Contacts', 'Neighbours (old street)',         'In person',             false, '', '', ''],
    [40, 'Personal Contacts', 'School — current enrolment',     'School office',         false, '', '', 'Withdrawal notice'],
    [41, 'Personal Contacts', 'Church / Community Groups',      'Direct contact',        false, '', '', 'If applicable'],
    [42, 'Personal Contacts', 'Children\'s Extra-Curricular',   'Coach / coordinator',   false, '', '', ''],
    // Extra entries to reach 45
    [43, 'Government & Official', 'VicRoads — Boat/Trailer Rego', 'vicroads.vic.gov.au', false, '', '', 'If applicable'],
    [44, 'Financial', 'Mortgage Lender / Landlord',       'Direct contact',        false, '', '', 'New home details'],
    [45, 'Subscriptions & Delivery', 'Newspaper / Local Delivery', 'Direct contact',   false, '', '', ''],
  ];

  const entryRows = entries.map(e => [e[0], e[1], e[2], e[3], e[4], e[5], e[6], e[7]]);
  values.push({ range: "'📋 Change of Address'!A6", values: entryRows });

  await valuesBatchUpdate(id, values, 'coa-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(COA, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(COA, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.sageDark),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Summary header row 2 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(COA, 1, 2, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Summary value row 3 ───────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(COA, 2, 3, 0, 3),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(COA, 2, 3, 3, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.forestGreen), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'PERCENT', pattern: '0%' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // ── Column headers row 5 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(COA, 4, 5, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Category-grouped background colours ───────────────────────────────────
  const categoryColors = {
    'Government & Official': C.linen,
    'Financial':             C.warmCream,
    'Insurance':             C.warmCream,
    'Subscriptions & Delivery': C.linen,
    'Healthcare':            C.warmCream,
    'Work & Professional':   C.linen,
    'Personal Contacts':     C.warmCream,
  };

  // Zebra stripe rows 6–50 (index 5–49)
  for (let r = 0; r < 45; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(COA, 5 + r, 6 + r, 0, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 140, 200, 160, 60, 110, 100, 200];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: COA, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: COA, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: COA, dimension: 'ROWS', startIndex: 1, endIndex: 5 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: COA, dimension: 'ROWS', startIndex: 5, endIndex: 50 },
      properties: { pixelSize: 22 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–5 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: COA, gridProperties: { frozenRowCount: 5 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'coa-format');
  console.log('Change of Address complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
