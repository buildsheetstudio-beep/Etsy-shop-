'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const AV = sheetMap['🔐 Account Vault'];

// 18 columns: A-R
// A: # | B: Service Name | C: Website / App | D: Username / Email
// E: Category | F: Password | G: MFA Enabled | H: MFA Type
// I: Recovery Email | J: Recovery Phone | K: Security Question Hint
// L: Password Strength | M: Status | N: Date Added | O: Last Updated
// P: Password Expires | Q: Notes | R: (spare / formula-locked)

const NCOLS = 18; // A-R

// Row constants
const ROW_TITLE      = 0; // row 1
const ROW_NOTICE_1   = 1; // row 2
const ROW_NOTICE_2   = 2; // row 3
const ROW_NOTICE_3   = 3; // row 4
const ROW_NOTICE_4   = 4; // row 5
const ROW_SEC_NOTICE = 5; // row 6
const ROW_SPACER     = 6; // row 7
const ROW_SUBTITLE   = 7; // row 8  ← section header
const ROW_HEADER     = 8; // row 9  ← column headers
const ROW_DATA_START = 9; // row 10 ← first data row
const ROW_DATA_END   = 109; // row 110 (100 rows)

// Freeze: 9 rows, 2 cols — but merges in rows 1-9 must start from col C (index 2)

const COL_WIDTHS = [
  30,   // A: #
  160,  // B: Service Name
  180,  // C: Website
  200,  // D: Username/Email
  150,  // E: Category
  140,  // F: Password (hidden by color)
  80,   // G: MFA?
  140,  // H: MFA Type
  180,  // I: Recovery Email
  130,  // J: Recovery Phone
  180,  // K: Security Question Hint
  110,  // L: Strength
  90,   // M: Status
  90,   // N: Date Added
  90,   // O: Last Updated
  90,   // P: Expires
  200,  // Q: Notes
  40,   // R: spare
];

// 40 sample accounts
const ACCOUNTS = [
  // [#, Service, Website, Username/Email, Category, Password, MFA, MFAType, RecoveryEmail, RecoveryPhone, SecHint, Strength, Status, DateAdded, LastUpdated, Expires, Notes]
  [1, 'Gmail (Personal)',        'mail.google.com',       'sarah.chen@gmail.com',       'Email & Communication',       'Tr0ub4dor&3!',     'Yes','Authenticator App','backup@icloud.com',    '+1 555-0101','Childhood pet name',   'Strong',  'Active',      '15 Jan 2020','01 Jun 2025','01 Jun 2026','Primary personal email'],
  [2, 'Gmail (Work)',            'mail.google.com',       'sarah.chen@company.com',      'Work & Productivity',         'C0mp@nyS3cure#8',  'Yes','Authenticator App','sarah.chen@gmail.com', '+1 555-0101','First car model',      'Strong',  'Active',      '03 Mar 2021','15 May 2025','15 May 2026','Work email via Google Workspace'],
  [3, 'iCloud / Apple ID',       'appleid.apple.com',     'sarah.chen@icloud.com',       'Cloud Storage',               'App!3S3cur1ty$$',  'Yes','Authenticator App','sarah.chen@gmail.com', '+1 555-0101','Mothers maiden name',  'Strong',  'Active',      '10 Sep 2018','20 Apr 2025','20 Apr 2026','Linked to all Apple devices'],
  [4, 'Facebook',                'facebook.com',          'sarah.chen@gmail.com',        'Social Media',                'F@c3b00k2024!',    'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','High school mascot',   'Strong',  'Active',      '22 Jun 2015','10 Mar 2025','10 Mar 2026','Personal & family account'],
  [5, 'Instagram',               'instagram.com',         'sarahchen_life',              'Social Media',                'Inst@gr@m#99!',    'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','Favourite teacher',    'Strong',  'Active',      '14 Aug 2017','05 Feb 2025','05 Feb 2026','Photography & lifestyle'],
  [6, 'LinkedIn',                'linkedin.com',          'sarah.chen@company.com',      'Work & Productivity',         'L1nk3dIn!Prof#',   'Yes','Email',           'sarah.chen@gmail.com', '+1 555-0101','First job title',      'Strong',  'Active',      '30 Jan 2019','22 Jan 2025','22 Jan 2026','Professional networking'],
  [7, 'Chase Bank',              'chase.com',             'sarahchen_banking',           'Banking & Finance',           'Ch@s3Secur3!2024', 'Yes','Authenticator App','sarah.chen@gmail.com', '+1 555-0101','Childhood street',     'Strong',  'Active',      '08 May 2016','01 Jun 2025','01 Jun 2026','Primary checking & savings'],
  [8, 'Bank of America',         'bankofamerica.com',     'sarahchen_bofа',              'Banking & Finance',           'B@nkAm3r1c@!',     'Yes','Hardware Key',    'sarah.chen@gmail.com', '+1 555-0101','Elementary school',    'Strong',  'Active',      '12 Nov 2018','15 Apr 2025','15 Apr 2026','Joint savings account'],
  [9, 'PayPal',                  'paypal.com',            'sarah.chen@gmail.com',        'Banking & Finance',           'P@yP@l$ecure9!',   'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','Childhood nickname',   'Strong',  'Active',      '20 Jul 2017','28 Feb 2025','28 Feb 2026','Online payments'],
  [10,'Coinbase',                'coinbase.com',          'sarah.chen@gmail.com',        'Banking & Finance',           'C01nb@se!Crypt0',  'Yes','Authenticator App','backup@icloud.com',    '+1 555-0101','First crypto bought',  'Strong',  'Active',      '05 Feb 2021','10 May 2025','10 May 2026','Crypto exchange account'],
  [11,'Amazon',                  'amazon.com',            'sarah.chen@gmail.com',        'Shopping & Retail',           'Am@z0nSh0p#22!',   'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','Favourite colour',     'Strong',  'Active',      '18 Dec 2014','03 Jun 2025','03 Jun 2026','Prime member'],
  [12,'eBay',                    'ebay.com',              'sarahsells2024',              'Shopping & Retail',           'eB@ySeller!99',    'No', 'None',            'sarah.chen@gmail.com', '',           'Favourite sport',      'Medium',  'Active',      '25 Mar 2016','01 Jan 2025','01 Jan 2026','Occasional selling'],
  [13,'Netflix',                 'netflix.com',           'sarah.chen@gmail.com',        'Streaming & Entertainment',   'N3tfl1x$tream!',   'No', 'None',            'sarah.chen@gmail.com', '+1 555-0101','Favourite movie',      'Medium',  'Active',      '10 Oct 2019','20 May 2025','20 May 2026','Family plan - 4 users'],
  [14,'Spotify',                 'spotify.com',           'sarah.chen@gmail.com',        'Streaming & Entertainment',   'Sp0t1fy!Mus1c#',   'No', 'None',            'sarah.chen@gmail.com', '',           'Favourite band',       'Medium',  'Active',      '01 Apr 2018','10 Apr 2025','10 Apr 2026','Premium individual'],
  [15,'Disney+',                 'disneyplus.com',        'sarah.chen@gmail.com',        'Streaming & Entertainment',   'D1sn3yPlus#22',    'No', 'None',            'sarah.chen@gmail.com', '',           'Favourite character',  'Medium',  'Active',      '22 Nov 2019','15 Mar 2025','15 Mar 2026','Bundle with Hulu'],
  [16,'YouTube Premium',         'youtube.com',           'sarah.chen@gmail.com',        'Streaming & Entertainment',   'Y0uTub3Pr3m!',     'Yes','Authenticator App','sarah.chen@gmail.com', '+1 555-0101','Childhood friend',     'Strong',  'Active',      '14 Mar 2021','01 Feb 2025','01 Feb 2026','Linked to Google account'],
  [17,'Slack',                   'slack.com',             'sarah.chen@company.com',      'Work & Productivity',         'Sl@ck!W0rk$2024',  'Yes','Authenticator App','sarah.chen@gmail.com', '+1 555-0101','Team name',            'Strong',  'Active',      '03 Mar 2021','05 Jun 2025','05 Jun 2026','Company workspace'],
  [18,'Zoom',                    'zoom.us',               'sarah.chen@company.com',      'Work & Productivity',         'Z00m#V1deo!Call',  'Yes','Email',           'sarah.chen@gmail.com', '+1 555-0101','Conference theme',     'Strong',  'Active',      '20 Mar 2020','20 Mar 2025','20 Mar 2026','Business plan'],
  [19,'Microsoft 365',           'microsoft.com',         'sarah.chen@company.com',      'Work & Productivity',         'M$365!0ff1ce#',    'Yes','Authenticator App','sarah.chen@gmail.com', '+1 555-0101','Favourite OS',         'Strong',  'Active',      '01 Jan 2021','01 Jan 2025','01 Jan 2026','Work subscription'],
  [20,'Dropbox',                 'dropbox.com',           'sarah.chen@gmail.com',        'Cloud Storage',               'Dr0pb0x!Cl0ud$',   'Yes','Email',           'sarah.chen@gmail.com', '',           'First folder name',    'Strong',  'Active',      '05 May 2017','08 May 2025','08 May 2026','2TB business plan'],
  [21,'GitHub',                  'github.com',            'sarahchen-dev',               'Work & Productivity',         'G1tHub!D3v#2024',  'Yes','Authenticator App','sarah.chen@gmail.com', '',           'First repo name',      'Strong',  'Active',      '10 Jun 2020','10 Jun 2025','10 Jun 2026','Personal & work repos'],
  [22,'Twitter / X',             'x.com',                 'sarahchen_tweets',            'Social Media',                'Tw1tt3r!X#2024',   'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','First tweet topic',    'Strong',  'Active',      '08 Aug 2016','15 Apr 2025','15 Apr 2026','Personal & professional'],
  [23,'TikTok',                  'tiktok.com',            'sarahchen_creates',           'Social Media',                'T1kT0k!Cr3ate',    'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','First video topic',    'Medium',  'Active',      '12 Jan 2022','22 Feb 2025','22 Feb 2026','Content creator account'],
  [24,'Pinterest',               'pinterest.com',         'sarah.chen@gmail.com',        'Social Media',                'p1nter3st',        'No', 'None',            'sarah.chen@gmail.com', '',           'Favourite board',      'Weak',    'Active',      '30 Jun 2018','01 Oct 2024','01 Oct 2025','Inspiration boards'],
  [25,'Reddit',                  'reddit.com',            'u/sarahchentech',             'Social Media',                'R3dd1t!Anon#',     'Yes','Authenticator App','sarah.chen@gmail.com', '',           'First subreddit',      'Strong',  'Active',      '14 Feb 2020','14 Feb 2025','14 Feb 2026','Tech & design communities'],
  [26,'Airbnb',                  'airbnb.com',            'sarah.chen@gmail.com',        'Travel & Transport',          'A1rBnB!Tr@vel#',   'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','Favourite city',       'Strong',  'Active',      '20 Jul 2019','20 Jul 2025','20 Jul 2026','Host & guest account'],
  [27,'Booking.com',             'booking.com',           'sarah.chen@gmail.com',        'Travel & Transport',          'B00k1ng#Hotels!',  'No', 'None',            'sarah.chen@gmail.com', '',           'Favourite hotel',      'Medium',  'Active',      '05 Mar 2020','05 Mar 2025','05 Mar 2026','Genius Level 2'],
  [28,'Uber',                    'uber.com',              'sarah.chen@gmail.com',        'Travel & Transport',          'Ub3r!R1de#2024',   'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','Hometown area',        'Strong',  'Active',      '10 Sep 2018','10 Sep 2025','10 Sep 2026','Ride & Eats'],
  [29,'American Airlines',       'aa.com',                'sarah.chen@gmail.com',        'Travel & Transport',          '@mAirl1nes!FF',    'No', 'None',            'sarah.chen@gmail.com', '',           'Frequent flyer #',     'Medium',  'Active',      '22 Apr 2017','22 Apr 2025','22 Apr 2026','AAdvantage Gold member'],
  [30,'Steam',                   'store.steampowered.com','sarahgamer2024',              'Gaming',                      'St3@mG@m3r#!',     'Yes','Authenticator App','sarah.chen@gmail.com', '',           'First game bought',    'Strong',  'Active',      '25 Dec 2015','01 Jan 2025','01 Jan 2026','400+ games library'],
  [31,'Epic Games',              'epicgames.com',         'sarahchen_epic',              'Gaming',                      'Ep1cG@mes!Fn',     'Yes','Email',           'sarah.chen@gmail.com', '',           'Favourite game char',  'Strong',  'Active',      '05 Apr 2020','05 Apr 2025','05 Apr 2026','Fortnite & free games'],
  [32,'Myfitnesspal',            'myfitnesspal.com',      'sarah.chen@gmail.com',        'Health & Wellness',           'F1tness!Goals#',   'No', 'None',            'sarah.chen@gmail.com', '',           'Fitness goal',         'Medium',  'Active',      '01 Jan 2023','01 Jun 2025','01 Jun 2026','Calorie & fitness tracking'],
  [33,'Headspace',               'headspace.com',         'sarah.chen@gmail.com',        'Health & Wellness',           'M1ndful!H3ad$',    'No', 'None',            'sarah.chen@gmail.com', '',           'Meditation style',     'Weak',    'Active',      '14 Feb 2022','14 Feb 2025','14 Feb 2026','Annual subscription'],
  [34,'1Password',               '1password.com',         'sarah.chen@gmail.com',        'Security & Password',         '1P@ssw0rd!M@st3r', 'Yes','Hardware Key',    'backup@icloud.com',    '+1 555-0101','Secret word',          'Critical','Active',      '01 Jun 2023','01 Jun 2025','01 Jun 2026','Master vault — DO NOT FORGET'],
  [35,'Authy',                   'authy.com',             'sarah.chen@gmail.com',        'Security & Password',         '@uth3ntic@tor!',   'Yes','Biometric',       'sarah.chen@gmail.com', '+1 555-0101','Backup code hint',     'Strong',  'Active',      '15 Jul 2023','15 Jul 2025','15 Jul 2026','2FA backup app'],
  [36,'IRS Online',              'irs.gov',               'sarah.chen@gmail.com',        'Government & Legal',          '1RS!G0v#S3cur3',   'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','Tax year reference',   'Strong',  'Active',      '10 Apr 2022','10 Apr 2025','10 Apr 2026','Tax filing & refund status'],
  [37,'Social Security',         'ssa.gov',               'sarah.chen@gmail.com',        'Government & Legal',          'SS@Gov!2024#',     'Yes','SMS',             'sarah.chen@gmail.com', '+1 555-0101','Card last 4 digits',   'Strong',  'Active',      '20 Aug 2021','20 Aug 2025','20 Aug 2026','SS benefits & records'],
  [38,'Comcast / Xfinity',       'xfinity.com',           'sarah.chen@gmail.com',        'Utilities & Home',            'C0mc@st!Home#X',   'No', 'None',            'sarah.chen@gmail.com', '+1 555-0101','Account number hint',  'Medium',  'Active',      '01 Mar 2019','01 Mar 2025','01 Mar 2026','Internet & cable'],
  [39,'State Farm',              'statefarm.com',         'sarah.chen@gmail.com',        'Utilities & Home',            'St@t3F@rm!Ins#',   'Yes','Email',           'sarah.chen@gmail.com', '+1 555-0101','Policy number hint',   'Strong',  'Active',      '15 Sep 2020','15 Sep 2025','15 Sep 2026','Home & auto insurance'],
  [40,'Old Yahoo Email',         'yahoo.com',             'sarahchen_old@yahoo.com',     'Email & Communication',       'yahoo123',         'No', 'None',            'sarah.chen@gmail.com', '',           'Old nickname',         'Weak',    'Inactive',    '01 Jan 2010','01 Jan 2020','01 Jan 2025','Old account — consider closing'],
];

(async () => {
  const reqs = [];

  // ═══════════════════════════════════════════════════════════════════════
  // COLUMN WIDTHS
  // ═══════════════════════════════════════════════════════════════════════
  COL_WIDTHS.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: AV, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w },
      fields: 'pixelSize',
    }});
  });

  // ROW HEIGHTS
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: AV, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});
  [ROW_NOTICE_1, ROW_NOTICE_2, ROW_NOTICE_3, ROW_NOTICE_4].forEach(r => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: AV, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize',
    }});
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: AV, dimension: 'ROWS', startIndex: ROW_SEC_NOTICE, endIndex: ROW_SEC_NOTICE + 1 },
    properties: { pixelSize: 20 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: AV, dimension: 'ROWS', startIndex: ROW_SPACER, endIndex: ROW_SPACER + 1 },
    properties: { pixelSize: 10 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: AV, dimension: 'ROWS', startIndex: ROW_SUBTITLE, endIndex: ROW_SUBTITLE + 1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: AV, dimension: 'ROWS', startIndex: ROW_HEADER, endIndex: ROW_HEADER + 1 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: AV, dimension: 'ROWS', startIndex: ROW_DATA_START, endIndex: ROW_DATA_END },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  }});

  // ═══════════════════════════════════════════════════════════════════════
  // TITLE ROW (row 1) — full indigo, cols A+B separate, C:R merged
  // ═══════════════════════════════════════════════════════════════════════
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_TITLE, ROW_TITLE + 1, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepIndigo) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  // Merge C1:R1 for title text (cols 2-17)
  reqs.push({ mergeCells: { range: gridRange(AV, ROW_TITLE, ROW_TITLE + 1, 2, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_TITLE, ROW_TITLE + 1, 2, NCOLS),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Montserrat' },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'LEFT',
    }},
    fields: 'userEnteredFormat(textFormat,verticalAlignment,horizontalAlignment)',
  }});
  // Lock icon accent for col A,B area
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_TITLE, ROW_TITLE + 1, 0, 2),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 14 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(textFormat,verticalAlignment,horizontalAlignment)',
  }});

  // ═══════════════════════════════════════════════════════════════════════
  // SECURITY NOTICE BLOCK (rows 2-6)
  // Warm amber notice: "⚠️ SECURITY NOTICE — keep private"
  // ═══════════════════════════════════════════════════════════════════════
  // Background for entire notice block rows 2-6, cols A-B match
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_NOTICE_1, ROW_SEC_NOTICE + 1, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.noticeBg) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  // A & B cols: same noticeBg (already done above)
  // Merge C2:R2 notice heading
  reqs.push({ mergeCells: { range: gridRange(AV, ROW_NOTICE_1, ROW_NOTICE_1 + 1, 2, NCOLS), mergeType: 'MERGE_ALL' }});
  // Merge C3:R3
  reqs.push({ mergeCells: { range: gridRange(AV, ROW_NOTICE_2, ROW_NOTICE_2 + 1, 2, NCOLS), mergeType: 'MERGE_ALL' }});
  // Merge C4:R4
  reqs.push({ mergeCells: { range: gridRange(AV, ROW_NOTICE_3, ROW_NOTICE_3 + 1, 2, NCOLS), mergeType: 'MERGE_ALL' }});
  // Merge C5:R5
  reqs.push({ mergeCells: { range: gridRange(AV, ROW_NOTICE_4, ROW_NOTICE_4 + 1, 2, NCOLS), mergeType: 'MERGE_ALL' }});
  // Merge C6:R6
  reqs.push({ mergeCells: { range: gridRange(AV, ROW_SEC_NOTICE, ROW_SEC_NOTICE + 1, 2, NCOLS), mergeType: 'MERGE_ALL' }});

  // Notice heading text format (row 2)
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_NOTICE_1, ROW_NOTICE_1 + 1, 2, NCOLS),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.noticeFg), bold: true, fontSize: 10 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(textFormat,verticalAlignment)',
  }});
  // Notice body text format (rows 3-6)
  [ROW_NOTICE_2, ROW_NOTICE_3, ROW_NOTICE_4, ROW_SEC_NOTICE].forEach(r => {
    reqs.push({ repeatCell: {
      range: gridRange(AV, r, r + 1, 2, NCOLS),
      cell: { userEnteredFormat: {
        textFormat: { foregroundColor: hex(C.noticeFg), fontSize: 9, italic: r === ROW_SEC_NOTICE },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(textFormat,verticalAlignment)',
    }});
  });

  // Notice border — bottom of row 6
  reqs.push({ updateBorders: {
    range: gridRange(AV, ROW_SEC_NOTICE, ROW_SEC_NOTICE + 1, 0, NCOLS),
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.noticeBorder), width: 2 },
  }});

  // ═══════════════════════════════════════════════════════════════════════
  // SECTION SUBTITLE (row 8)
  // ═══════════════════════════════════════════════════════════════════════
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_SUBTITLE, ROW_SUBTITLE + 1, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.midnightIndigo) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(AV, ROW_SUBTITLE, ROW_SUBTITLE + 1, 2, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_SUBTITLE, ROW_SUBTITLE + 1, 0, 2),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midnightIndigo),
      textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 11 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_SUBTITLE, ROW_SUBTITLE + 1, 2, NCOLS),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.clearBlue), fontSize: 10, italic: true },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(textFormat,verticalAlignment)',
  }});

  // ═══════════════════════════════════════════════════════════════════════
  // COLUMN HEADERS (row 9)
  // ═══════════════════════════════════════════════════════════════════════
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_HEADER, ROW_HEADER + 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.mediumIndigo),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  }});
  // Password column header — keep same style but note it's hidden
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_HEADER, ROW_HEADER + 1, 5, 6), // col F
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midnightIndigo),
      textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});

  // ═══════════════════════════════════════════════════════════════════════
  // DATA ROWS — alternating pale indigo / white
  // ═══════════════════════════════════════════════════════════════════════
  for (let r = ROW_DATA_START; r < ROW_DATA_END; r++) {
    const bg = (r - ROW_DATA_START) % 2 === 0 ? C.paleIndigo : C.silkyWhite;
    reqs.push({ repeatCell: {
      range: gridRange(AV, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // # column (A) — center, small
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_DATA_START, ROW_DATA_END, 0, 1),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.mutedIndigo), fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
  }});

  // MFA Enabled (col G) — center
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_DATA_START, ROW_DATA_END, 6, 7),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Date columns (N, O, P) — center
  [13, 14, 15].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(AV, ROW_DATA_START, ROW_DATA_END, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'TEXT' } }},
      fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
    }});
  });

  // PASSWORD COLUMN (F = col 5) — dark-on-dark hidden text
  // Text color AND background both = deepIndigo so text is invisible until selected
  for (let r = ROW_DATA_START; r < ROW_DATA_END; r++) {
    const bg = (r - ROW_DATA_START) % 2 === 0 ? C.deepIndigo : C.midnightIndigo;
    reqs.push({ repeatCell: {
      range: gridRange(AV, r, r + 1, 5, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.deepIndigo),
        textFormat: { foregroundColor: hex(C.deepIndigo), fontSize: 9, fontFamily: 'Courier New' },
        horizontalAlignment: 'LEFT',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
  }

  // Strength column (L = col 11) — center
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_DATA_START, ROW_DATA_END, 11, 12),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Status column (M = col 12) — center
  reqs.push({ repeatCell: {
    range: gridRange(AV, ROW_DATA_START, ROW_DATA_END, 12, 13),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // ═══════════════════════════════════════════════════════════════════════
  // BORDERS on data table
  // ═══════════════════════════════════════════════════════════════════════
  reqs.push({ updateBorders: {
    range: gridRange(AV, ROW_HEADER, ROW_DATA_END, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'vault-format');

  // ═══════════════════════════════════════════════════════════════════════
  // VALUES
  // ═══════════════════════════════════════════════════════════════════════
  const vals = [];

  // Title
  vals.push({ range: "'🔐 Account Vault'!C1", values: [['  🔐  All-in-One Password & Account Vault']] });

  // Notice block
  vals.push({ range: "'🔐 Account Vault'!C2", values: [['⚠️  SECURITY NOTICE — This spreadsheet contains sensitive credentials. Keep it private and never share access.']] });
  vals.push({ range: "'🔐 Account Vault'!C3", values: [['🔒  Passwords are hidden by colour-matching text. Click any password cell to reveal it in the formula bar.']] });
  vals.push({ range: "'🔐 Account Vault'!C4", values: [['📌  Enable Google Sheets 2-Step Verification and restrict sharing. Do NOT share this file publicly.']] });
  vals.push({ range: "'🔐 Account Vault'!C5", values: [['🛡️  Consider using a dedicated password manager (1Password, Bitwarden) as your primary vault.']] });
  vals.push({ range: "'🔐 Account Vault'!C6", values: [['   Last reviewed: Jun 2025  |  Owner: Sarah Chen  |  Classification: CONFIDENTIAL']] });

  // Subtitle row
  vals.push({ range: "'🔐 Account Vault'!C8", values: [['Account Directory  ›  40 accounts tracked  ›  Scroll right to see all columns  ›  Password column (F) is colour-protected']] });

  // Column headers row 9
  vals.push({ range: "'🔐 Account Vault'!A9", values: [['#','Service Name','Website / App','Username / Email','Category','🔒 Password','MFA?','MFA Type','Recovery Email','Recovery Phone','Security Question Hint','Strength','Status','Date Added','Last Updated','Pwd Expires','Notes','']] });

  // Data rows
  const dataValues = ACCOUNTS.map(row => row);
  vals.push({ range: `'🔐 Account Vault'!A10`, values: dataValues });

  await valuesBatchUpdate(id, vals, 'vault-values');

  console.log('Account Vault complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
