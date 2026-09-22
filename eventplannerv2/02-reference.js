'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 0; // Reference Data sheetId

const CATEGORIES = ['Business','Conference','Workshop','Networking','Community','Nonprofit',
  'School','Fundraiser','Social','Birthday','Wedding','Family','Holiday','Sports','Virtual','Other'];

const CAT_COLORS = {
  Business:'#C9DCEF', Conference:'#D8CDE8', Workshop:'#F2D2B6', Networking:'#C5E0DD',
  Community:'#CBDCC5', Nonprofit:'#E9DBA7', School:'#D5E4F2',  Fundraiser:'#E8C0B8',
  Social:'#E9CCD8',   Birthday:'#F3D0BB',   Wedding:'#DDD0DF',  Family:'#EEE2C8',
  Holiday:'#D0E5D8',  Sports:'#C6D9E8',     Virtual:'#DDD5ED',  Other:'#E3E1DE',
};

const LISTS = {
  'Event Category':    CATEGORIES,
  'Event Status':      ['Planning','Confirmed','In Progress','Completed','Cancelled','Postponed'],
  'Event Type':        ['Single-Day','Multi-Day','Virtual','Hybrid','Recurring'],
  'Priority':          ['Critical','High','Medium','Low'],
  'Task Status':       ['Not Started','In Progress','Completed','Blocked','Deferred'],
  'Venue Type':        ['Hotel','Conference Center','Community Hall','Outdoor','Virtual','Restaurant','Office','School','Church','Other'],
  'Vendor Category':   ['Venue','Catering','AV/Tech','Photography','Videography','Decoration','Entertainment','Security','Transportation','Marketing','Printing','Staffing','Other'],
  'Contract Status':   ['Not Started','Sent','Signed','Expired','Cancelled'],
  'Payment Status':    ['Not Paid','Deposit Paid','Partial','Paid in Full','Refunded'],
  'Attendance Status': ['Registered','Confirmed','Attended','No-Show','Cancelled','Waitlist'],
  'Ticket Type':       ['General Admission','VIP','Staff','Speaker','Sponsor','Complimentary','Online'],
  'Meal Choice':       ['Standard','Vegetarian','Vegan','Gluten-Free','Halal','Kosher','No Meal','Unknown'],
  'Equipment Status':  ['Available','In Use','Ordered','Returned','Damaged','Missing'],
  'Expense Category':  ['Venue','Catering','AV/Tech','Marketing','Printing','Staffing','Transportation','Décor','Photography','Entertainment','Permits','Insurance','Other'],
  'Revenue Type':      ['Ticket Sales','Sponsorship','Donation','Registration Fee','Merchandise','Grant','Other'],
  'Recurrence':        ['None','Daily','Weekly','Bi-Weekly','Monthly','Annual'],
};

(async () => {
  const values = [];
  const formatRequests = [];
  const namedRanges = [];

  // Column A: Category Color Map header + data
  values.push({ range: "'Reference Data'!A1", values: [['Event Category Color Map']] });
  values.push({ range: "'Reference Data'!A2", values: [['Category']] });
  values.push({ range: "'Reference Data'!B2", values: [['Hex Color']] });

  const catRows = CATEGORIES.map((cat, i) => [cat, CAT_COLORS[cat]]);
  values.push({ range: "'Reference Data'!A3", values: catRows });

  // Style category color map
  formatRequests.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex('#2F3336'),
        textFormat: { bold: true, foregroundColor: hex('#FFFFFF'), fontSize: 11 },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat',
    }
  });
  formatRequests.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex('#4A5056'),
        textFormat: { bold: true, foregroundColor: hex('#FFFFFF') },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat',
    }
  });

  // Color each category row with its pastel
  CATEGORIES.forEach((cat, i) => {
    formatRequests.push({
      repeatCell: {
        range: gridRange(SID, 2 + i, 3 + i, 0, 2),
        cell: { userEnteredFormat: { backgroundColor: hex(CAT_COLORS[cat]) } },
        fields: 'userEnteredFormat.backgroundColor',
      }
    });
  });

  // Named range for Category Color Map
  namedRanges.push({
    addNamedRange: {
      namedRange: { name: 'REF_CATEGORY_COLORS', range: gridRange(SID, 2, 2 + CATEGORIES.length, 0, 2) }
    }
  });

  // Lookup lists starting at column D (index 3)
  const listNames = Object.keys(LISTS);
  listNames.forEach((listName, ci) => {
    const col = ci + 3; // D=3, E=4, ...
    const items = LISTS[listName];
    const colLetter = String.fromCharCode(65 + col);

    values.push({ range: `'Reference Data'!${colLetter}1`, values: [[listName]] });
    values.push({ range: `'Reference Data'!${colLetter}2`, values: items.map(v => [v]) });

    formatRequests.push({
      repeatCell: {
        range: gridRange(SID, 0, 1, col, col + 1),
        cell: { userEnteredFormat: {
          backgroundColor: hex('#2F3336'),
          textFormat: { bold: true, foregroundColor: hex('#FFFFFF') },
          horizontalAlignment: 'CENTER',
        }},
        fields: 'userEnteredFormat',
      }
    });

    const rangeName = 'REF_' + listName.toUpperCase().replace(/[^A-Z0-9]/g, '_');
    namedRanges.push({
      addNamedRange: {
        namedRange: { name: rangeName, range: gridRange(SID, 1, 1 + items.length, col, col + 1) }
      }
    });
  });

  // Column widths
  formatRequests.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 160 }, fields: 'pixelSize',
  }});
  formatRequests.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 2, endIndex: 20 },
    properties: { pixelSize: 175 }, fields: 'pixelSize',
  }});

  // Freeze row 1
  formatRequests.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 1 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  // Hide Reference Data tab
  formatRequests.push({ updateSheetProperties: {
    properties: { sheetId: SID, hidden: true },
    fields: 'hidden',
  }});

  await valuesBatchUpdate(id, values, '02-reference values');
  await batchUpdate(id, [...formatRequests, ...namedRanges], '02-reference format');
  console.log('✅ Reference Data done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
