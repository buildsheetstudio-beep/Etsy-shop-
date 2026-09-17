'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = 'Reference Data';

const LISTS = {
  'Living Status':       ['Living','Deceased','Unknown'],
  'Sex / Gender Record': ['Female','Male','Nonbinary / Other','Unknown','Not Recorded'],
  'Relationship Types':  ['Biological Parent','Adoptive Parent','Step Parent','Foster Parent','Guardian','Child','Adopted Child','Step Child','Foster Child','Spouse','Partner','Former Spouse','Sibling','Half Sibling','Step Sibling','Grandparent','Grandchild','Other'],
  'Relationship Status': ['Confirmed','Probable','Possible','Disputed','Unknown'],
  'Vital Event Types':   ['Birth','Baptism / Christening','Marriage','Civil Union','Divorce','Separation','Death','Burial','Cremation','Probate','Other'],
  'Date Precision':      ['Exact','Month / Year','Year Only','Before','After','About','Estimated','Unknown'],
  'Location Event Types':['Birthplace','Childhood Residence','Residence','Census Residence','Immigration Origin','Immigration Destination','Migration','Marriage Location','Employment Location','Military Location','Death Location','Burial Location','Other'],
  'Source Types':        ['Birth Certificate','Marriage Certificate','Death Certificate','Census','Church Record','Cemetery Record','Obituary','Newspaper','Probate Record','Will','Military Record','Immigration Record','Naturalization Record','Passenger List','City Directory','Land / Property Record','Family Bible','Photograph','Letter','Diary','Interview','Oral History','Published Book','Website','Family Tree','DNA Evidence','Other'],
  'Repository Types':    ['Personal Collection','Family Collection','Local Archive','State Archive','National Archive','Library','Church','Cemetery','Courthouse','Online Database','Genealogy Website','Newspaper Archive','Historical Society','Other'],
  'Evidence Confidence': ['Confirmed','Strong','Moderate','Tentative','Conflicting','Unknown'],
  'Source Quality':      ['Original Record','Derivative Record','Authored Narrative','Oral History','User-Contributed Tree','Unknown'],
  'Citation Status':     ['Not Cited','Partial Citation','Complete Citation','Needs Review'],
  'Research Status':     ['Not Started','Researching','Lead Found','Evidence Found','Needs Corroboration','Conflicting Evidence','Resolved','Brick Wall','Paused','Archived'],
  'Research Outcome':    ['No Result','New Person Found','New Relationship Found','New Event Found','New Location Found','Source Found','Existing Fact Confirmed','Existing Fact Challenged','Follow-Up Needed','Other'],
  'Generation Labels':   ['Self / Root','Parent','Grandparent','Great-Grandparent','2x Great-Grandparent','3x Great-Grandparent','4x Great-Grandparent','Descendant','Other'],
  'Yes / No':            ['Yes','No'],
};

// Reciprocal relationship mapping
const RECIPROCALS = [
  ['Biological Parent','Child'],
  ['Adoptive Parent','Adopted Child'],
  ['Step Parent','Step Child'],
  ['Foster Parent','Foster Child'],
  ['Guardian','Other'],
  ['Child','Biological Parent'],
  ['Adopted Child','Adoptive Parent'],
  ['Step Child','Step Parent'],
  ['Foster Child','Foster Parent'],
  ['Spouse','Spouse'],
  ['Partner','Partner'],
  ['Former Spouse','Former Spouse'],
  ['Sibling','Sibling'],
  ['Half Sibling','Half Sibling'],
  ['Step Sibling','Step Sibling'],
  ['Grandparent','Grandchild'],
  ['Grandchild','Grandparent'],
  ['Other','Other'],
];

(async () => {
  const reqs = [];
  const vals = [];

  // Hide Reference Data sheet
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, hidden: true },
    fields: 'hidden',
  }});

  // Tab background
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColor: hex(C.primary) },
    fields: 'tabColor',
  }});

  // Write list headers and values
  let col = 0;
  for (const [label, items] of Object.entries(LISTS)) {
    // Header
    vals.push({ range: `'${S}'!${String.fromCharCode(65+col)}1`, values: [[label]] });
    // Items
    items.forEach((item, r) => {
      vals.push({ range: `'${S}'!${String.fromCharCode(65+col)}${r+2}`, values: [[item]] });
    });

    // Header format
    reqs.push({ repeatCell: {
      range: gridRange(SID, 0, 1, col, col+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primaryDeep),
        textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9, fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});

    // Column width
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: col, endIndex: col+1 },
      properties: { pixelSize: 180 },
      fields: 'pixelSize',
    }});

    col++;
  }

  // Reciprocal mapping — write after lists
  const rcol = col;
  vals.push({ range: `'${S}'!${String.fromCharCode(65+rcol)}1`, values: [['Relationship Type']] });
  vals.push({ range: `'${S}'!${String.fromCharCode(66+rcol)}1`, values: [['Reciprocal']] });
  RECIPROCALS.forEach(([rt, rec], r) => {
    vals.push({ range: `'${S}'!${String.fromCharCode(65+rcol)}${r+2}`, values: [[rt]] });
    vals.push({ range: `'${S}'!${String.fromCharCode(66+rcol)}${r+2}`, values: [[rec]] });
  });
  [0,1].forEach(offset => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 0, 1, rcol+offset, rcol+offset+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondaryDeep),
        textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9, fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
  });

  await batchUpdate(id, reqs, 'ref-fmt');
  await valuesBatchUpdate(id, vals, 'ref-vals');
  console.log('✓ Reference Data complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
