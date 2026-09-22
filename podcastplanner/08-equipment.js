'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const EQP = sheetMap['🎧 Equipment Checklist'];

// A=Item | B=Category | C=Status | D=Last Maintenance/Check Date | E=Notes

const HEADERS = ['Item','Category','Status','Last Maintenance / Check Date','Notes'];
const WIDTHS  = [220, 130, 120, 160, 320];

// 12 rows — all 3 statuses (Owned, Need to Buy, Rented), all 8 categories covered
const sampleData = [
  ['Shure SM7B Dynamic Microphone','Microphone','Owned','2024-05-01','Primary recording mic — checked for hiss; all clear'],
  ['Focusrite Scarlett 2i2 Audio Interface','Audio Interface','Owned','2024-04-15','USB connection tested; drivers updated to latest version'],
  ['Sony ZV-1 Compact Camera (for video)','Camera','Owned','2024-03-20','Used for YouTube + social clips; sensor cleaned'],
  ['Elgato Key Light Air (2-pack)','Lighting','Owned','2024-02-10','Both lights working — temperature set to 4500K'],
  ['Descript (editing + transcription software)','Software','Owned','2024-06-01','Subscription renewed; project organized in Overdub folder'],
  ['Riverside.fm (remote recording)','Software','Owned','2024-06-01','Plan: Standard — up to 4 participants; auto-backup active'],
  ['Western Digital 2TB Portable SSD','Backup/Storage','Owned','2024-05-20','Monthly backups verified; 60% capacity used'],
  ['Sony WH-1000XM5 Headphones','Headphones','Owned','2024-04-01','Noise cancelling for monitoring during remote interviews'],
  ['Kaotica Eyeball Acoustic Booth (travel)','Microphone','Rented','','Rented for travel episodes — next needed: July conference'],
  ['Rode NT-USB Mini (backup mic)','Microphone','Need to Buy','','Backup for guest on-site recording — budget $100-120'],
  ['Aputure MC RGB LED Travel Light','Lighting','Need to Buy','','Travel lighting for video ep — budget approx $80'],
  ['Shure A55M Mic Adapter (pop filter mount)','Pop Filter','Need to Buy','','Upgrade existing foam windscreen — $15-20 on Amazon'],
];

(async () => {
  const reqs = [];

  // Row 0: Title (full-width merge)
  reqs.push({ mergeCells: { range: gridRange(EQP, 0, 1, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(EQP, 0, 1, 0, 5),
      cell: {
        userEnteredValue: { stringValue: '🎧 Equipment & Recording Checklist' },
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
    range: { sheetId: EQP, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(EQP, 1, 2, 0, 5),
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
    range: { sheetId: EQP, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows
  for (let r = 2; r < 14; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(EQP, r, r+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: EQP, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(EQP, 14, 200, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: EQP, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Formats
  reqs.push({ repeatCell: {
    range: gridRange(EQP, 2, 200, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(EQP, 2, 200, 1, 3),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'eqp-format');

  // Values
  const data = [];
  data.push({ range: `'🎧 Equipment Checklist'!A2:E2`, values: [HEADERS] });
  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    data.push({ range: `'🎧 Equipment Checklist'!A${row}:E${row}`, values: [sampleData[i]] });
  }

  await valuesBatchUpdate(id, data, 'eqp-values');
  console.log('Equipment Checklist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
