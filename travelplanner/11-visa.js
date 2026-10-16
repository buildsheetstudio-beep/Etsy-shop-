const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 1488564375; // Visa & Entry Requirements

// 3 visa rows (all "No Visa Required" for UK passport to EU)
const visaRows = [
  // Country, VisaType, PassportNat, Cost, ValidityDays, AppDate, ApprovalDate, ExpiryDate, Status, Notes
  ['France','No Visa Required','British','£0','90 (Schengen)','N/A','N/A','N/A','Not Required','UK passport holders: up to 90 days in any 180-day period in Schengen area.'],
  ['Spain','No Visa Required','British','£0','90 (Schengen)','N/A','N/A','N/A','Not Required','Same Schengen rules apply. Days in France + Spain + Italy count together.'],
  ['Italy','No Visa Required','British','£0','90 (Schengen)','N/A','N/A','N/A','Not Required','Check passport expiry — must be valid for entire trip. No entry stamp needed.'],
];

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 11),
      cell: {
        userEnteredValue: { stringValue: '🛂 VISA & ENTRY REQUIREMENTS' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 2: Info banner ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 11),
      cell: {
        userEnteredValue: { stringValue: '⚠  Always verify entry requirements with official government sources before travel. Requirements can change.' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.lightAmber),
          textFormat: { foregroundColor: hex(COLOR.amber), italic: true, bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 3: Column headers ──
  const headers = ['#','Country','Visa Type','Passport Nationality','Visa Cost','Validity','Application Date','Approval Date','Expiry Date','Status','Notes'];
  headers.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 2, 3, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // ── Data rows (3 sample + blanks) ──
  const totalRows = 30;
  for (let i = 0; i < totalRows; i++) {
    const row = 3 + i;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);

    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(B${row+1}="","",ROW()-2)` },
          userEnteredFormat: { backgroundColor: hex(COLOR.champagne), textFormat: { bold: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    for (let c = 1; c <= 10; c++) {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, c, c+1),
          cell: {
            userEnteredFormat: {
              backgroundColor: c === 9 ? hex(COLOR.formulaBg) : bg,
              numberFormat: (c === 6 || c === 7 || c === 8) ? { type: 'DATE', pattern: 'dd mmm yyyy' } : undefined,
              horizontalAlignment: (c === 2 || c === 4 || c === 5 || c === 9) ? 'CENTER' : 'LEFT',
              verticalAlignment: 'MIDDLE',
              wrapStrategy: c === 10 ? 'WRAP' : 'CLIP',
            },
          },
          fields: 'userEnteredFormat',
        },
      });
    }
  }

  // ── Passport checklist section (rows 34+) ──
  const checkRow = 3 + totalRows + 1;
  reqs.push({ mergeCells: { range: gridRange(SID, checkRow, checkRow+1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, checkRow, checkRow+1, 0, 11),
      cell: {
        userEnteredValue: { stringValue: '📄  TRAVEL DOCUMENT CHECKLIST' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  const docChecklist = [
    ['Passport valid for entire trip duration'],
    ['Passport has 2+ blank pages'],
    ['Travel insurance policy printed / downloaded'],
    ['E-visa/ETA applied for (if applicable)'],
    ['Hotel confirmations printed / in app'],
    ['Flight tickets / boarding passes ready'],
    ['Emergency contact list saved offline'],
    ['Copies of documents in separate bag'],
    ['EHIC / GHIC card (if applicable)'],
    ['Currency & travel cards loaded'],
  ];

  docChecklist.forEach((item, i) => {
    const row = checkRow + 1 + i;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);
    // Checkbox col A
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredFormat: { backgroundColor: bg, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredFormat',
      },
    });
    // Item text
    reqs.push({ mergeCells: { range: gridRange(SID, row, row+1, 1, 11), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 1, 11),
        cell: {
          userEnteredValue: { stringValue: item[0] },
          userEnteredFormat: {
            backgroundColor: bg,
            textFormat: { foregroundColor: hex(COLOR.nearBlack), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Column widths
  const widths = [35,130,150,140,80,100,110,110,110,120,220];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 + totalRows + 1 + docChecklist.length }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'visa-structure');

  // ── Values ──
  const data = [];
  data.push(vr(`'Visa & Entry Requirements'!B4`, visaRows.map(r => [r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],r[8],r[9]])));

  // Document checklist checkboxes (FALSE default)
  const checkboxData = docChecklist.map(() => [false]);
  data.push(vr(`'Visa & Entry Requirements'!A${checkRow+2}`, checkboxData));

  await valuesBatchUpdate(id, data, 'visa-values');

  // Apply BOOLEAN checkbox validation to checklist
  const valReqs = [{
    setDataValidation: {
      range: gridRange(SID, checkRow+1, checkRow+1+docChecklist.length, 0, 1),
      rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true },
    },
  }];
  await batchUpdate(id, valReqs, 'visa-checkboxes');

  console.log('Visa & Entry Requirements built with', visaRows.length, 'countries + doc checklist');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
