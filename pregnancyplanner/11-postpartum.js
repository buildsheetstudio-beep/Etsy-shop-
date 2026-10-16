'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PP = sheetMap['🌙 Postpartum Recovery Tracker'];

// 11 cols, 25 rows
// Row 0: Title (B1:K1 — A1 blank for col A freeze)
// Row 1: PROMINENT DISCLAIMER (merged, eye-catching)
// Row 2: Summary labels, Row 3: Summary values
// Row 4: Col headers
// Row 5-16: 12 postpartum weeks

(async () => {
  const reqs = [];

  // Leave A1 blank, merge B1:K1 for title
  reqs.push({ mergeCells: { range: gridRange(PP, 0, 1, 1, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PP, 0, 1, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PP, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // PROMINENT DISCLAIMER row 1 — deep crimson background
  reqs.push({ mergeCells: { range: gridRange(PP, 1, 2, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PP, 1, 2, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex('#9B2335'), textFormat: { foregroundColor: hex(C.white), bold: true, italic: true, fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment,wrapStrategy)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PP, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Summary rows 2-3
  reqs.push({ repeatCell: {
    range: gridRange(PP, 2, 3, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(PP, 3, 4, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PP, dimension: 'ROWS', startIndex: 2, endIndex: 4 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(PP, 4, 5, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PP, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Alt rows 5-16
  for (let r = 5; r < 17; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(PP, r, r+1, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Wrap + tall rows for data
  reqs.push({ repeatCell: { range: gridRange(PP, 5, 17, 0, 11), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: PP, dimension: 'ROWS', startIndex: 5, endIndex: 17 }, properties: { pixelSize: 70 }, fields: 'pixelSize' } });

  // Formula col A (date) — pale lavender
  reqs.push({ repeatCell: { range: gridRange(PP, 5, 17, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment)' }});

  // Column widths
  const colW = [100, 60, 120, 120, 90, 120, 120, 180, 120, 180, 180];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: PP, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(PP, 4, 17, 0, 11), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});

  await batchUpdate(id, reqs, 'pp-format');

  const vdata = [];
  vdata.push({ range: `'🌙 Postpartum Recovery Tracker'!B1`, values: [['🌙 Postpartum Recovery Tracker']] });
  vdata.push({ range: `'🌙 Postpartum Recovery Tracker'!A2`, values: [['⚕️ IMPORTANT: This tracker is for personal journaling only. Postpartum depression, anxiety, and other conditions require professional support. If you are struggling, please contact your healthcare provider or call the Postpartum Support International helpline: 1-800-944-4773.']] });

  vdata.push({ range: `'🌙 Postpartum Recovery Tracker'!A3:K3`, values: [['Weeks Tracked','Avg Wellbeing','Avg Sleep (hrs)','Avg Energy (1-10)','Current Feeding','','','','','','']] });
  vdata.push({ range: `'🌙 Postpartum Recovery Tracker'!A4:K4`, values: [[
    '=COUNTA(B6:B17)',
    '=IFERROR(INDEX(\'📋 Reference Data\'!$P$2:$P$7,MODE(MATCH(C6:C17,\'📋 Reference Data\'!$P$2:$P$7,0))),"—")',
    '=IFERROR(ROUND(AVERAGE(D6:D17),1),"—")',
    '=IFERROR(ROUND(AVERAGE(E6:E17),1),"—")',
    '=IFERROR(INDEX(G6:G17,MATCH(MAX(A6:A17),A6:A17,0)),"—")',
    '','','','','','',
  ]]});

  vdata.push({ range: `'🌙 Postpartum Recovery Tracker'!A5:K5`, values: [[
    'Week Date','Week #','Wellbeing','Sleep (hrs/night)','Energy (1-10)','Pain Level (0-10)','Feeding Method','Physical Recovery Notes','Emotional Notes','Wins This Week','Questions for Provider',
  ]]});

  // 12 postpartum weeks — dates start at due date 2025-12-15
  // Week 1 = Dec 15, Week 2 = Dec 22, etc.
  const weeks = [];
  for (let w = 1; w <= 12; w++) {
    const rowNum = w + 4; // row 5 = week 1
    weeks.push([
      `=IFERROR('🌸 Dashboard'!$B$3+(${w-1}*7),"")`,
      w,
      w <= 4 ? (w === 1 ? 'Struggling' : w === 2 ? 'Fair' : w === 3 ? 'Fair' : 'Good') : w <= 8 ? 'Good' : 'Great',
      w === 1 ? 2 : w === 2 ? 3 : w <= 4 ? 4 : w <= 6 ? 5 : 6,
      w === 1 ? 3 : w <= 3 ? 4 : w <= 6 ? 6 : w <= 9 ? 7 : 8,
      w === 1 ? 7 : w === 2 ? 6 : w <= 4 ? 5 : w <= 6 ? 3 : w <= 8 ? 2 : 1,
      'Breastfeeding',
      w === 1 ? 'Stitches healing, very sore. Taking pain relief.' :
      w === 2 ? 'Lochia reducing. Sore but mobile.' :
      w === 3 ? 'Healing well. Still tender but much better.' :
      w === 4 ? '6-week check coming up. Feel almost normal.' :
      w === 6 ? '6-week check done — cleared for gentle exercise!' :
      w === 8 ? 'Trying gentle walks. Incision/tear healed.' :
      w === 10 ? 'Back to light yoga. Energy improving.' :
      '',
      w === 1 ? 'Baby blues hitting hard. Crying a lot. Normal per midwife.' :
      w === 2 ? 'Better days than bad now. Partner support crucial.' :
      w === 4 ? 'Starting to find a rhythm. Still emotional but positive.' :
      w === 6 ? 'Feeling much more like myself. Bonding is beautiful.' :
      w === 10 ? 'Genuinely happy. Hard but amazing.' :
      '',
      w === 1 ? 'Baby latched! We did it!' :
      w === 2 ? 'Survived first solo night.' :
      w === 4 ? 'First smile (gas, but still...)' :
      w === 6 ? 'Left house alone for 30 minutes!' :
      w === 8 ? '2-month milestone — baby thriving!' :
      w === 10 ? 'First laugh! My heart!' :
      '',
      w === 2 ? 'When does lochia stop?' :
      w === 4 ? 'Clearance for exercise at 6-week check?' :
      w === 6 ? 'Birth control options while breastfeeding?' :
      '',
    ]);
  }
  vdata.push({ range: `'🌙 Postpartum Recovery Tracker'!A6:K17`, values: weeks });

  await valuesBatchUpdate(id, vdata, 'pp-values');
  console.log('Postpartum Recovery Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
