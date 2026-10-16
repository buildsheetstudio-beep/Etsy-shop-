'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WJ = sheetMap['📅 Weekly Pregnancy Journal'];

// Tab: 12 cols, 48 rows
// Row 0: Title (B1:L1 merged — A1 blank to allow col A freeze)
// Row 1: Disclaimer
// Row 2: Column headers
// Row 3-41: Week 4-42 data

(async () => {
  const reqs = [];

  // ── Title row: leave A1 empty, merge B1:L1 ──────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(WJ, 0, 1, 1, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(WJ, 0, 1, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: WJ, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // ── Disclaimer row 1 ─────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(WJ, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(WJ, 1, 2, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream), textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 8 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment,wrapStrategy)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: WJ, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Header row 2 ─────────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(WJ, 2, 3, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: WJ, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // ── Alt row shading rows 3-41 ─────────────────────────────────────────────
  for (let r = 3; r < 42; r++) {
    const bg = r % 2 === 1 ? C.lavenderMist : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(WJ, r, r+1, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(bg) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // ── Input cell styling (cols 3-11 = D-L, rows 3-41) ─────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(WJ, 3, 42, 3, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.ivory), wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } },
    fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment)',
  }});

  // Formula cells A, B, C (auto-fill): pale lavender
  reqs.push({ repeatCell: {
    range: gridRange(WJ, 3, 42, 0, 3),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment)',
  }});

  // ── Column widths ────────────────────────────────────────────────────────────
  const colW = [65, 120, 180, 180, 140, 180, 180, 180, 90, 90, 90, 90];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: WJ, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // ── Row heights for data rows ────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WJ, dimension: 'ROWS', startIndex: 3, endIndex: 42 },
    properties: { pixelSize: 60 }, fields: 'pixelSize',
  }});

  // ── Borders ──────────────────────────────────────────────────────────────────
  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(WJ, 2, 42, 0, 12), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border } });

  await batchUpdate(id, reqs, 'wj-format');

  // ── Values ────────────────────────────────────────────────────────────────────
  const vdata = [];

  vdata.push({ range: `'📅 Weekly Pregnancy Journal'!B1`, values: [['📅 Weekly Pregnancy Journal']] });
  vdata.push({ range: `'📅 Weekly Pregnancy Journal'!A2`, values: [['⚠️ For personal tracking only. Always consult your healthcare provider for medical advice.']] });

  vdata.push({ range: `'📅 Weekly Pregnancy Journal'!A3:L3`, values: [[
    'Week #','Baby Size','Milestone',
    'How I\'m Feeling','Physical Symptoms','Emotional Notes',
    'Things I\'m Grateful For','Questions for Doctor',
    'Weight (lbs)','BP Systolic','BP Diastolic','Notes',
  ]]});

  // Pre-populate weeks 4-42 with formulas in A, B, C; sample data in D-L
  const journalRows = [];
  const sampleFeelings = [
    'Tired but excited — just found out!',
    'Still in shock — pregnancy test confirmed!',
    'A bit nauseous in the mornings',
    'Fatigue is real. Napping every afternoon.',
    'Morning sickness hitting hard this week.',
    'Feeling emotional but so grateful.',
    'Energy starting to return slightly.',
    'Heartburn has started — keeping antacids close.',
    'Starting to feel more like myself again.',
    'The bump is showing! So excited.',
    'Felt baby move for the first time!',
    'Really enjoying this second trimester glow.',
    'Getting a bit more tired again.',
    'Nursery planning has begun — so fun!',
    'Baby shower planning underway!',
    'Feeling big but beautiful.',
    'Back pain starting — doing prenatal yoga.',
    'Third trimester — so close now!',
    'Nesting instinct in full swing!',
    'Counting down the weeks.',
    'Sleep is getting harder.',
    'Feeling ready but also nervous.',
    'Braxton Hicks picking up.',
    'So eager to meet this little one.',
    'Packing the hospital bag this week.',
    'Everything is ready — just waiting!',
    'Come on, baby!',
    'Still waiting... staying patient.',
    'Monitoring closely — feeling good.',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
  ];
  const sampleQuestions = [
    'When should I announce?','What vitamins do I need?','How much weight gain is normal?',
    'Is this fatigue normal?','What can I take for nausea?','When is my first ultrasound?',
    'Can I exercise?','What foods should I avoid?','Is spotting normal?',
    'Anatomy scan timing?','Should I take childbirth classes?','What is the anatomy scan?',
    'Birth plan options?','Pain relief options?','Kick count guidance?',
    'When to call the hospital?','Signs of preterm labour?','GBS test timing?',
    'Hospital bag checklist?','What is a birth plan?','Epidural vs natural?',
    'Signs labour is near?','What is a doula?','When to go to hospital?',
    'What happens during induction?','C-section recovery?','Breastfeeding support?',
    'Postpartum care plan?','','','','','','','','','','','',
  ];

  for (let w = 4; w <= 42; w++) {
    const r = w - 1; // row index 3 = week 4
    const i = w - 4; // 0-based for sample data arrays
    journalRows.push([
      w,
      `=IFERROR(VLOOKUP(A${r+1},'📋 Reference Data'!$A$2:$C$40,2,FALSE),"")`,
      `=IFERROR(VLOOKUP(A${r+1},'📋 Reference Data'!$A$2:$C$40,3,FALSE),"")`,
      sampleFeelings[i] || '',
      '',
      '',
      '',
      sampleQuestions[i] || '',
      '',
      '',
      '',
      '',
    ]);
  }
  vdata.push({ range: `'📅 Weekly Pregnancy Journal'!A4:L42`, values: journalRows });

  await valuesBatchUpdate(id, vdata, 'wj-values');
  console.log('Weekly Journal complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
