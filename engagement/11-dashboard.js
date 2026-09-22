'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🎉 Dashboard'];

// Dashboard layout (cols A:M, 13 cols, indices 0-12):
// Row 0: Title banner (merge A:M)
// Row 1: Couple names subtitle (merge A:M)
// Row 2: Section headers
// Rows 3-9: Party Details (A:D) | KPI block (E:H)
// Row 10: Section divider
// Rows 11-12: Budget section header + col labels
// Rows 13-22: Budget by category (10 cats, SUMIF formulas)
// Row 23: spacer
// Rows 24-25: RSVP section header + col labels
// Rows 26-29: RSVP breakdown (4 statuses)
// Row 30: spacer
// Row 31: chart anchor note
// (Charts placed via 13-charts.js)

const GST_TAB = `'👥 Guest List & RSVP'`;
const BUD_TAB = `'💰 Budget & Vendor Tracker'`;
const CHK_TAB = `'✅ Party Checklist'`;
const TML_TAB = `'⏰ Day-Of Timeline'`;

// Budget categories (from Reference Data col F, items 2-11 → rows 3-12 in 1-indexed)
const BUD_CATS = ['Venue','Catering','Photography','Videography','Florals','Music/DJ','Hair & Makeup','Cake/Desserts','Invitations','Other'];
// RSVP statuses (from Reference Data col A)
const RSVP_STATUSES = ['Attending','Declined','Maybe','Awaiting Reply'];

function fmtCell(bg, textColor, bold, size, align) {
  return {
    backgroundColor: hex(bg),
    textFormat: { foregroundColor: hex(textColor), bold: !!bold, fontSize: size || 10 },
    horizontalAlignment: align || 'LEFT',
    verticalAlignment: 'MIDDLE',
  };
}

(async () => {
  const reqs = [];

  // ── Row 0: Grand title ──
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 13), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 0, 1, 0, 13),
    cell: { userEnteredFormat: { ...fmtCell(C.deepPlum, C.white, true, 18, 'CENTER'), wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  }});

  // ── Row 1: Subtitle / Couple input ──
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 13), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 1, 2, 0, 13),
    cell: { userEnteredFormat: fmtCell(C.roseGold, C.white, true, 11, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 2: Section labels (Party Details | Quick Stats) ──
  reqs.push({ mergeCells: { range: gridRange(DSH, 2, 3, 0, 4), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 2, 3, 0, 4),
    cell: { userEnteredFormat: fmtCell(C.mutedRust, C.white, true, 11, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ mergeCells: { range: gridRange(DSH, 2, 3, 4, 9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 2, 3, 4, 9),
    cell: { userEnteredFormat: fmtCell(C.sageGreen, C.white, true, 11, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Rows 3-9: Left block (Party Details) ──
  for (let r = 3; r < 10; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 0, 2),
      cell: { userEnteredFormat: fmtCell(C.palePlum, C.deepPlum, true, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 2, 4),
      cell: { userEnteredFormat: fmtCell(C.ivoryCream, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  }

  // ── Rows 3-9: Right block (KPIs) ──
  for (let r = 3; r < 10; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 4, 7),
      cell: { userEnteredFormat: fmtCell(C.lightGreen, C.deepPlum, true, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 7, 9),
      cell: { userEnteredFormat: fmtCell(C.ivoryCream, C.deepPlum, true, 12, 'CENTER') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  }

  // ── Row 10: Divider ──
  reqs.push({ mergeCells: { range: gridRange(DSH, 10, 11, 0, 13), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 10, 11, 0, 13),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 11, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Rows 11-22: Budget breakdown ──
  reqs.push({ mergeCells: { range: gridRange(DSH, 11, 12, 0, 5), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 11, 12, 0, 5),
    cell: { userEnteredFormat: fmtCell(C.mutedGold, C.deepPlum, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  for (let r = 12; r < 22; r++) {
    const bg = r % 2 === 0 ? C.lightGold : C.ivoryCream;
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 0, 5),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Right-align B, C, D
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 1, 4),
      cell: { userEnteredFormat: { horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    // Currency format
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 1, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
    // Sparkline col (D, index 3)
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 4, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(bg) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // ── Row 23: Spacer ──
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 22, 23, 0, 13),
    cell: { userEnteredFormat: { backgroundColor: hex(C.white) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── Row 24: RSVP divider ──
  reqs.push({ mergeCells: { range: gridRange(DSH, 23, 24, 0, 13), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 23, 24, 0, 13),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 11, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 25: RSVP col headers ──
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 24, 25, 0, 3),
    cell: { userEnteredFormat: fmtCell(C.roseGold, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Rows 26-29: RSVP data ──
  for (let r = 25; r < 29; r++) {
    const bg = r % 2 === 0 ? C.altRow : C.ivoryCream;
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 0, 3),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'CENTER') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 50 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 10 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 10, endIndex: 30 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // Col widths: A=160 B=120 C=120 D=120 E=160 F=80 G=80 H=120 I=20 J-M=chart area 200
  const cw = [160, 120, 120, 120, 160, 80, 80, 120, 20, 80, 80, 80, 80];
  cw.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Borders for detail sections
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(DSH, 3, 10, 0, 9),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});
  reqs.push({ updateBorders: {
    range: gridRange(DSH, 11, 22, 0, 5),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});
  reqs.push({ updateBorders: {
    range: gridRange(DSH, 24, 29, 0, 3),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  await batchUpdate(id, reqs, 'dashboard-format');

  // ── Values ──
  const data = [];

  // Row 0: Title
  data.push({ range: `'🎉 Dashboard'!A1`, values: [['🎉  ULTIMATE ENGAGEMENT PARTY PLANNER']] });

  // Row 1: Dynamic subtitle
  data.push({
    range: `'🎉 Dashboard'!A2`,
    values: [[`=IFERROR(C4&" & "&C5&" — "&TEXT(C6,"MMMM D, YYYY"),"Enter your details below ↓")`]],
  });

  // Row 2: Section labels
  data.push({ range: `'🎉 Dashboard'!A3`, values: [['📋  PARTY DETAILS']] });
  data.push({ range: `'🎉 Dashboard'!E3`, values: [['✨  AT A GLANCE']] });

  // Rows 3-9: Party Details labels (cols A:B) + inputs (cols C:D)
  const detailLabels = ['Partner 1 Name','Partner 2 Name','Party Date','Party Time','Venue','Theme / Style','Host Contact'];
  detailLabels.forEach((label, i) => {
    data.push({ range: `'🎉 Dashboard'!A${i + 4}`, values: [[label]] });
    data.push({ range: `'🎉 Dashboard'!C${i + 4}`, values: [['← Enter here']] });
  });

  // Rows 3-9: KPI labels (cols E:F) + values (cols G:H)
  const kpis = [
    ['Total Guests',       `=IFERROR(COUNTA(${GST_TAB}!B4:B200),0)`],
    ['Attending',          `=IFERROR(COUNTIF(${GST_TAB}!F4:F200,"Attending"),0)`],
    ['Total Budget',       `=IFERROR("$"&TEXT(SUM(${BUD_TAB}!D4:D200),"#,##0"),"—")`],
    ['Amount Spent',       `=IFERROR("$"&TEXT(SUM(${BUD_TAB}!E4:E200),"#,##0"),"—")`],
    ['Days to Party',      `=IFERROR(C6-TODAY()&" days","Set date ←")`],
    ['Checklist Done',     `=IFERROR(COUNTIF(${CHK_TAB}!H4:H200,TRUE)&" / "&COUNTA(${CHK_TAB}!B4:B200),"—")`],
    ['Vendors Fully Paid', `=IFERROR(COUNTIF(${BUD_TAB}!G4:G200,"Fully Paid")&" / "&COUNTA(${BUD_TAB}!B4:B200),"—")`],
  ];
  kpis.forEach(([label, formula], i) => {
    data.push({ range: `'🎉 Dashboard'!E${i + 4}`, values: [[label]] });
    data.push({ range: `'🎉 Dashboard'!H${i + 4}`, values: [[formula]] });
  });

  // Row 10: Budget section header
  data.push({ range: `'🎉 Dashboard'!A11`, values: [['💰  BUDGET OVERVIEW']] });

  // Row 11: Budget col labels
  data.push({ range: `'🎉 Dashboard'!A12:E12`, values: [['Category','Budgeted','Spent','Remaining','Spend Bar']] });

  // Rows 12-21: Budget by category
  const budgetRows = BUD_CATS.map((cat, i) => {
    const row = i + 13;
    const catQ = `"${cat}"`;
    const budgeted = `=IFERROR(SUMIF(${BUD_TAB}!$C$4:$C$200,${catQ},${BUD_TAB}!$D$4:$D$200),0)`;
    const spent    = `=IFERROR(SUMIF(${BUD_TAB}!$C$4:$C$200,${catQ},${BUD_TAB}!$E$4:$E$200),0)`;
    const remain   = `=IFERROR(B${row}-C${row},0)`;
    const sparkline = `=IFERROR(IF(B${row}>0,SPARKLINE(C${row}/B${row},{"charttype","bar";"color1","#C9A667";"color2","#EAD9C4";"max",1}),""),"")`;
    return [cat, budgeted, spent, remain, sparkline];
  });
  data.push({ range: `'🎉 Dashboard'!A13:E22`, values: budgetRows });

  // Row 23: RSVP section header
  data.push({ range: `'🎉 Dashboard'!A24`, values: [['👥  RSVP SUMMARY']] });

  // Row 24: RSVP col labels
  data.push({ range: `'🎉 Dashboard'!A25:C25`, values: [['Status','Count','%']] });

  // Rows 25-28: RSVP breakdown
  const rsvpRows = RSVP_STATUSES.map((status, i) => {
    const row = i + 26;
    const cnt = `=IFERROR(COUNTIF(${GST_TAB}!$F$4:$F$200,"${status}"),0)`;
    const pct = `=IFERROR(B${row}/COUNTA(${GST_TAB}!$B$4:$B$200),0)`;
    return [status, cnt, pct];
  });
  data.push({ range: `'🎉 Dashboard'!A26:C29`, values: rsvpRows });

  await valuesBatchUpdate(id, data, 'dashboard-values');

  // Percentage format for RSVP col C
  const pctReqs = [];
  for (let r = 25; r < 29; r++) {
    pctReqs.push({ repeatCell: {
      range: gridRange(DSH, r, r + 1, 2, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  await batchUpdate(id, pctReqs, 'dashboard-pct');

  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
