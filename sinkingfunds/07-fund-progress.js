'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Fund Progress'];
const S = "'Fund Progress'";
const SETUP = "'Fund Setup & Goals'";
const LOG   = "'Contribution Log'";

(async () => {
  const fmt = [];
  const vals = [];

  // Multi-pastel card colors cycling
  const CARD_COLORS = [
    C.seafoam, C.dustyBlue, C.softLilac, C.mutedRose, C.paleOlive,
    C.softPeach, C.powderAqua, C.mistyPeri, C.warmBlush, C.softSage,
  ];
  const CARD_DARK = [
    '#7EADA5','#8DAABF','#ADA0C2','#BB9093','#A3A87A',
    '#C9A080','#97BFC0','#A0A7C9','#BFA0A9','#9EBA97',
  ];

  // ── Tab header ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['FUND PROGRESS OVERVIEW']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A2`, values: [['A visual snapshot of every sinking fund — balance, progress, and status at a glance']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.textMid) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // ── KPI summary row (row 4) ──────────────────────────────────────────────────
  // 4 KPI cards: Total Funds | Total Saved | Total Goal | Overall Progress
  const KPI_COLS = [0, 3, 7, 11];
  const KPI_SPANS = [3, 4, 4, 3];
  const KPI_LABELS = ['TOTAL FUNDS','TOTAL SAVED','TOTAL GOAL AMOUNT','OVERALL PROGRESS'];
  const KPI_VALS = [
    `=COUNTA('Fund Setup & Goals'!A8:A33)`,
    `=IFERROR(SUM('Fund Setup & Goals'!I8:I33),0)`,
    `=IFERROR(SUM('Fund Setup & Goals'!G8:G33),0)`,
    `=IFERROR(SUM('Fund Setup & Goals'!I8:I33)/SUM('Fund Setup & Goals'!G8:G33),0)`,
  ];
  const KPI_FMTS = ['#,##0','"$"#,##0.00','"$"#,##0.00','0.0%'];

  KPI_COLS.forEach((col, ki) => {
    const endCol = col + KPI_SPANS[ki];
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, col, endCol), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, col, endCol), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(col)}4`, values: [[KPI_LABELS[ki]]] });
    vals.push({ range: `${S}!${colL(col)}5`, values: [[KPI_VALS[ki]]] });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, col, endCol), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, col, endCol), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', numberFormat: { type: ki === 3 ? 'PERCENT' : ki === 0 ? 'NUMBER' : 'CURRENCY', pattern: KPI_FMTS[ki] } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)' } });
  });

  // ── Fund Progress Cards (rows 7 onwards, 2 cards per row, 13 rows = 26 funds) ─
  // Each card takes 7 rows × 7 cols
  // Layout: left card cols 0-6, right card cols 7-13
  // Card rows (relative): 0=header, 1=name, 2=category/owner, 3=balance/goal, 4=progress bar visual, 5=status/method, 6=gap

  const FUND_IDS = Array.from({ length: 26 }, (_, i) => `FUND-${String(i+1).padStart(3,'0')}`);

  FUND_IDS.forEach((fid, fi) => {
    const cardRow = Math.floor(fi / 2);
    const cardCol = (fi % 2) * 7;
    const baseRow = 6 + cardRow * 8; // 8 rows per card-row (7 card + 1 gap)
    const colOff  = cardCol;
    const colorIdx = fi % 10;
    const cardBg   = CARD_COLORS[colorIdx];
    const cardDark  = CARD_DARK[colorIdx];

    // Merge the card header
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow, baseRow+1, colOff, colOff+7), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(colOff)}${baseRow+1}`, values: [[`${fid}`]] });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow, baseRow+1, colOff, colOff+7), cell: { userEnteredFormat: { backgroundColor: hex(cardDark), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

    // Fund Name (row baseRow+1)
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+1, baseRow+2, colOff, colOff+7), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(colOff)}${baseRow+2}`, values: [[`=IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$B$33,2,0),"")`]] });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+1, baseRow+2, colOff, colOff+7), cell: { userEnteredFormat: { backgroundColor: hex(cardBg), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.textDark) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

    // Category | Owner (row baseRow+2)
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+2, baseRow+3, colOff, colOff+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+2, baseRow+3, colOff+3, colOff+7), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(colOff)}${baseRow+3}`, values: [[`=IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$C$33,3,0),"")`]] });
    vals.push({ range: `${S}!${colL(colOff+3)}${baseRow+3}`, values: [[`=IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$D$33,4,0),"")`]] });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+2, baseRow+3, colOff, colOff+7), cell: { userEnteredFormat: { backgroundColor: hex(cardBg), textFormat: { fontSize: 8, foregroundColor: hex(C.textMid) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

    // Balance / Goal (row baseRow+3) — left: current balance, right: goal
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+3, baseRow+4, colOff, colOff+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+3, baseRow+4, colOff+3, colOff+7), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(colOff)}${baseRow+4}`, values: [[`=IFERROR("$"&TEXT(VLOOKUP("${fid}",${SETUP}!$A$8:$I$33,9,0),"#,##0.00"),"")`]] });
    vals.push({ range: `${S}!${colL(colOff+3)}${baseRow+4}`, values: [[`=IFERROR("Goal: $"&TEXT(VLOOKUP("${fid}",${SETUP}!$A$8:$G$33,7,0),"#,##0.00"),"")`]] });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+3, baseRow+4, colOff, colOff+3), cell: { userEnteredFormat: { backgroundColor: hex(cardBg), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.textDark) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+3, baseRow+4, colOff+3, colOff+7), cell: { userEnteredFormat: { backgroundColor: hex(cardBg), textFormat: { fontSize: 9, foregroundColor: hex(C.textMid) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

    // Progress % (row baseRow+4)
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+4, baseRow+5, colOff, colOff+7), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(colOff)}${baseRow+5}`, values: [[`=IFERROR(TEXT(MIN(1,VLOOKUP("${fid}",${SETUP}!$A$8:$I$33,9,0)/VLOOKUP("${fid}",${SETUP}!$A$8:$G$33,7,0)),"0.0%")&" Funded","")`]] });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+4, baseRow+5, colOff, colOff+7), cell: { userEnteredFormat: { backgroundColor: hex(cardBg), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(cardDark) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

    // Status | Funding Method (row baseRow+5)
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+5, baseRow+6, colOff, colOff+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow+5, baseRow+6, colOff+3, colOff+7), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colL(colOff)}${baseRow+6}`, values: [[`=IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$V$33,22,0),"")`]] });
    vals.push({ range: `${S}!${colL(colOff+3)}${baseRow+6}`, values: [[`=IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$F$33,6,0),"")`]] });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+5, baseRow+6, colOff, colOff+3), cell: { userEnteredFormat: { backgroundColor: hex(cardDark), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+5, baseRow+6, colOff+3, colOff+7), cell: { userEnteredFormat: { backgroundColor: hex(cardDark), textFormat: { fontSize: 8, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

    // Gap row
    fmt.push({ repeatCell: { range: gridRange(SID, baseRow+6, baseRow+8, colOff, colOff+7), cell: { userEnteredFormat: { backgroundColor: hex(C.white) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // ── Row heights for card rows ─────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 5 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  for (let cardRow = 0; cardRow < 13; cardRow++) {
    const baseRow = 6 + cardRow * 8;
    const rowHeights = [22, 28, 20, 28, 22, 22, 8, 8];
    rowHeights.forEach((h, ri) => {
      fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: baseRow+ri, endIndex: baseRow+ri+1 }, properties: { pixelSize: h }, fields: 'pixelSize' } });
    });
  }

  // ── Column widths ─────────────────────────────────────────────────────────────
  for (let i = 0; i < 14; i++) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } });
  }

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'progress-fmt');
  await valuesBatchUpdate(id, vals, 'progress-vals');

  console.log('✓ Fund Progress complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
