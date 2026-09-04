'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CK = sheetMap['✅ Executor Checklist'];

// 50 pre-filled tasks
// Columns: A=# B=Phase C=Task D=Priority E=Status F=Due Date G=Done? H=Notes
// Phase 1+2 = Completed, Phase 3 mix, Phase 4-6 = Not Started
const TASKS = [
  // Phase 1: Immediate (tasks 1-10)
  [1, 'Phase 1: Immediate', 'Obtain death certificate (minimum 5 certified copies)', 'High', 'Completed', '2025-04-14', true, 'Ordered 8 copies from registry'],
  [2, 'Phase 1: Immediate', 'Locate original Will and any codicils', 'High', 'Completed', '2025-04-14', true, 'Located in safe at family home'],
  [3, 'Phase 1: Immediate', 'Arrange funeral and notify immediate family', 'High', 'Completed', '2025-04-15', true, 'Service held 2025-04-19'],
  [4, 'Phase 1: Immediate', 'Secure deceased\'s property and valuables', 'High', 'Completed', '2025-04-14', true, 'Locks changed, contents insured'],
  [5, 'Phase 1: Immediate', 'Notify Centrelink/DHS of death', 'High', 'Completed', '2025-04-16', true, 'Online notification submitted'],
  [6, 'Phase 1: Immediate', 'Cancel Medicare and pension payments', 'High', 'Completed', '2025-04-16', true, 'Services Australia notified'],
  [7, 'Phase 1: Immediate', 'Notify immediate beneficiaries of passing', 'High', 'Completed', '2025-04-15', true, 'All 4 beneficiaries notified in person'],
  [8, 'Phase 1: Immediate', 'Check for pre-paid funeral arrangements', 'Medium', 'Completed', '2025-04-13', true, 'No pre-paid plan found'],
  [9, 'Phase 1: Immediate', 'Secure pets and dependants', 'Medium', 'Completed', '2025-04-13', true, 'N/A — no dependants'],
  [10, 'Phase 1: Immediate', 'Redirect mail temporarily', 'Low', 'Completed', '2025-04-17', true, 'Australia Post redirection set 6 months'],
  // Phase 2: First Month (tasks 11-20)
  [11, 'Phase 2: First Month', 'Engage solicitor to apply for Probate', 'High', 'Completed', '2025-04-25', true, 'Mitchell & Partners engaged'],
  [12, 'Phase 2: First Month', 'Open estate bank account', 'High', 'Completed', '2025-04-22', true, 'NAB Estate account opened'],
  [13, 'Phase 2: First Month', 'Compile full asset and liability inventory', 'High', 'Completed', '2025-04-30', true, 'See Assets & Debts tab'],
  [14, 'Phase 2: First Month', 'Obtain property valuations', 'High', 'Completed', '2025-05-02', true, 'Bay Valuation Pty Ltd engaged'],
  [15, 'Phase 2: First Month', 'Notify all financial institutions of death', 'High', 'Completed', '2025-04-24', true, 'Bank, super, insurer notified'],
  [16, 'Phase 2: First Month', 'Cancel direct debits and subscriptions', 'Medium', 'Completed', '2025-04-28', true, 'All recurring payments cancelled'],
  [17, 'Phase 2: First Month', 'Engage accountant for estate tax returns', 'High', 'Completed', '2025-05-01', true, 'Green & Associates engaged'],
  [18, 'Phase 2: First Month', 'Notify ATO of death', 'High', 'Completed', '2025-04-30', true, 'TFN cancellation in progress'],
  [19, 'Phase 2: First Month', 'Review and claim life insurance policies', 'High', 'Completed', '2025-05-05', true, 'AIA claim submitted'],
  [20, 'Phase 2: First Month', 'Gather and secure important documents', 'Medium', 'Completed', '2025-04-28', true, 'Filed in estate documents folder'],
  // Phase 3: 1-3 Months (tasks 21-30, mix)
  [21, 'Phase 3: 1-3 Months', 'Lodge probate application with Supreme Court', 'High', 'Completed', '2025-05-15', true, 'Application lodged PROB-2025-0442'],
  [22, 'Phase 3: 1-3 Months', 'Advertise for creditors (statutory notice)', 'High', 'Completed', '2025-05-20', true, 'Notice published in NSW Gazette'],
  [23, 'Phase 3: 1-3 Months', 'Prepare estate accounts (provisional)', 'High', 'In Progress', '2025-06-15', false, 'Draft with accountant'],
  [24, 'Phase 3: 1-3 Months', 'Arrange sale or transfer of real property', 'High', 'In Progress', '2025-07-01', false, 'Home listed with agent'],
  [25, 'Phase 3: 1-3 Months', 'Transfer share portfolio to estate', 'Medium', 'In Progress', '2025-06-30', false, 'ASX transfer forms submitted'],
  [26, 'Phase 3: 1-3 Months', 'Claim superannuation death benefit', 'High', 'Completed', '2025-05-28', true, 'AustralianSuper paid to estate'],
  [27, 'Phase 3: 1-3 Months', 'Pay urgent estate debts and expenses', 'High', 'Completed', '2025-06-01', true, 'Funeral and legal invoices paid'],
  [28, 'Phase 3: 1-3 Months', 'Notify all creditors of estate', 'Medium', 'In Progress', '2025-06-15', false, 'Letters sent, awaiting responses'],
  [29, 'Phase 3: 1-3 Months', 'Prepare list of estate liabilities', 'Medium', 'Completed', '2025-05-30', true, 'See Assets & Debts tab'],
  [30, 'Phase 3: 1-3 Months', 'File final personal tax return for deceased', 'High', 'In Progress', '2025-10-31', false, 'With accountant'],
  // Phase 4: 3-6 Months (tasks 31-40)
  [31, 'Phase 4: 3-6 Months', 'Obtain Grant of Probate from court', 'High', 'Not Started', '2025-08-01', false, 'Awaiting court processing'],
  [32, 'Phase 4: 3-6 Months', 'Complete transfer of real estate title', 'High', 'Not Started', '2025-09-01', false, 'Pending probate grant'],
  [33, 'Phase 4: 3-6 Months', 'Sell estate assets as required', 'High', 'Not Started', '2025-09-30', false, ''],
  [34, 'Phase 4: 3-6 Months', 'Pay all confirmed estate debts', 'High', 'Not Started', '2025-09-30', false, ''],
  [35, 'Phase 4: 3-6 Months', 'Prepare estate distribution accounts', 'High', 'Not Started', '2025-10-15', false, ''],
  [36, 'Phase 4: 3-6 Months', 'Obtain beneficiary consent/release forms', 'Medium', 'Not Started', '2025-10-30', false, ''],
  [37, 'Phase 4: 3-6 Months', 'Close digital accounts and online profiles', 'Medium', 'Not Started', '2025-09-30', false, 'See Digital Assets tab'],
  [38, 'Phase 4: 3-6 Months', 'Transfer personal property to beneficiaries', 'Medium', 'Not Started', '2025-10-01', false, ''],
  [39, 'Phase 4: 3-6 Months', 'File estate income tax return', 'High', 'Not Started', '2025-10-31', false, ''],
  [40, 'Phase 4: 3-6 Months', 'Update executor records with court', 'Low', 'Not Started', '2025-10-15', false, ''],
  // Phase 5: 6-12 Months (tasks 41-46)
  [41, 'Phase 5: 6-12 Months', 'Make interim distributions to beneficiaries', 'High', 'Not Started', '2025-11-30', false, ''],
  [42, 'Phase 5: 6-12 Months', 'Lodge estate administration accounts with court', 'High', 'Not Started', '2025-12-15', false, ''],
  [43, 'Phase 5: 6-12 Months', 'Obtain tax clearance from ATO', 'High', 'Not Started', '2026-01-31', false, ''],
  [44, 'Phase 5: 6-12 Months', 'Resolve any contested claims or disputes', 'High', 'Not Started', '2026-01-31', false, 'None anticipated'],
  [45, 'Phase 5: 6-12 Months', 'Make final residuary distribution', 'High', 'Not Started', '2026-02-28', false, ''],
  [46, 'Phase 5: 6-12 Months', 'Obtain receipts from all beneficiaries', 'Medium', 'Not Started', '2026-03-15', false, ''],
  // Phase 6: Finalisation (tasks 47-50)
  [47, 'Phase 6: Finalisation', 'File final executor accounts', 'High', 'Not Started', '2026-04-30', false, ''],
  [48, 'Phase 6: Finalisation', 'Close estate bank account', 'High', 'Not Started', '2026-05-31', false, ''],
  [49, 'Phase 6: Finalisation', 'File final tax return for estate', 'High', 'Not Started', '2026-06-30', false, ''],
  [50, 'Phase 6: Finalisation', 'Archive all estate records (retain 7 years)', 'Medium', 'Not Started', '2026-07-31', false, ''],
];

(async () => {
  const reqs = [];

  // Freeze rows 3 (title + disclaimer + header), no col freeze (full-width merges in rows 0-1)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: CK, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 56 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 48 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'ROWS', startIndex: 3, endIndex: 53 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Column widths: A=40 B=160 C=360 D=80 E=110 F=110 G=60 H=240
  const colWidths = [40, 160, 360, 80, 110, 110, 60, 240];
  colWidths.forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // ── Row 0: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(CK, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CK, 0, 1, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 1: Disclaimer ─────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(CK, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CK, 1, 2, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.disclaimerBg),
      textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 2: Column headers ─────────────────────────────────────────────────
  reqs.push({ repeatCell: { range: gridRange(CK, 2, 3, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: {
        bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
      },
    },
  }, fields: 'userEnteredFormat' }});

  // ── Data rows alternating bg ──────────────────────────────────────────────
  for (let r = 0; r < 50; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(CK, r+3, r+4, 0, 8), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // Center columns A, D, E, F, G
  [0, 3, 4, 5, 6].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(CK, 3, 53, c, c+1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  // Wrap col C (Task)
  reqs.push({ repeatCell: { range: gridRange(CK, 3, 53, 2, 3), cell: {
    userEnteredFormat: { wrapStrategy: 'WRAP' },
  }, fields: 'userEnteredFormat.wrapStrategy' }});

  // Borders on data rows
  reqs.push({ updateBorders: { range: gridRange(CK, 2, 53, 0, 8),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  // Summary stats row — row 53 (index 53)
  reqs.push({ mergeCells: { range: gridRange(CK, 53, 54, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CK, 53, 54, 0, 8), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.paleSage),
      textFormat: { foregroundColor: hex(C.mutedSage), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  await batchUpdate(id, reqs, 'checklist-fmt');

  // ── Values ────────────────────────────────────────────────────────────────
  const vals = [];

  vals.push({
    range: `'✅ Executor Checklist'!A1`,
    values: [['✅ EXECUTOR CHECKLIST — Estate of Margaret Eleanor Williams']],
  });
  vals.push({
    range: `'✅ Executor Checklist'!A2`,
    values: [['IMPORTANT DISCLAIMER: This checklist is provided as a general administrative guide only and does not constitute legal or financial advice. Each estate is unique — consult a qualified solicitor and accountant for advice specific to your circumstances.']],
  });
  vals.push({
    range: `'✅ Executor Checklist'!A3:H3`,
    values: [['#', 'Phase', 'Task', 'Priority', 'Status', 'Due Date', 'Done?', 'Notes']],
  });

  // Task data rows
  const taskRows = TASKS.map(t => [t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7]]);
  vals.push({
    range: `'✅ Executor Checklist'!A4`,
    values: taskRows,
  });

  // Summary row
  vals.push({
    range: `'✅ Executor Checklist'!A54:H54`,
    values: [['SUMMARY', '', `=COUNTIF(E4:E53,"Completed")&" / "&COUNTA(C4:C53)&" tasks completed"`, `=COUNTIF(D4:D53,"High")&" High"`, '', `=COUNTIF(G4:G53,TRUE)&" done"`, '', `=TEXT(COUNTIF(G4:G53,TRUE)/COUNTA(C4:C53),"0%")&" complete"`]],
  });

  await valuesBatchUpdate(id, vals, 'checklist-vals');
  console.log('Executor Checklist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
