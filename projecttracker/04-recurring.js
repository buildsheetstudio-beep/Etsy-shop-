'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recurring Tasks'];
const S = "'Recurring Tasks'";
const NC = 16; // A-P

// Column layout:
// A(0)=Template ID(fml) B(1)=Task Name C(2)=Assigned To D(3)=Project ID
// E(4)=Task Type F(5)=Priority G(6)=Frequency H(7)=Day of Week
// I(8)=Custom Interval J(9)=Interval Unit K(10)=Start Date
// L(11)=Next Due(fml) M(12)=Last Completed N(13)=Status O(14)=Est Hrs P(15)=Notes
// Data at row 8 (1-indexed) = row 7 (0-indexed)

const R = (name, assignee, projId, type, pri, freq, dow, interval, unit, start, lastDone, status, estH, notes) => {
  const row = Array(NC).fill('');
  row[1]  = name;
  row[2]  = assignee;
  row[3]  = projId;
  row[4]  = type;
  row[5]  = pri;
  row[6]  = freq;
  row[7]  = dow;
  row[8]  = interval;
  row[9]  = unit;
  row[10] = start;
  row[12] = lastDone;
  row[13] = status;
  row[14] = estH;
  row[15] = notes;
  return row;
};

const TEMPLATES = [
  // Daily
  R('Daily team standup','Sarah Chen','','Meeting','High','Daily','','','Days','1/5/2026','8/17/2026','Active',0.5,'15-min standup: blockers, priorities, progress'),
  R('Daily EOD task log review','Self','','Task','Medium','Daily','','','Days','1/5/2026','8/17/2026','Active',0.25,'Review and update all in-progress task statuses before end of day'),
  R('Daily email and message triage','Self','','Task','High','Weekdays','','','Days','1/5/2026','8/17/2026','Active',0.5,'Process inbox to zero; flag action items; archive read items'),

  // Weekly
  R('Weekly project status report — all active','Sarah Chen','','Deliverable','High','Weekly','Monday','','Weeks','1/5/2026','8/11/2026','Active',1.5,'1-page status update distributed to all stakeholders by Monday 9am'),
  R('Weekly team sync — full team','Olivia Park','','Meeting','High','Weekly','Wednesday','','Weeks','1/7/2026','8/13/2026','Active',1,'60-min team meeting: project updates, decisions, announcements'),
  R('Weekly content calendar update','Zara Osei','','Task','Medium','Weekly','Monday','','Weeks','7/1/2026','8/11/2026','Active',1,'Update and schedule next 7 days of social content across all channels'),
  R('Weekly expense review and approval','Tyler Banks','','Task','Medium','Weekly','Friday','','Weeks','1/2/2026','8/15/2026','Active',0.5,'Review team expense submissions; approve within policy; flag exceptions'),
  R('Weekly product roadmap update','Marcus Webb','','Task','Medium','Weekly','Thursday','','Weeks','3/5/2026','8/14/2026','Active',1,'Update priority order and status in roadmap doc; share to Slack'),
  R('Weekly bug triage meeting','Marcus Webb','','Meeting','High','Weekly','Tuesday','','Weeks','4/7/2026','8/12/2026','Active',1,'Review new bugs; assign severity; assign owners; set sprint targets'),
  R('Weekly 1:1 — Sarah and Marcus','Sarah Chen','','Meeting','Medium','Weekly','Thursday','','Weeks','2/5/2026','8/14/2026','Active',0.75,'30-min check-in: workload, blockers, career development topics'),
  R('Weekly 1:1 — Sarah and Tyler','Sarah Chen','','Meeting','Medium','Weekly','Thursday','','Weeks','2/5/2026','8/14/2026','Active',0.75,'30-min check-in; project velocity and delivery risk review'),
  R('Weekly 1:1 — Sarah and Olivia','Sarah Chen','','Meeting','Medium','Weekly','Friday','','Weeks','2/6/2026','8/15/2026','Active',0.75,'30-min check-in; team morale and event pipeline review'),
  R('Weekly 1:1 — Sarah and Zara','Sarah Chen','','Meeting','Medium','Weekly','Friday','','Weeks','2/6/2026','8/15/2026','Active',0.75,'30-min check-in; marketing performance and campaign review'),
  R('Weekly website uptime and performance check','Tyler Banks','','Task','Medium','Weekly','Monday','','Weeks','5/4/2026','8/10/2026','Active',0.5,'Check Core Web Vitals, uptime log, and error rates in GA4'),
  R('Weekly social media analytics review','Zara Osei','','Task','Medium','Weekly','Friday','','Weeks','7/3/2026','8/15/2026','Active',0.5,'Review reach, engagement, and follower growth across all platforms'),

  // Biweekly
  R('Biweekly sprint planning and retrospective','Marcus Webb','','Meeting','High','Biweekly','Monday','','Weeks','4/6/2026','8/10/2026','Active',2,'2-hour session: sprint retro (45 min) + next sprint planning (75 min)'),
  R('Biweekly all-hands company update','Olivia Park','','Meeting','Medium','Biweekly','Friday','','Weeks','2/6/2026','8/8/2026','Active',1,'30-min company-wide update: metrics, wins, announcements, Q&A'),
  R('Biweekly CRM and pipeline review','Sarah Chen','','Meeting','High','Biweekly','Wednesday','','Weeks','3/4/2026','8/5/2026','Active',1.5,'Review deal stages, forecast, and at-risk accounts'),
  R('Biweekly automated backup verification','Tyler Banks','','Task','Medium','Biweekly','Saturday','','Weeks','1/3/2026','8/9/2026','Active',0.5,'Confirm cloud and local backup integrity; test restore on 1 file'),

  // Monthly
  R('Monthly financial close and P&L review','Sarah Chen','','Task','Critical','Monthly','','','Months','1/31/2026','7/31/2026','Active',3,'Close books; prepare P&L, balance sheet, and cash flow summary'),
  R('Monthly invoice generation — all clients','Tyler Banks','','Task','Critical','Monthly','','','Months','1/31/2026','7/31/2026','Active',1.5,'Generate and send invoices by last business day of month'),
  R('Monthly payroll processing','Tyler Banks','','Task','Critical','Monthly','','','Months','1/25/2026','7/25/2026','Active',2,'Process payroll for all employees via ADP; confirm direct deposits'),
  R('Monthly client newsletter','Zara Osei','','Deliverable','Medium','Monthly','','','Months','2/28/2026','7/31/2026','Active',3,'Write, design, and send monthly client newsletter (target open rate 35%)'),
  R('Monthly analytics and KPI report','Zara Osei','','Deliverable','High','Monthly','','','Months','1/31/2026','7/31/2026','Active',2,'Compile web, social, and ad metrics into monthly performance deck'),
  R('Monthly vendor payment run','Tyler Banks','','Task','High','Monthly','','','Months','1/15/2026','7/15/2026','Active',1,'Process all vendor invoices due in the month; update AP ledger'),
  R('Monthly 30-day onboarding check-in — new clients','Tyler Banks','','Meeting','Medium','Monthly','','','Months','2/28/2026','7/31/2026','Active',2,'30-min check-in call with each client in their first 90 days'),

  // Quarterly
  R('Quarterly business review — leadership','Sarah Chen','','Meeting','Critical','Quarterly','','','Months','3/31/2026','6/30/2026','Active',4,'Full-day QBR: financial results, OKR review, next quarter planning'),
  R('Quarterly supplier performance scorecard','Tyler Banks','','Task','High','Quarterly','','','Months','3/31/2026','6/30/2026','Active',3,'Score all suppliers on delivery, quality, responsiveness, pricing'),
  R('Quarterly estimated tax payment','Self','','Task','Critical','Quarterly','','','Months','3/15/2026','6/15/2026','Active',0.5,'Submit IRS estimated quarterly tax payment (Form 1040-ES)'),
  R('Quarterly employee engagement survey','Olivia Park','','Task','Medium','Quarterly','','','Months','3/31/2026','6/30/2026','Active',3,'Deploy Lattice survey; analyze results; share action plan within 2 weeks'),

  // Annual
  R('Annual business insurance renewal review','Tyler Banks','','Task','High','Annual','','','Years','12/1/2025','12/1/2025','Active',2,'Review all policies (GL, E&O, cyber, property) and get competing quotes'),
  R('Annual software license audit and renewal','Tyler Banks','','Task','Medium','Annual','','','Years','12/15/2025','12/15/2025','Active',3,'Audit all SaaS subscriptions; remove unused; negotiate renewals'),
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['RECURRING TASKS']] });

  // ── Subtitle ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Define recurring task templates. Next Due Date is auto-calculated from Last Completed and Frequency. Reference Template ID in Master Task Log.']] });

  // ── Instructions ──────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['FORMULA COLUMNS (blue): Template ID, Next Due Date. To log a completed occurrence: update "Last Completed" date. To spawn a task: add entry to Master Task Log with this Template ID.']] });

  // ── Stats row ─────────────────────────────────────────────────────────────
  const stats = [
    ['Active Templates', `=COUNTIF(N8:N${8+TEMPLATES.length+200},"Active")`],
    ['Due This Week',    `=COUNTIFS(L8:L${8+TEMPLATES.length+200},">="&TODAY(),L8:L${8+TEMPLATES.length+200},"<="&(TODAY()+7),N8:N${8+TEMPLATES.length+200},"Active")`],
    ['Daily',            `=COUNTIF(G8:G${8+TEMPLATES.length+200},"Daily")`],
    ['Weekly',           `=COUNTIF(G8:G${8+TEMPLATES.length+200},"Weekly")`],
    ['Monthly',          `=COUNTIF(G8:G${8+TEMPLATES.length+200},"Monthly")`],
    ['Paused',           `=COUNTIF(N8:N${8+TEMPLATES.length+200},"Paused")`],
  ];
  const statW = Math.floor(NC / stats.length); // 2 cols each for 6 stats
  stats.forEach(([label, fml], i) => {
    const c = i * statW;
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+statW), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, c, c+statW), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, c, c+statW), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, c, c+statW), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}5`, values: [[fml]] });
  });

  // Spacer row 5 (0-indexed)
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column Headers (row 6, 0-indexed = row 7) ────────────────────────────
  const COLS = ['Template ID','Task Name','Assigned To','Project ID','Task Type','Priority',
    'Frequency','Day of Week','Custom Interval','Interval Unit','Start Date',
    'Next Due Date','Last Completed','Status','Est Hrs','Notes'];
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A7`, values: [COLS] });

  // ── Freeze ────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Data Validation ───────────────────────────────────────────────────────
  const dvList = (col, items) => ({
    setDataValidation: {
      range: gridRange(SID, 7, 507, col, col+1),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) },
        showCustomUi: true, strict: false,
      },
    },
  });
  fmt.push(dvList(4, ['Task','Subtask','Milestone Task','Meeting','Approval','Deliverable','Follow-Up','Recurring Task','Other']));
  fmt.push(dvList(5, ['Low','Medium','High','Urgent','Critical']));
  fmt.push(dvList(6, ['None','Daily','Weekdays','Weekly','Biweekly','Monthly','Quarterly','Semiannual','Annual','Custom Interval']));
  fmt.push(dvList(7, ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']));
  fmt.push(dvList(9, ['Days','Weeks','Months','Years']));
  fmt.push(dvList(13, ['Active','Paused','Complete','Archived']));

  // ── Formula columns (A = Template ID, L = Next Due Date) ─────────────────
  const nextDueFml = (r) =>
    `=IFERROR(IF(M${r}="",K${r},IF(G${r}="Daily",M${r}+1,IF(G${r}="Weekdays",M${r}+1,IF(G${r}="Weekly",M${r}+7,IF(G${r}="Biweekly",M${r}+14,IF(G${r}="Monthly",EDATE(M${r},1),IF(G${r}="Quarterly",EDATE(M${r},3),IF(G${r}="Semiannual",EDATE(M${r},6),IF(G${r}="Annual",EDATE(M${r},12),IF(G${r}="Custom Interval",M${r}+IF(J${r}="Weeks",I${r}*7,IF(J${r}="Months",I${r}*30,IF(J${r}="Years",I${r}*365,I${r}))),M${r})))))))))),"")`;

  const N_ROWS = 500;
  const idFormulas   = Array.from({ length: N_ROWS }, (_, i) => [`=IF(B${8+i}="","","REC-"&TEXT(ROW()-7,"000"))`]);
  const nextDueFmls  = Array.from({ length: N_ROWS }, (_, i) => [nextDueFml(8+i)]);
  vals.push({ range: `${S}!A8:A${7+N_ROWS}`, values: idFormulas });
  vals.push({ range: `${S}!L8:L${7+N_ROWS}`, values: nextDueFmls });

  // ── Number formats ────────────────────────────────────────────────────────
  [10, 11, 12].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, 14, 15), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Formula column backgrounds
  [0, 11].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Alternating rows
  for (let r = 0; r < 200; r++) {
    if (r % 2 !== 0) {
      fmt.push({ repeatCell: { range: gridRange(SID, 7+r, 8+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }

  // Row/col dimensions
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  const colWidths = [75,220,90,80,85,70,90,80,65,80,90,90,90,80,55,220];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, fmt, 'rec-fmt');
  vals.push({ range: `${S}!A8`, values: TEMPLATES });
  await valuesBatchUpdate(id, vals, 'rec-vals');
  console.log(`✓ Recurring Tasks complete — ${TEMPLATES.length} templates written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
