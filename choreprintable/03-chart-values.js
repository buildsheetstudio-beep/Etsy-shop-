'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const T = "'Weekly Chore Chart'";

// Column layout (0-indexed):
// A=0 (person label)
// Mon: B=1(chk) C=2 D=3(desc merged C:D)
// Tue: E=4(chk) F=5 G=6
// Wed: H=7(chk) I=8 J=9
// Thu: K=10(chk) L=11 M=12
// Fri: N=13(chk) O=14 P=15
// Sat: Q=16(chk) R=17 S=18
// Sun: T=19(chk) U=20 V=21

// Sample chores [member][row0..3][day0..6]  — just the task text (left desc cell)
// day order: Mon,Tue,Wed,Thu,Fri,Sat,Sun
const SAMPLES = [
  // Member 1 — Mom (rows 14-17)
  [
    ['Make bed','Complete homework','Wipe kitchen counters','','','Help with yard work',''],
    ['','','','Pack lunch','Help prepare dinner','','Water plants'],
    ['Fold laundry','','','','','',''],
    ['','Unload dishwasher','','','','',''],
  ],
  // Member 2 — Dad (rows 18-21)
  [
    ['Take out trash','','','','','Help with yard work','Bring in mail'],
    ['','','Vacuum','','','',''],
    ['','','','','','',''],
    ['','','','','','',''],
  ],
  // Member 3 — Child 1 (rows 22-25)
  [
    ['Make bed','Make bed','Make bed','Make bed','Make bed','Make bed','Make bed'],
    ['Complete homework','Complete homework','Complete homework','Complete homework','Complete homework','',''],
    ['','','','','','','Read for 20 minutes'],
    ['Pack school bag','Pack school bag','Pack school bag','Pack school bag','Pack school bag','',''],
  ],
  // Member 4 — Child 2 (rows 26-29)
  [
    ['Make bed','Make bed','Make bed','Make bed','Make bed','Make bed','Make bed'],
    ['Feed pet','Feed pet','Feed pet','Feed pet','Feed pet','Feed pet','Feed pet'],
    ['Complete homework','Complete homework','Complete homework','Complete homework','Complete homework','',''],
    ['','','','','','Walk dog','Walk dog'],
  ],
  // Member 5 — Child 3 (rows 30-33)
  [
    ['Make bed','Make bed','Make bed','Make bed','Make bed','Make bed','Make bed'],
    ['Put clothes away','','Put clothes away','','Put clothes away','',''],
    ['','Wipe bathroom mirror','','Clean bathroom sink','','',''],
    ['','','','','','','Read for 20 minutes'],
  ],
  // Member 6 — Child 4 (rows 34-37)
  [
    ['Make bed','Make bed','Make bed','Make bed','Make bed','Make bed','Make bed'],
    ['Refill water bowl','Refill water bowl','Refill water bowl','Refill water bowl','Refill water bowl','Refill water bowl','Refill water bowl'],
    ['','Sort laundry','','','Put laundry away','',''],
    ['','','','','','',''],
  ],
];

// Checkboxes to pre-check: [member][row][day] → TRUE
// Show some completed chores for visual demo
const CHECKED = [
  // Mom
  [[0,0],[0,1],[0,2],[1,3],[1,4]],
  // Dad
  [[0,0],[0,6],[0,5]],
  // Child 1
  [[0,0],[0,1],[0,2],[0,3],[0,4],[1,0],[1,1],[1,2],[1,3],[1,4],[3,0],[3,1],[3,2],[3,3],[3,4]],
  // Child 2
  [[0,0],[0,1],[0,2],[0,3],[0,4],[1,0],[1,1],[1,2],[1,3],[1,4],[1,5],[1,6]],
  // Child 3
  [[0,0],[0,1],[0,2],[0,3],[0,4],[1,0],[1,2],[1,4]],
  // Child 6
  [[0,0],[0,1],[0,2],[0,3],[0,4],[1,0],[1,1],[1,2],[1,3],[1,4],[1,5],[1,6]],
];

// Checkbox col letters for Mon-Sun
const CHK_COLS = ['B','E','H','K','N','Q','T'];
// Task (left desc cell) col letters
const TASK_COLS = ['C','F','I','L','O','R','U'];

const MEM_START = [14,18,22,26,30,34]; // 1-indexed first data row per member

(async () => {
  const data = [];

  // ── Rows 1-2: Title ───────────────────────────────────────────────────────
  data.push({ range: `${T}!A1`, values: [['PRINTABLE FAMILY CHORE CHART']] });

  // ── Row 3: Subtitle ───────────────────────────────────────────────────────
  data.push({ range: `${T}!A3`, values: [['UP TO 6 FAMILY MEMBERS • PRINTABLE • REUSABLE EVERY WEEK']] });

  // ── Row 5-7: Setup labels ─────────────────────────────────────────────────
  data.push({ range: `${T}!A5`, values: [['CHORE CHART']] });
  data.push({ range: `${T}!G5`, values: [['WEEK START DATE']] });
  data.push({ range: `${T}!G6`, values: [['2025-07-07']] });  // sample Monday

  // Summary card labels and formulas
  data.push({ range: `${T}!K5`, values: [['TOTAL CHORES']] });
  data.push({ range: `${T}!N5`, values: [['COMPLETED']] });
  data.push({ range: `${T}!Q5`, values: [['REMAINING']] });
  data.push({ range: `${T}!T5`, values: [['COMPLETION %']] });

  // Summary formulas (large value cells)
  const totalF = `=COUNTA(C14:C37,F14:F37,I14:I37,L14:L37,O14:O37,R14:R37,U14:U37)`;
  const compF  = `=COUNTIFS(B14:B37,TRUE,C14:C37,"<>")+COUNTIFS(E14:E37,TRUE,F14:F37,"<>")+COUNTIFS(H14:H37,TRUE,I14:I37,"<>")+COUNTIFS(K14:K37,TRUE,L14:L37,"<>")+COUNTIFS(N14:N37,TRUE,O14:O37,"<>")+COUNTIFS(Q14:Q37,TRUE,R14:R37,"<>")+COUNTIFS(T14:T37,TRUE,U14:U37,"<>")`;
  const remF   = `=MAX(0,K6-N6)`;
  const pctF   = `=IFERROR(N6/K6,0)`;
  data.push({ range: `${T}!K6`, values: [[totalF]] });
  data.push({ range: `${T}!N6`, values: [[compF]] });
  data.push({ range: `${T}!Q6`, values: [[remF]] });
  data.push({ range: `${T}!T6`, values: [[pctF]] });

  // ── Row 9-11: Family member names ─────────────────────────────────────────
  data.push({ range: `${T}!A9`,  values: [['FAMILY MEMBERS']] });
  data.push({ range: `${T}!D9`,  values: [["='Reference Data'!B2"]] });
  data.push({ range: `${T}!G9`,  values: [["='Reference Data'!B3"]] });
  data.push({ range: `${T}!J9`,  values: [["='Reference Data'!B4"]] });
  data.push({ range: `${T}!M9`,  values: [["='Reference Data'!B5"]] });
  data.push({ range: `${T}!P9`,  values: [["='Reference Data'!B6"]] });
  data.push({ range: `${T}!S9`,  values: [["='Reference Data'!B7"]] });

  // ── Row 13: Day headers ───────────────────────────────────────────────────
  data.push({ range: `${T}!A13`, values: [['PERSON']] });
  data.push({ range: `${T}!B13`, values: [['MONDAY']] });
  data.push({ range: `${T}!E13`, values: [['TUESDAY']] });
  data.push({ range: `${T}!H13`, values: [['WEDNESDAY']] });
  data.push({ range: `${T}!K13`, values: [['THURSDAY']] });
  data.push({ range: `${T}!N13`, values: [['FRIDAY']] });
  data.push({ range: `${T}!Q13`, values: [['SATURDAY']] });
  data.push({ range: `${T}!T13`, values: [['SUNDAY']] });

  // ── Member labels in column A ─────────────────────────────────────────────
  const memRefs = ["='Reference Data'!B2","='Reference Data'!B3","='Reference Data'!B4",
                   "='Reference Data'!B5","='Reference Data'!B6","='Reference Data'!B7"];
  for (let m = 0; m < 6; m++) {
    const startRow = MEM_START[m];
    data.push({ range: `${T}!A${startRow}`, values: [[memRefs[m]]] });
  }

  // ── Sample chore tasks ────────────────────────────────────────────────────
  for (let m = 0; m < 6; m++) {
    const startRow = MEM_START[m];
    for (let r = 0; r < 4; r++) {
      const row = startRow + r;
      for (let d = 0; d < 7; d++) {
        const task = SAMPLES[m][r][d] || '';
        if (task) {
          data.push({ range: `${T}!${TASK_COLS[d]}${row}`, values: [[task]] });
        }
      }
    }
  }

  // ── Pre-checked checkboxes ────────────────────────────────────────────────
  for (let m = 0; m < 6; m++) {
    const startRow = MEM_START[m];
    const checks = CHECKED[m] || [];
    for (const [r, d] of checks) {
      const row = startRow + r;
      data.push({ range: `${T}!${CHK_COLS[d]}${row}`, values: [['TRUE']] });
    }
  }

  // ── Notes label ───────────────────────────────────────────────────────────
  data.push({ range: `${T}!A39`, values: [['NOTES']] });

  // ── Footer ────────────────────────────────────────────────────────────────
  data.push({ range: `${T}!A46`, values: [['INSTANT DOWNLOAD • 1 EASY-TO-USE TAB • GOOGLE SHEETS + EXCEL FRIENDLY']] });

  await valuesBatchUpdate(id, data, 'chart-values');
  console.log('Chart values complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
