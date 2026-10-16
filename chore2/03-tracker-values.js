'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// ── Column layout (0-indexed) ──────────────────────────────────────────────
//  A=0  label
//  B=1  Sun Hrs   C=2  Sun Allow
//  D=3  Mon Hrs   E=4  Mon Allow
//  F=5  Tue Hrs   G=6  Tue Allow
//  H=7  Wed Hrs   I=8  Wed Allow
//  J=9  Thu Hrs   K=10 Thu Allow
//  L=11 Fri Hrs   M=12 Fri Allow
//  N=13 Sat Hrs   O=14 Sat Allow
//  P=15 Wk Total Hrs   Q=16 Wk Total Allow
//
// Header rows 1-12 (1-indexed):
//   Row 1: Title (A1) | Rate table header (F1) | Family member header (I1)
//   Row 2: Start Date (A2:B2) | Col hdrs F2:H2 | Col hdrs I2:K2
//   Rows 3-8: Rate table (F3:H8) | Family members 1-6 (I3:K8)
//   Rows 9-12: Family members 7-10 (I9:K12)
//
// Rate table: F=Chore Name, G=Rate, H=Type (VLOOKUP range $F$3:$H$8)
// Family block: I=Name, J=Weekly Hours (formula), K=Weekly Allowance (formula)
//
// Grid rows:
//   Row 13: day headers  Row 14: H/$ sub-headers
//   Row 15: TOTAL header  Rows 16-21: TOTAL chore rows
//   Row 22+: member blocks (1 header + 6 chore rows each, 7 rows per slot × 10 slots)

const CHORES = ['Feeding the pet','Mow the lawn','Put dishes away','Laundry','Homework check-in','Make your bed'];
const RATES  = [3.50, 8.00, 2.00, 1.50, 0.75, 1.00];
const TYPES  = ['Hourly','Flat Rate','Flat Rate','Hourly','Hourly','Flat Rate'];

const DAYS   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const DAY_H  = ['B','D','F','H','J','L','N'];   // hour cols per day
const DAY_A  = ['C','E','G','I','K','M','O'];   // allowance cols per day

// Base row (first chore row) for each of 10 member slots
// Slot m (0-indexed): header = 22 + m*7, chores = 23 + m*7 … 28 + m*7
const MEM_BASE = m => 23 + m * 7;

const MEMBERS = ['Anna','Peter','John','','','','','','',''];
// Hours[member][chore][day] — 0-indexed: chore 0-5, day Sun-Sat
const HOURS = {
  Anna: [
    [0,1,0,1,0,1,0],  // Feeding the pet
    [0,0,0,0,0,0,1],  // Mow the lawn (flat, counts=instances)
    [0,1,0,1,0,1,0],  // Put dishes away
    [0,0,2,0,0,0,0],  // Laundry
    [0,1,1,1,1,1,0],  // Homework check-in
    [1,1,1,1,1,1,1],  // Make your bed
  ],
  Peter: [
    [0,0,1,0,1,0,0],
    [1,0,0,0,0,0,0],
    [0,1,1,0,0,0,1],
    [1.5,0,0,0,1.5,0,0],
    [0,0.5,0,0.5,0,0.5,0],
    [0,1,1,1,1,1,0],
  ],
  John: [
    [1,0,0,1,0,0,1],
    [0,0,0,0,0,0,0],
    [1,0,1,0,1,0,0],
    [0,2,0,0,0,0,0],
    [0,0,1,1,1,0,0],
    [0,1,1,1,1,1,1],
  ],
};
const MEM_HOURS = [HOURS.Anna, HOURS.Peter, HOURS.John,
                   null,null,null,null,null,null,null];

// ── Formula helpers ────────────────────────────────────────────────────────

function allowF(row, d) {
  const h = DAY_H[d];
  // Hourly: h*rate | Flat Rate: if h>0 then rate else 0
  return `=IFERROR(IF(VLOOKUP($A${row},$F$3:$H$8,3,FALSE)="Hourly",${h}${row}*VLOOKUP($A${row},$F$3:$H$8,2,FALSE),IF(${h}${row}>0,VLOOKUP($A${row},$F$3:$H$8,2,FALSE),0)),0)`;
}

function wkHrs(row)   { return `=SUM(${DAY_H.map(c=>`${c}${row}`).join(',')})`;  }
function wkAllow(row) { return `=SUM(${DAY_A.map(c=>`${c}${row}`).join(',')})`;  }

function totalH(c, d) {
  const col = DAY_H[d];
  return `=SUM(${[0,1,2,3,4,5,6,7,8,9].map(m=>`${col}${MEM_BASE(m)+c}`).join(',')})`;
}
function totalA(c, d) {
  const col = DAY_A[d];
  return `=SUM(${[0,1,2,3,4,5,6,7,8,9].map(m=>`${col}${MEM_BASE(m)+c}`).join(',')})`;
}

function b3Hrs(base) {
  return `SUM(${DAY_H.map(c=>`${c}${base}:${c}${base+5}`).join(',')})`;
}
function b3Allow(base) {
  return `SUM(${DAY_A.map(c=>`${c}${base}:${c}${base+5}`).join(',')})`;
}

// Build a 17-cell row (A-Q) for a member chore row
function memberRow(choreName, hrs, row) {
  const v = [choreName];
  for (let d = 0; d < 7; d++) {
    v.push(hrs ? (hrs[d] || 0) : 0);
    v.push(allowF(row, d));
  }
  v.push(wkHrs(row));
  v.push(wkAllow(row));
  return v;
}

// Build a 17-cell row for a TOTAL chore row
function totalRow(choreName, c, row) {
  const v = [choreName];
  for (let d = 0; d < 7; d++) {
    v.push(totalH(c, d));
    v.push(totalA(c, d));
  }
  v.push(wkHrs(row));
  v.push(wkAllow(row));
  return v;
}

(async () => {
  const T = "'Chore & Allowance Tracker'";
  const data = [];

  // ── Header rows 1-12 ────────────────────────────────────────────────────
  // Row 1 block labels
  data.push({ range: `${T}!A1`, values: [['KIDS CHORE CHART & ALLOWANCE']] });
  data.push({ range: `${T}!F1`, values: [['CHORE RATE TABLE']] });
  data.push({ range: `${T}!I1`, values: [['FAMILY MEMBERS']] });

  // Row 2
  data.push({ range: `${T}!A2:B2`, values: [['Start Date:','2025-06-01']] });
  data.push({ range: `${T}!F2:H2`, values: [['Chore Name','Rate','Type']] });
  data.push({ range: `${T}!I2:K2`, values: [['Family Member','Weekly Hours','Weekly Allowance']] });

  // Rate table rows 3-8
  for (let i = 0; i < 6; i++) {
    data.push({ range: `${T}!F${i+3}:H${i+3}`, values: [[CHORES[i], RATES[i], TYPES[i]]] });
  }

  // Family member block rows 3-12
  for (let m = 0; m < 10; m++) {
    const hr = m + 3;  // header row for this member in block 3
    const base = MEM_BASE(m);
    data.push({
      range: `${T}!I${hr}:K${hr}`,
      values: [[
        MEMBERS[m],
        `=IF(I${hr}="","",${b3Hrs(base)})`,
        `=IF(I${hr}="","",${b3Allow(base)})`
      ]]
    });
  }

  // ── Grid header rows ────────────────────────────────────────────────────
  // Row 13: day headers
  const dayRow = ['Person / Chore'];
  for (const d of DAYS) dayRow.push(d, '');
  dayRow.push('Weekly Achievement', '');
  data.push({ range: `${T}!A13:Q13`, values: [dayRow] });

  // Row 14: H/$ sub-headers
  const subRow = [''];
  for (let d = 0; d < 7; d++) subRow.push('Hrs', 'Allow.');
  subRow.push('Total Hrs', 'Total Allow.');
  data.push({ range: `${T}!A14:Q14`, values: [subRow] });

  // ── TOTAL block rows 15-21 ──────────────────────────────────────────────
  data.push({ range: `${T}!A15`, values: [['TOTAL']] });
  for (let c = 0; c < 6; c++) {
    data.push({ range: `${T}!A${16+c}:Q${16+c}`, values: [totalRow(CHORES[c], c, 16+c)] });
  }

  // ── Member blocks rows 22-91 ────────────────────────────────────────────
  const MEM_NAMES_GRID = ['ANNA','PETER','JOHN','','','','','','',''];
  for (let m = 0; m < 10; m++) {
    const headerRow = 22 + m * 7;
    const base = MEM_BASE(m);
    // Member section header (label in col A only — no cross-boundary merge)
    data.push({ range: `${T}!A${headerRow}`, values: [[MEM_NAMES_GRID[m]]] });
    // 6 chore rows
    for (let c = 0; c < 6; c++) {
      const row = base + c;
      data.push({ range: `${T}!A${row}:Q${row}`, values: [memberRow(CHORES[c], MEM_HOURS[m] ? MEM_HOURS[m][c] : null, row)] });
    }
  }

  await valuesBatchUpdate(id, data, 'tracker-values');
  console.log('Tracker values complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
