'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Team Dashboard'];

// Root cause: the original 09-dashboard.js applied two bad merges in order:
//   1. gridRange(DB,0,2,...) — merged rows 1-2 (title + subtitle row) into one cell.
//      Subtitle text at A2 (index 1) was absorbed and lost; only A1's title shows.
//   2. gridRange(DB,2,4,...) — merged rows 3-4 (the controls row) into one blank cell.
//      B4=year, D4=period, F4=member, H4=dept, J4=project were ALL deleted by this merge.
//
// KPI formulas reference $B$4, $F$4, $H$4, $J$4 → all return "" → year filter is
// YEAR(range)="" → FALSE for every row → SUMPRODUCT = 0.
//
// Fix:
//   1. Unmerge both bad merges
//   2. Re-merge title as just row 1 (index 0)
//   3. Re-merge subtitle as rows 2-3 (indices 1-2), master cell A2
//   4. Re-apply formatting for title, subtitle, and controls rows
//   5. Push title, subtitle text, and all control values back

(async () => {
  const reqs = [];

  // ── 1. Unmerge the two bad merges ─────────────────────────────────────────
  reqs.push({ unmergeCells: { range: gridRange(DB, 0, 2, 0, 12) } }); // was title rows 1-2
  reqs.push({ unmergeCells: { range: gridRange(DB, 2, 4, 0, 12) } }); // was subtitle rows 3-4 (swallowed controls)

  // ── 2. Re-merge correctly ─────────────────────────────────────────────────
  // Title: just row 1 (index 0)
  reqs.push({ mergeCells: { range: gridRange(DB, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  // Subtitle: rows 2-3 (indices 1-2), so A2 is the master cell
  reqs.push({ mergeCells: { range: gridRange(DB, 1, 3, 0, 12), mergeType: 'MERGE_ALL' } });

  // ── 3. Re-apply row formatting ────────────────────────────────────────────
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(DB,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });

  // Title row (index 0 only)
  fmt(0,1,0,12,{ userEnteredFormat:{
    backgroundColor: hex(C.primary),
    textFormat:{ foregroundColor: hex(C.white), bold:true, fontSize:22, fontFamily:'Arial' },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
  }});

  // Subtitle rows 2-3 (indices 1-2)
  fmt(1,3,0,12,{ userEnteredFormat:{
    backgroundColor: hex(C.bg),
    textFormat:{ foregroundColor: hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
  }});

  // Controls row 4 (index 3) — 5 label/value pairs across A-J
  for (let i = 0; i < 5; i++) {
    const c = i * 2;
    fmt(3,4,c,c+1,{ userEnteredFormat:{
      backgroundColor: hex(C.secondary),
      textFormat:{ foregroundColor: hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' },
      horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE',
    }});
    fmt(3,4,c+1,c+2,{ userEnteredFormat:{
      backgroundColor: hex(C.input),
      textFormat:{ fontSize:10, fontFamily:'Arial' },
      verticalAlignment:'MIDDLE',
    }});
  }

  await batchUpdate(id, reqs, 'fix-db-layout');

  // ── 4. Re-push title, subtitle text, and control values ──────────────────
  const data = [];
  data.push({ range: `'Team Dashboard'!A1`, values: [['TEAM DASHBOARD']] });
  data.push({ range: `'Team Dashboard'!A2`, values: [['Your real-time command center for team task management. Use the controls below to filter by year, period, member, department, or project.']] });

  // Controls — B4=year, D4=period, F4=member, H4=dept, J4=project
  data.push({ range: `'Team Dashboard'!A4`, values: [['Year:']] });
  data.push({ range: `'Team Dashboard'!B4`, values: [[2026]] });
  data.push({ range: `'Team Dashboard'!C4`, values: [['Period:']] });
  data.push({ range: `'Team Dashboard'!D4`, values: [['All Time']] });
  data.push({ range: `'Team Dashboard'!E4`, values: [['Team Member:']] });
  data.push({ range: `'Team Dashboard'!F4`, values: [['All Team Members']] });
  data.push({ range: `'Team Dashboard'!G4`, values: [['Department:']] });
  data.push({ range: `'Team Dashboard'!H4`, values: [['All Departments']] });
  data.push({ range: `'Team Dashboard'!I4`, values: [['Project:']] });
  data.push({ range: `'Team Dashboard'!J4`, values: [['All Projects']] });

  await valuesBatchUpdate(id, data, 'fix-db-controls');

  console.log('Dashboard layout fixed — control cells restored');
  console.log('KPI formulas reference $B$4 (year), $F$4 (member), $H$4 (dept), $J$4 (project)');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
