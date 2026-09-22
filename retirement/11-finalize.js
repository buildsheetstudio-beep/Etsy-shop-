'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const RD = sheetMap['Reference Data'];
const PS = sheetMap['Party Setup'];
const GL = sheetMap['Guest List & RSVP'];
const IT = sheetMap['Invitation Tracker'];
const BE = sheetMap['Budget & Expenses'];
const EM = sheetMap['Entertainment & Music'];
const TY = sheetMap['Thank-You Cards'];
const PP = sheetMap['Prep & Party Day'];
const DB = sheetMap['Dashboard'];

function colWidth(sheetId, colIndex, pixels) {
  return {
    updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: colIndex, endIndex: colIndex + 1 },
      properties: { pixelSize: pixels },
      fields: 'pixelSize'
    }
  };
}

function rowHeight(sheetId, rowIndex, pixels) {
  return {
    updateDimensionProperties: {
      range: { sheetId, dimension: 'ROWS', startIndex: rowIndex, endIndex: rowIndex + 1 },
      properties: { pixelSize: pixels },
      fields: 'pixelSize'
    }
  };
}

function rowHeightRange(sheetId, startIndex, endIndex, pixels) {
  return {
    updateDimensionProperties: {
      range: { sheetId, dimension: 'ROWS', startIndex, endIndex },
      properties: { pixelSize: pixels },
      fields: 'pixelSize'
    }
  };
}

function tabColor(sheetId, colorHex) {
  return {
    updateSheetProperties: {
      properties: { sheetId, tabColorStyle: { rgbColor: hex(colorHex) } },
      fields: 'tabColorStyle'
    }
  };
}

(async () => {
  const reqs = [];

  // ── Tab Colors ──────────────────────────────────────────────────────────────
  reqs.push(tabColor(RD, C.secText));       // Reference Data — muted grey
  reqs.push(tabColor(PS, C.primary));        // Party Setup — slate indigo
  reqs.push(tabColor(GL, C.mutedBlue));      // Guest List — muted blue
  reqs.push(tabColor(IT, C.mutedMauve));     // Invitation Tracker — muted mauve
  reqs.push(tabColor(BE, C.warning));        // Budget & Expenses — amber
  reqs.push(tabColor(EM, C.success));        // Entertainment & Music — sage
  reqs.push(tabColor(TY, C.secondary));      // Thank-You Cards — champagne copper
  reqs.push(tabColor(PP, C.attention));      // Prep & Party Day — rust
  reqs.push(tabColor(DB, C.secondary));      // Dashboard — champagne copper

  // ── Reference Data column widths ────────────────────────────────────────────
  for (let c = 0; c <= 24; c += 2) {
    reqs.push(colWidth(RD, c, 160));   // label cols
    reqs.push(colWidth(RD, c + 1, 20)); // spacer
  }

  // ── Party Setup column widths ────────────────────────────────────────────────
  reqs.push(colWidth(PS, 0, 30));   // A — index
  reqs.push(colWidth(PS, 1, 200));  // B — label
  reqs.push(colWidth(PS, 2, 220));  // C — value
  reqs.push(colWidth(PS, 3, 30));   // D — spacer
  reqs.push(colWidth(PS, 4, 130));  // E — summary label left
  reqs.push(colWidth(PS, 5, 130));  // F — summary label right
  reqs.push(colWidth(PS, 6, 110));  // G — value
  reqs.push(colWidth(PS, 7, 110));  // H — value continued

  // Party Setup row heights
  reqs.push(rowHeight(PS, 0, 8));
  reqs.push(rowHeight(PS, 1, 52));  // title
  reqs.push(rowHeight(PS, 2, 36));  // section header row
  reqs.push(rowHeight(PS, 3, 28));  // column header row
  reqs.push(rowHeightRange(PS, 4, 23, 28));  // data rows

  // ── Guest List column widths ─────────────────────────────────────────────────
  reqs.push(colWidth(GL, 0, 28));   // A — row num
  reqs.push(colWidth(GL, 1, 90));   // B — Guest ID
  reqs.push(colWidth(GL, 2, 160));  // C — First Name
  reqs.push(colWidth(GL, 3, 160));  // D — Last Name
  reqs.push(colWidth(GL, 4, 170));  // E — Household/Group
  reqs.push(colWidth(GL, 5, 90));   // F — Guest Type
  reqs.push(colWidth(GL, 6, 110));  // G — Relationship
  reqs.push(colWidth(GL, 7, 210));  // H — Email
  reqs.push(colWidth(GL, 8, 130));  // I — Phone
  reqs.push(colWidth(GL, 9, 100));  // J — RSVP Status
  reqs.push(colWidth(GL, 10, 90));  // K — RSVP Date
  reqs.push(colWidth(GL, 11, 80));  // L — Party Count
  reqs.push(colWidth(GL, 12, 80));  // M — Gift/Contribution (checkbox)
  reqs.push(colWidth(GL, 13, 180));  // N — Gift Description
  reqs.push(colWidth(GL, 14, 130));  // O — Dietary Restrictions
  reqs.push(colWidth(GL, 15, 80));  // P — Table Assignment
  reqs.push(colWidth(GL, 16, 80));  // Q — Thank-You Needed
  reqs.push(colWidth(GL, 17, 170));  // R — Notes

  // Guest List row heights
  reqs.push(rowHeight(GL, 0, 8));
  reqs.push(rowHeight(GL, 1, 52));
  reqs.push(rowHeight(GL, 2, 28));
  reqs.push(rowHeight(GL, 3, 28));
  reqs.push(rowHeight(GL, 4, 28));
  reqs.push(rowHeightRange(GL, 5, 45, 26));  // data rows

  // ── Invitation Tracker column widths ─────────────────────────────────────────
  reqs.push(colWidth(IT, 0, 80));   // A — Guest ID
  reqs.push(colWidth(IT, 1, 160));  // B — Guest Name (VLOOKUP)
  reqs.push(colWidth(IT, 2, 150));  // C — Household (VLOOKUP)
  reqs.push(colWidth(IT, 3, 110));  // D — Invitation Method
  reqs.push(colWidth(IT, 4, 80));   // E — Sent (checkbox)
  reqs.push(colWidth(IT, 5, 90));   // F — Date Sent
  reqs.push(colWidth(IT, 6, 90));   // G — Date Due
  reqs.push(colWidth(IT, 7, 100));  // H — RSVP Status (VLOOKUP)
  reqs.push(colWidth(IT, 8, 80));   // I — Follow-Up
  reqs.push(colWidth(IT, 9, 90));   // J — Follow-Up Date
  reqs.push(colWidth(IT, 10, 80));  // K — Reminder Sent (checkbox)
  reqs.push(colWidth(IT, 11, 90));  // L — Reminder Date
  reqs.push(colWidth(IT, 12, 130)); // M — Contact (VLOOKUP)
  reqs.push(colWidth(IT, 13, 170)); // N — Notes

  // Invitation row heights
  reqs.push(rowHeight(IT, 0, 8));
  reqs.push(rowHeight(IT, 1, 52));
  reqs.push(rowHeight(IT, 2, 28));
  reqs.push(rowHeight(IT, 3, 28));
  reqs.push(rowHeight(IT, 4, 28));
  reqs.push(rowHeightRange(IT, 5, 35, 26));

  // ── Budget & Expenses column widths ──────────────────────────────────────────
  reqs.push(colWidth(BE, 0, 150));  // A — Category
  reqs.push(colWidth(BE, 1, 100));  // B — Planned Amount
  reqs.push(colWidth(BE, 2, 100));  // C — Actual Spent
  reqs.push(colWidth(BE, 3, 80));   // D — Difference
  reqs.push(colWidth(BE, 4, 90));   // E — Status
  reqs.push(colWidth(BE, 5, 30));   // F — spacer
  reqs.push(colWidth(BE, 6, 30));   // G — spacer
  // Expense Log cols
  reqs.push(colWidth(BE, 7, 90));   // H — pivot label
  reqs.push(colWidth(BE, 8, 100));  // I — pivot amount
  reqs.push(colWidth(BE, 9, 30));   // J — spacer
  reqs.push(colWidth(BE, 10, 120)); // K — payment status
  reqs.push(colWidth(BE, 11, 100)); // L — payment total

  // Budget row heights
  reqs.push(rowHeight(BE, 0, 8));
  reqs.push(rowHeight(BE, 1, 52));
  reqs.push(rowHeight(BE, 2, 28));
  reqs.push(rowHeight(BE, 3, 28));
  reqs.push(rowHeight(BE, 4, 28));
  reqs.push(rowHeightRange(BE, 5, 30, 26));

  // ── Entertainment & Music column widths ───────────────────────────────────────
  reqs.push(colWidth(EM, 0, 80));   // A — #
  reqs.push(colWidth(EM, 1, 90));   // B — Ent ID
  reqs.push(colWidth(EM, 2, 160));  // C — Performer/Act
  reqs.push(colWidth(EM, 3, 110));  // D — Type
  reqs.push(colWidth(EM, 4, 90));   // E — Start Time
  reqs.push(colWidth(EM, 5, 90));   // F — End Time
  reqs.push(colWidth(EM, 6, 70));   // G — Duration
  reqs.push(colWidth(EM, 7, 100));  // H — Status
  reqs.push(colWidth(EM, 8, 100));  // I — Cost
  reqs.push(colWidth(EM, 9, 100));  // J — Deposit Paid
  reqs.push(colWidth(EM, 10, 100)); // K — Contact
  reqs.push(colWidth(EM, 11, 110)); // L — Phone/Email
  reqs.push(colWidth(EM, 12, 90));  // M — Contract Signed (checkbox)
  reqs.push(colWidth(EM, 13, 90));  // N — AV Needs (checkbox)
  reqs.push(colWidth(EM, 14, 170)); // O — Notes

  // Entertainment row heights
  reqs.push(rowHeight(EM, 0, 8));
  reqs.push(rowHeight(EM, 1, 52));
  reqs.push(rowHeight(EM, 2, 28));
  reqs.push(rowHeight(EM, 3, 28));
  reqs.push(rowHeight(EM, 4, 28));
  reqs.push(rowHeightRange(EM, 5, 100, 26));

  // ── Thank-You Cards column widths ────────────────────────────────────────────
  reqs.push(colWidth(TY, 0, 28));   // A — #
  reqs.push(colWidth(TY, 1, 80));   // B — Guest ID
  reqs.push(colWidth(TY, 2, 80));   // C — Thank-You ID
  reqs.push(colWidth(TY, 3, 150));  // D — Guest Name (VLOOKUP)
  reqs.push(colWidth(TY, 4, 150));  // E — Household (VLOOKUP)
  reqs.push(colWidth(TY, 5, 80));   // F — Gift Received (VLOOKUP)
  reqs.push(colWidth(TY, 6, 180));  // G — Gift Description (VLOOKUP)
  reqs.push(colWidth(TY, 7, 90));   // H — Card Status
  reqs.push(colWidth(TY, 8, 90));   // I — Date Sent
  reqs.push(colWidth(TY, 9, 90));   // J — Days Since Party
  reqs.push(colWidth(TY, 10, 80));  // K — Email Thank-You (checkbox)
  reqs.push(colWidth(TY, 11, 80));  // L — Email Sent Date
  reqs.push(colWidth(TY, 12, 80));  // M — Handwritten (checkbox)
  reqs.push(colWidth(TY, 13, 90));  // N — Personal Note Added (checkbox)
  reqs.push(colWidth(TY, 14, 80));  // O — Priority
  reqs.push(colWidth(TY, 15, 170)); // P — Notes

  // Thank-You row heights
  reqs.push(rowHeight(TY, 0, 8));
  reqs.push(rowHeight(TY, 1, 52));
  reqs.push(rowHeight(TY, 2, 28));
  reqs.push(rowHeight(TY, 3, 28));
  reqs.push(rowHeight(TY, 4, 28));
  reqs.push(rowHeight(TY, 5, 28));
  reqs.push(rowHeightRange(TY, 6, 35, 26));

  // ── Prep & Party Day column widths ───────────────────────────────────────────
  reqs.push(colWidth(PP, 0, 28));   // A — #
  reqs.push(colWidth(PP, 1, 90));   // B — Task ID
  reqs.push(colWidth(PP, 2, 220));  // C — Task Name
  reqs.push(colWidth(PP, 3, 110));  // D — Phase
  reqs.push(colWidth(PP, 4, 90));   // E — Priority
  reqs.push(colWidth(PP, 5, 100));  // F — Assigned To
  reqs.push(colWidth(PP, 6, 90));   // G — Due Date
  reqs.push(colWidth(PP, 7, 80));   // H — Days Remaining
  reqs.push(colWidth(PP, 8, 80));   // I — Overdue
  reqs.push(colWidth(PP, 9, 80));   // J — Done (checkbox)
  reqs.push(colWidth(PP, 10, 80));  // K — Completed Date
  reqs.push(colWidth(PP, 11, 170)); // L — Notes

  // Prep row heights
  reqs.push(rowHeight(PP, 0, 8));
  reqs.push(rowHeight(PP, 1, 52));
  reqs.push(rowHeight(PP, 2, 28));
  reqs.push(rowHeight(PP, 3, 28));
  reqs.push(rowHeight(PP, 4, 28));
  reqs.push(rowHeight(PP, 5, 28));
  reqs.push(rowHeight(PP, 6, 28));
  reqs.push(rowHeight(PP, 7, 28));
  reqs.push(rowHeightRange(PP, 8, 60, 26));

  // ── Dashboard column widths ────────────────────────────────────────────────
  // 12 main cols (A-L) each 80px wide, plus pivot cols J-L override
  for (let c = 0; c < 9; c++) {
    reqs.push(colWidth(DB, c, 90));
  }
  reqs.push(colWidth(DB, 9, 30));    // J — spacer before pivot
  reqs.push(colWidth(DB, 10, 160));  // K — pivot label
  reqs.push(colWidth(DB, 11, 110));  // L — pivot value

  // Dashboard row heights
  reqs.push(rowHeight(DB, 0, 8));
  reqs.push(rowHeight(DB, 1, 52));   // title
  reqs.push(rowHeight(DB, 2, 28));   // subtitle
  reqs.push(rowHeight(DB, 3, 14));   // spacer
  reqs.push(rowHeight(DB, 4, 28));   // KPI label row
  reqs.push(rowHeight(DB, 5, 48));   // KPI value row
  reqs.push(rowHeight(DB, 6, 14));   // spacer
  reqs.push(rowHeight(DB, 7, 28));   // KPI label row 2
  reqs.push(rowHeight(DB, 8, 48));   // KPI value row 2
  reqs.push(rowHeight(DB, 9, 14));   // spacer
  reqs.push(rowHeight(DB, 10, 28));  // KPI label row 3
  reqs.push(rowHeight(DB, 11, 48));  // KPI value row 3
  reqs.push(rowHeight(DB, 12, 14));  // spacer
  reqs.push(rowHeight(DB, 13, 28));  // KPI label row 4
  reqs.push(rowHeight(DB, 14, 48));  // KPI value row 4
  reqs.push(rowHeight(DB, 15, 14));  // spacer
  reqs.push(rowHeight(DB, 16, 28));  // section header
  reqs.push(rowHeight(DB, 17, 26));  // column header
  reqs.push(rowHeightRange(DB, 18, 29, 26));  // upcoming tasks
  reqs.push(rowHeight(DB, 29, 14));  // spacer
  reqs.push(rowHeight(DB, 30, 28));  // section header
  reqs.push(rowHeight(DB, 31, 26));  // column header
  reqs.push(rowHeightRange(DB, 32, 39, 26));  // entertainment snapshot
  reqs.push(rowHeight(DB, 39, 14));  // spacer
  reqs.push(rowHeight(DB, 40, 28));  // section header
  reqs.push(rowHeight(DB, 41, 26));  // column header
  reqs.push(rowHeightRange(DB, 42, 50, 26));  // weights table

  await batchUpdate(id, reqs);
  console.log('Finalize complete — column widths, row heights, and tab colors applied');
})();
