'use strict';
// Two fixes:
// 1. Master Event Log column X ("Next Occurrence") — format cells as dates instead of serials.
// 2. Weekly Schedule + Monthly Calendar — replace the single-purple CF rules with
//    13 category-specific pastel colors drawn from Master Event Log column G, plus
//    a visible color legend row on each tab so viewers know what each color means.
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const CHUNK = 400;
const VCHUNK = 200;

// 13 distinct pastel swatches — one per category in first-seen order
const PALETTE = [
  { red: 0.682, green: 0.839, blue: 0.863 }, // light blue     – Conference
  { red: 0.961, green: 0.773, blue: 0.639 }, // light orange   – Fundraiser
  { red: 0.710, green: 0.835, blue: 0.659 }, // light green    – Community
  { red: 0.961, green: 0.875, blue: 0.659 }, // light amber    – Workshop
  { red: 0.659, green: 0.835, blue: 0.773 }, // light teal     – Business
  { red: 0.835, green: 0.918, blue: 0.659 }, // light lime     – Sports
  { red: 0.773, green: 0.722, blue: 0.910 }, // light violet   – Virtual
  { red: 0.941, green: 0.722, blue: 0.784 }, // light rose     – Networking
  { red: 0.659, green: 0.835, blue: 0.918 }, // light sky      – School
  { red: 0.941, green: 0.784, blue: 0.847 }, // light pink     – Birthday
  { red: 0.941, green: 0.898, blue: 0.659 }, // light gold     – Holiday
  { red: 0.784, green: 0.847, blue: 0.722 }, // light sage     – Nonprofit
  { red: 0.847, green: 0.784, blue: 0.910 }, // light lavender – Social
];

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

(async () => {
  const sheets = await getSheets();

  // ─── Metadata: IDs + existing CF counts ──────────────────────────────────
  console.log('Reading spreadsheet metadata...');
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(sheetId,title),conditionalFormats)',
  });
  const sheetMap = {};
  const cfCount  = {};
  for (const s of meta.data.sheets) {
    sheetMap[s.properties.title] = s.properties.sheetId;
    cfCount[s.properties.sheetId] = (s.conditionalFormats || []).length;
  }
  const wsId  = sheetMap['Weekly Schedule'];
  const mcId  = sheetMap['Monthly Calendar'];
  const melId = sheetMap['Master Event Log'];
  console.log(`  WS=${wsId}(${cfCount[wsId]} CF) MC=${mcId}(${cfCount[mcId]} CF) MEL=${melId}`);

  // ─── Read time-slot rows from WS col A ───────────────────────────────────
  const wsColAResp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Weekly Schedule'!A1:A80",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const wsColA = (wsColAResp.data.values || []).map(r => r[0]);
  const tsRows = []; // 0-indexed time-slot rows
  for (let i = 0; i < wsColA.length; i++) {
    const v = wsColA[i];
    if (typeof v === 'number' && v > 0 && v < 1) tsRows.push(i);
  }
  const wsFirst = tsRows[0];
  const wsLast  = tsRows[tsRows.length - 1];
  console.log(`  ${tsRows.length} time-slot rows: ${wsFirst}–${wsLast} (0-indexed)`);

  // ─── Unique categories from MEL col G (in first-seen order) ──────────────
  const catResp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Master Event Log'!G6:G500",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const seen = new Set();
  const cats = [];
  for (const row of (catResp.data.values || [])) {
    const v = row[0];
    if (v && !seen.has(v)) { seen.add(v); cats.push(v); }
  }
  console.log(`  ${cats.length} categories: ${cats.join(', ')}`);
  const catColor = {};
  cats.forEach((c, i) => { catColor[c] = PALETTE[i % PALETTE.length]; });

  const batch1      = []; // MEL format + CF deletions (committed first)
  const batch2      = []; // legend formatting + new CF rules
  const valueUpdates = [];

  // ─── FIX 1: Format MEL col X as date ──────────────────────────────────────
  batch1.push({
    repeatCell: {
      range: {
        sheetId: melId,
        startRowIndex: 5, endRowIndex: 500,
        startColumnIndex: 23, endColumnIndex: 24, // column X
      },
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ─── FIX 2: Delete ALL existing CF rules from WS and MC ───────────────────
  for (const [sid, cnt] of [[wsId, cfCount[wsId]], [mcId, cfCount[mcId]]]) {
    for (let idx = cnt - 1; idx >= 0; idx--) {
      batch1.push({ deleteConditionalFormatRule: { sheetId: sid, index: idx } });
    }
  }
  console.log(`  Queued ${cfCount[wsId] + cfCount[mcId]} CF deletions in batch1.`);

  // ─── FIX 3: WS color legend — row 5 (currently empty) ────────────────────
  valueUpdates.push({ range: "'Weekly Schedule'!A5", values: [['Category Colors:']] });
  cats.forEach((cat, i) => {
    const col = String.fromCharCode(66 + i); // B(0) … N(12)
    valueUpdates.push({ range: `'Weekly Schedule'!${col}5`, values: [[cat]] });
    batch2.push({
      repeatCell: {
        range: {
          sheetId: wsId,
          startRowIndex: 4, endRowIndex: 5,
          startColumnIndex: 1 + i, endColumnIndex: 2 + i,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: catColor[cat],
            horizontalAlignment: 'CENTER',
            textFormat: { bold: true, fontSize: 8 },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)',
      },
    });
  });

  // ─── FIX 4: MC color legend — row 3, cols H+ ──────────────────────────────
  valueUpdates.push({ range: "'Monthly Calendar'!H3", values: [['Legend:']] });
  batch2.push({
    repeatCell: {
      range: { sheetId: mcId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 7, endColumnIndex: 8 },
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat',
    },
  });
  cats.forEach((cat, i) => {
    const col = String.fromCharCode(73 + i); // I(0) … U(12)
    valueUpdates.push({ range: `'Monthly Calendar'!${col}3`, values: [[cat]] });
    batch2.push({
      repeatCell: {
        range: {
          sheetId: mcId,
          startRowIndex: 2, endRowIndex: 3,
          startColumnIndex: 8 + i, endColumnIndex: 9 + i,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: catColor[cat],
            horizontalAlignment: 'CENTER',
            textFormat: { bold: true, fontSize: 8 },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)',
      },
    });
  });

  // ─── FIX 5: WS category CF rules ─────────────────────────────────────────
  // COUNTIFS with INDIRECT cross-sheet refs — direct sheet refs are rejected by
  // the Sheets CF validator but INDIRECT is accepted.
  // INT((COLUMN()-2)/2) maps col B/C→0, D/E→1, …, N/O→6 (day-of-week offset).
  for (let ci = cats.length - 1; ci >= 0; ci--) {
    const cat   = cats[ci];
    const color = catColor[cat];
    const fml =
      `=COUNTIFS(INDIRECT("'Master Event Log'!D6:D500"),$B$4+INT((COLUMN()-2)/2),` +
      `INDIRECT("'Master Event Log'!G6:G500"),"${cat}")>0`;
    batch2.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{
            sheetId: wsId,
            startRowIndex: wsFirst, endRowIndex: wsLast + 1,
            startColumnIndex: 1, endColumnIndex: 15,
          }],
          booleanRule: {
            condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: fml }] },
            format: { backgroundColor: color },
          },
        },
        index: 0,
      },
    });
  }

  // ─── FIX 6: MC category CF rules ─────────────────────────────────────────
  // COUNTIFS: date criterion computed from month anchor + ROW/COLUMN offsets.
  const monthHeaderRows = [4, 13, 22, 31, 40, 49, 58, 67, 76, 85, 94, 103];

  for (const mhRow of monthHeaderRows) {
    const anchor1 = mhRow + 2; // 1-indexed first grid row
    for (let ci = cats.length - 1; ci >= 0; ci--) {
      const cat   = cats[ci];
      const color = catColor[cat];
      const dateCrit =
        `$A$${mhRow}+((ROW()-${anchor1})*7+(COLUMN()-1))-(WEEKDAY($A$${mhRow},1)-1)`;
      const fml =
        `=COUNTIFS(INDIRECT("'Master Event Log'!D6:D500"),${dateCrit},` +
        `INDIRECT("'Master Event Log'!G6:G500"),"${cat}")>0`;
      batch2.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: mcId,
              startRowIndex: mhRow + 1, endRowIndex: mhRow + 7,
              startColumnIndex: 0, endColumnIndex: 7,
            }],
            booleanRule: {
              condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: fml }] },
              format: { backgroundColor: color },
            },
          },
          index: 0,
        },
      });
    }
  }

  // ─── Execute ──────────────────────────────────────────────────────────────
  console.log(`\nBatch 1 (${batch1.length} requests: MEL format + CF deletions)...`);
  for (let i = 0; i < batch1.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: batch1.slice(i, i + CHUNK) },
    });
    console.log(`  Chunk ${Math.floor(i / CHUNK) + 1}/${Math.ceil(batch1.length / CHUNK)} ✓`);
  }
  console.log('✅  Batch 1 done.');

  console.log(`\nBatch 2 (${batch2.length} requests: legend + CF rules)...`);
  for (let i = 0; i < batch2.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: batch2.slice(i, i + CHUNK) },
    });
    console.log(`  Chunk ${Math.floor(i / CHUNK) + 1}/${Math.ceil(batch2.length / CHUNK)} ✓`);
  }
  console.log('✅  Batch 2 done.');

  console.log(`\nValue updates (${valueUpdates.length} ranges)...`);
  for (let i = 0; i < valueUpdates.length; i += VCHUNK) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: valueUpdates.slice(i, i + VCHUNK) },
    });
  }
  console.log('✅  Values done.');

  console.log('\n✅  All done:');
  console.log('    MEL col X       — formatted as dates (M/d/yyyy).');
  console.log(`    WS row 5        — category color legend (${cats.length} categories in B5–N5).`);
  console.log('    MC row 3 cols H+ — category color legend.');
  console.log(`    WS/MC CF rules  — ${cats.length} category rules on WS, ${cats.length * monthHeaderRows.length} on MC.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
