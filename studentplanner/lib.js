'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const SECRET = path.join(__dirname, 'client_secret.json');
const TOKEN  = path.join(__dirname, 'tokens.json');

const { client_id, client_secret, redirect_uris } = JSON.parse(fs.readFileSync(SECRET)).installed;
const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
auth.setCredentials(JSON.parse(fs.readFileSync(TOKEN)));

const sheets = google.sheets({ version: 'v4', auth });

// ── Palette ──────────────────────────────────────────────────────────────────
function hex(h) {
  const v = h.replace('#','');
  return { red: parseInt(v.slice(0,2),16)/255, green: parseInt(v.slice(2,4),16)/255, blue: parseInt(v.slice(4,6),16)/255 };
}

const COLOR = {
  skyBlue:      '#3B82F6',
  lightBlue:    '#60A5FA',
  mint:         '#34D399',
  lavender:     '#A78BFA',
  peach:        '#FB923C',
  gold:         '#FBBF24',
  paleSkyBlue:  '#DBEAFE',
  paleMint:     '#D1FAE5',
  paleLavender: '#EDE9FE',
  palePeach:    '#FFEDD5',
  iceBlue:      '#F8FAFF',
  subheaderBg:  '#EFF6FF',
  alertRed:     '#EF4444',
  nearBlack:    '#1E293B',
  paleGold:     '#FEF3C7',
  paleRed:      '#FEE2E2',
  white:        '#FFFFFF',
  softGrey:     '#F1F5F9',
  midGrey:      '#94A3B8',
  border:       '#3B82F6',
  altRow:       '#F8FAFF',
  inputBg:      '#FFFBF0',
  formulaBg:    '#F0F4FA',
  darkBlue:     '#1D4ED8',
  emerald:      '#059669',
  violet:       '#7C3AED',
  orange:       '#EA580C',
  amber:        '#D97706',
};

// ── Helpers ───────────────────────────────────────────────────────────────────
function colLetter(n) { // 0-indexed
  let s = '';
  n++;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

function vr(sheetId, r0, r1, c0, c1) { return gridRange(sheetId, r0, r1, c0, c1); }

function colFormulaRange(sheetId, r0, r1, col) {
  return gridRange(sheetId, r0, r1, col, col + 1);
}

// ── Batch update (chunks of 400) ──────────────────────────────────────────────
const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label = '') {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (label) console.log(`  [${label}] chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)} (${chunk.length} reqs)`);
  }
}

// ── Values batch update (chunks of 200) ───────────────────────────────────────
const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label = '') {
  for (let i = 0; i < data.length; i += VCHUNK) {
    const chunk = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (label) console.log(`  [${label}] values chunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)} (${chunk.length} ranges)`);
  }
}

function clearConditionalFormats(sheetId) {
  return { deleteConditionalFormatRule: { sheetId, index: 0 } };
}

module.exports = { sheets, hex, COLOR, colLetter, gridRange, vr, colFormulaRange, batchUpdate, valuesBatchUpdate, clearConditionalFormats };
