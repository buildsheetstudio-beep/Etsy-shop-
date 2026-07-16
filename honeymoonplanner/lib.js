'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// ── Auth ──────────────────────────────────────────────────────────────────────
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

function getAuth() {
  const creds = JSON.parse(fs.readFileSync(SECRET_PATH));
  const { client_id, client_secret, redirect_uris } = creds.installed || creds.web;
  const oAuth2 = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2;
}

function sheets() { return google.sheets({ version: 'v4', auth: getAuth() }); }

// ── Color palette ─────────────────────────────────────────────────────────────
const C = {
  dustyPeach:   '#E0A47F',
  slateTeal:    '#3E5C58',
  warmIvory:    '#FBF6EF',
  mutedSage:    '#7A9B7A',
  amber:        '#D9A548',
  rustRed:      '#B3453A',
  altRow:       '#F2E9DF',
  warmCharcoal: '#3A2E28',
  inputBg:      '#FFFDF8',
  formulaBg:    '#F6E4D6',
  white:        '#FFFFFF',
  lightSage:    '#D4E8D4',
  lightAmber:   '#FFF0D0',
  lightRust:    '#F9D6D4',
  lightTeal:    '#D0E4E2',
  mediumGray:   '#9E9E9E',
  deepTeal:     '#2C4440',
  darkText:     '#3A2E28',
};

function hex(h) {
  const r = parseInt(h.slice(1, 3), 16) / 255;
  const g = parseInt(h.slice(3, 5), 16) / 255;
  const b = parseInt(h.slice(5, 7), 16) / 255;
  return { red: r, green: g, blue: b };
}

// ── gridRange helper (0-indexed, exclusive end) ───────────────────────────────
function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

// ── API wrappers ──────────────────────────────────────────────────────────────
const CHUNK  = 400;
const VCHUNK = 200;

async function batchUpdate(spreadsheetId, requests, label) {
  const s = sheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await s.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (requests.length > CHUNK) console.log(`  ${label}: chunk ${i / CHUNK + 1}/${Math.ceil(requests.length / CHUNK)}`);
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const s = sheets();
  for (let i = 0; i < data.length; i += VCHUNK) {
    const chunk = data.slice(i, i + VCHUNK);
    await s.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (data.length > VCHUNK) console.log(`  ${label}: chunk ${i / VCHUNK + 1}/${Math.ceil(data.length / VCHUNK)}`);
  }
}

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
