'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

function getAuth() {
  const { client_secret, client_id, redirect_uris } =
    JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

const sheets = google.sheets({ version: 'v4', auth: getAuth() });

// Modern Teal & Terracotta palette
const C = {
  primary:   '#355F62',  // Deep Petrol Teal
  secondary: '#B9684D',  // Burnished Terracotta
  bg:        '#F4F2EE',  // Soft Stone
  panel:     '#FFFFFF',
  teal:      '#D8E6E4',  // Soft Teal Tint
  terra:     '#EAD1C7',  // Soft Terracotta Tint
  sand:      '#D8C59D',  // Accent Sand
  success:   '#8FA88F',
  warning:   '#D2B16A',
  attention: '#B8756F',
  info:      '#A9C3CE',
  text:      '#2D3435',
  secText:   '#6D7475',
  border:    '#C8C6C0',
  white:     '#FFFFFF',
  gray:      '#B0AFA9',
  input:     '#FFFDF7',
  formula:   '#EBF2F5',
  altRow:    '#F0EDE8',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

async function batchUpdate(id, requests, label) {
  const CHUNK = 400;
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: chunk } });
    if (label) console.log(`  [${label}] batch ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 300));
  }
}

async function valuesBatchUpdate(id, data, label) {
  const CHUNK = 200;
  for (let i = 0; i < data.length; i += CHUNK) {
    const chunk = data.slice(i, i + CHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: id,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (label) console.log(`  [${label}] values batch ${Math.floor(i/CHUNK)+1}/${Math.ceil(data.length/CHUNK)}`);
    if (i + CHUNK < data.length) await new Promise(r => setTimeout(r, 300));
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { getAuth, sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange };
