'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
];

function getAuth() {
  const creds = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const { client_id, client_secret, redirect_uris } = creds.installed || creds.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const tokens = JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json')));
  oAuth2Client.setCredentials(tokens);
  return oAuth2Client;
}

const auth = getAuth();
const sheets = google.sheets({ version: 'v4', auth });

const CHUNK = 400;
const VCHUNK = 200;

async function batchUpdate(spreadsheetId, requests, label) {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (requests.length > CHUNK) process.stdout.write(`  ${label} chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}\n`);
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (data.length > VCHUNK) process.stdout.write(`  ${label} vchunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)}\n`);
  }
}

// ── Color palette ─────────────────────────────────────────────────────────────
const C = {
  deepCharcoal:   '#2E3A35',
  mutedSage:      '#3D5A4A',
  warmTaupe:      '#8B7355',
  warmStone:      '#F5F3EE',
  ivory:          '#FFFDF7',
  paleSage:       '#EDF0EB',
  warmCream:      '#FAF8F4',
  deepRust:       '#7A3535',
  charcoal:       '#1C2420',
  warmGrey:       '#888888',
  white:          '#FFFFFF',
  stoneBorder:    '#D5CFC4',
  lightStoneBg:   '#E8E4DC',
  doneBg:         '#D4E4DA',
  doneFg:         '#3D5A4A',
  warnBg:         '#FFF3CD',
  warnFg:         '#8B7355',
  errorBg:        '#F0E0E0',
  errorFg:        '#7A3535',
  sageAccent:     '#4A6B58',
  softSage:       '#8FAF9A',
  taupeLight:     '#D6CBBF',
  disclaimerBg:   '#FEF9F0',
};

function hex(h) {
  const c = h.replace('#', '');
  const r = parseInt(c.slice(0,2),16)/255;
  const g = parseInt(c.slice(2,4),16)/255;
  const b = parseInt(c.slice(4,6),16)/255;
  return { red: r, green: g, blue: b };
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

// column index → 0-based
function C_(col) { return col; }

module.exports = { sheets, batchUpdate, valuesBatchUpdate, hex, C, gridRange, C_ };
