'use strict';
const { google } = require('googleapis');
const fs = require('fs');

// ── OAuth ─────────────────────────────────────────────────────────────────────
const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const { client_id, client_secret, redirect_uris } = secret.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
const tokens = JSON.parse(fs.readFileSync(__dirname + '/tokens.json'));
oAuth2Client.setCredentials(tokens);
const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

// ── Colour Palette ────────────────────────────────────────────────────────────
const C = {
  deepForest:    '#2C4A35',  // Primary headers / title banners
  mossGreen:     '#3D6B4F',  // Secondary headers
  forestGreen:   '#2D6A4F',  // Positive / done
  antiqueBrass:  '#B8922A',  // Accent / KPI values / input borders
  warmCream:     '#FBF5E6',  // Alternating rows / subtle fills
  ivory:         '#FFFDF7',  // Input cell bg
  amber:         '#B5830A',  // Warning / in progress
  deepRust:      '#8B2500',  // Negative / overdue / cancelled
  nearBlack:     '#1C1C1E',  // Body text
  white:         '#FFFFFF',
  warmSand:      '#D9C9A3',  // Divider / border
  linen:         '#F2EDE0',  // Light section bg
  readOnlyBg:    '#F0F4F0',  // Formula / auto-calculated cells
  // Status helpers
  doneBg:        '#D4EDDA',
  doneFg:        '#2D6A4F',
  warnBg:        '#FFF3CD',
  warnFg:        '#B5830A',
  errorBg:       '#F8D7DA',
  errorFg:       '#8B2500',
  // Extra
  slateBlue:     '#3A5C7A',
  warmBrown:     '#6B5C3A',
  olive:         '#4A5C3A',
  sageDark:      '#5C7A4A',
  tealGreen:     '#4A6B5C',
  darkBrass:     '#7A6B3A',
};

function hex(h) {
  const c = h.replace('#','');
  return {
    red:   parseInt(c.slice(0,2),16)/255,
    green: parseInt(c.slice(2,4),16)/255,
    blue:  parseInt(c.slice(4,6),16)/255,
  };
}

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

function colLetter(n) {
  let s = '';
  n++;
  while (n > 0) { const m = (n-1) % 26; s = String.fromCharCode(65+m) + s; n = (n-m-1)/26; }
  return s;
}

// ── API helpers ───────────────────────────────────────────────────────────────
const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const batch = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: batch } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  for (let i = 0; i < data.length; i += VCHUNK) {
    const batch = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: batch },
    });
    if (data.length > VCHUNK) console.log(`  ${label} values chunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)}`);
  }
}

module.exports = { sheets, batchUpdate, valuesBatchUpdate, hex, C, gridRange, colLetter };
