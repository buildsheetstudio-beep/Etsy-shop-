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
  const v = h.replace('#', '');
  return { red: parseInt(v.slice(0,2),16)/255, green: parseInt(v.slice(2,4),16)/255, blue: parseInt(v.slice(4,6),16)/255 };
}

const C = {
  dustyLavender:   '#9B8EC4',
  medLavender:     '#B8A9D9',
  blush:           '#F4A7B9',
  skyBlue:         '#93C5FD',
  softGold:        '#F9D77E',
  dustySage:       '#A8C5A0',
  paleLavender:    '#EDE9F5',
  paleBlush:       '#FDE8ED',
  paleSky:         '#DBEAFE',
  paleGold:        '#FEF9C3',
  paleSage:        '#E8F0E9',
  white:           '#FFFFFF',
  lavenderWhite:   '#FAF8FD',
  veryPaleLav:     '#F3F0FA',
  border:          '#D8CEF0',
  complete:        '#A8C5A0',
  completeBg:      '#E8F0E9',
  inProgress:      '#93C5FD',
  inProgressBg:    '#DBEAFE',
  notStarted:      '#D1D5DB',
  notStartedBg:    '#F9FAFB',
  onHold:          '#F9D77E',
  onHoldBg:        '#FEF9C3',
  bodyText:        '#1F1535',
  mutedText:       '#7C6FA0',
  deepPurple:      '#4C3D8F',
  // life area colors (bg)
  healthBg:        '#A8C5A0',
  careerBg:        '#93C5FD',
  financeBg:       '#F9D77E',
  relBg:           '#F4A7B9',
  personalBg:      '#9B8EC4',
  funBg:           '#FED7AA',
  softOrange:      '#F97316',
  // text on dark headers
  paleGoldText:    '#FEF9C3',
  // misc
  lightGray:       '#D1D5DB',
  nearWhite:       '#F9FAFB',
};

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

function colLetter(n) {
  let s = ''; n++;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label = '') {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (label) console.log(`  [${label}] chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)} (${chunk.length} reqs)`);
  }
}

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

module.exports = { sheets, hex, C, colLetter, gridRange, batchUpdate, valuesBatchUpdate };
