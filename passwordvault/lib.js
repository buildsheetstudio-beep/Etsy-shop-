'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const CHUNK  = 400;
const VCHUNK = 200;

const C = {
  deepIndigo:     '#2D2B8F',
  midnightIndigo: '#1A1560',
  clearBlue:      '#4A90D9',
  paleIndigo:     '#F0F2FF',
  veryPaleIndigo: '#E8ECFF',
  silkyWhite:     '#FAFBFF',
  deepGreen:      '#1A6B3A',
  darkAmber:      '#8B6000',
  deepRed:        '#8B0000',
  bodyText:       '#1C1C2E',
  mutedIndigo:    '#666699',
  white:          '#FFFFFF',
  border:         '#C5C8E8',
  mediumIndigo:   '#4A4A9F',
  softSilver:     '#E8EAF6',
  lightBorder:    '#D8DAF0',
  doneBg:         '#D4EDDA',
  doneFg:         '#1A6B3A',
  warnBg:         '#FFF3CD',
  warnFg:         '#8B6000',
  errorBg:        '#FFE8E8',
  errorFg:        '#8B0000',
  noticeBg:       '#FFF8E7',
  noticeBorder:   '#F5A623',
  noticeFg:       '#7A4400',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname,'client_secret.json')));
  const { client_id, client_secret, redirect_uris } = secret.installed;
  const oAuth2 = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const tokens = JSON.parse(fs.readFileSync(path.join(__dirname,'tokens.json')));
  oAuth2.setCredentials(tokens);
  return oAuth2;
}

async function batchUpdate(spreadsheetId, requests, label) {
  const auth   = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const auth   = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < data.length; i += VCHUNK) {
    const chunk = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (data.length > VCHUNK) console.log(`  ${label} chunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)}`);
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

// Column index helper: A=0, B=1, …
function col(letter) { return letter.charCodeAt(0) - 65; }

module.exports = { C, hex, getAuth, batchUpdate, valuesBatchUpdate, gridRange, col, google };
