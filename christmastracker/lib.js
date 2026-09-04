'use strict';
const { google } = require('googleapis');
const fs = require('fs');

const C = {
  deepCranberry: '#8C1F2B',
  warmGold:      '#C9A227',
  snowWhite:     '#F7F8F7',
  mutedSage:     '#7A9B7A',
  amber:         '#D9A23C',
  trueRed:       '#B23A3A',
  altRow:        '#F0EAE4',
  darkText:      '#2E1F22',
  inputBg:       '#FFFFFE',
  formulaBg:     '#F2E3E5',
  white:         '#FFFFFF',
  lightSage:     '#D4E8D4',
  lightRed:      '#F9D6D6',
  lightAmber:    '#FFF0D0',
  lightGold:     '#FFF3CC',
  deepGreen:     '#2D5016',
  lightGray:     '#EEEEEE',
  mediumGray:    '#9E9E9E',
};

function hex(h) {
  const r = parseInt(h.slice(1, 3), 16) / 255;
  const g = parseInt(h.slice(3, 5), 16) / 255;
  const b = parseInt(h.slice(5, 7), 16) / 255;
  return { red: r, green: g, blue: b };
}

function getAuth() {
  const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = secret.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const tokens = JSON.parse(fs.readFileSync(__dirname + '/tokens.json'));
  oAuth2Client.setCredentials(tokens);
  return oAuth2Client;
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label = '') {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (label) console.log(`  ${label} chunk ${Math.floor(i / CHUNK) + 1}/${Math.ceil(requests.length / CHUNK)}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label = '') {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (label) console.log(`  ${label} chunk ${Math.floor(i / VCHUNK) + 1}/${Math.ceil(data.length / VCHUNK)}`);
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { C, hex, getAuth, batchUpdate, valuesBatchUpdate, gridRange };
