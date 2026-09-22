'use strict';
const { google } = require('googleapis');
const fs = require('fs');

const C = {
  cobaltBlue:  '#2454A6',
  limeGreen:   '#7CB342',
  offWhite:    '#FAFAFA',
  steelGray:   '#607D8B',
  coral:       '#E8534A',
  amber:       '#FFC107',
  darkText:    '#212121',
  headerBlue:  '#1A3A7A',
  mutedBlue:   '#5B7EC9',
  teal:        '#00897B',
  purple:      '#7B1FA2',
  white:       '#FFFFFF',
  lightBlue:   '#DBEAFE',
  lightGreen:  '#DCEDC8',
  lightCoral:  '#FFCDD2',
  lightAmber:  '#FFF9C4',
  altRow:      '#F5F5F5',
  inputBg:     '#F8FBFF',
  formulaBg:   '#EEF2FF',
  mediumGray:  '#90A4AE',
  darkGray:    '#455A64',
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
