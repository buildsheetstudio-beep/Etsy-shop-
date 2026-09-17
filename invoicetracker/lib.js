'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TOKEN_PATH  = path.join(__dirname, 'tokens.json');
const SECRET_PATH = path.join(__dirname, 'client_secret.json');

function getAuth() {
  const { client_secret, client_id, redirect_uris } =
    JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

function getSheets() { return google.sheets({ version: 'v4', auth: getAuth() }); }

const CHUNK  = 400;
const VCHUNK = 200;

async function batchUpdate(spreadsheetId, requests, label) {
  const sheets = getSheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 300));
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const sheets = getSheets();
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice }
    });
    if (i + VCHUNK < data.length) await new Promise(r => setTimeout(r, 300));
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function hex(h) {
  const c = h.replace('#', '');
  const r = parseInt(c.slice(0, 2), 16) / 255;
  const g = parseInt(c.slice(2, 4), 16) / 255;
  const b = parseInt(c.slice(4, 6), 16) / 255;
  return { red: r, green: g, blue: b };
}

const C = {
  primary:   '#5A3A52',  // Deep Mulberry
  secondary: '#78998B',  // Muted Eucalyptus
  bg:        '#F7F4EF',  // Soft Alabaster
  panel:     '#FFFFFF',
  success:   '#8DAA91',  // soft sage
  warning:   '#D2AD66',  // soft amber
  attention: '#B8736F',  // muted rust
  mutedBlue: '#AFC2D0',
  neutral:   '#A9A6A2',  // neutral gray
  mainText:  '#2F2C30',
  secText:   '#6E686D',
  border:    '#CBC5BF',
  input:     '#FEFCE8',  // pale yellow for user inputs
  formula:   '#EEF2F6',  // pale blue-gray for formula cells
  altRow:    '#F0EDE8',
  white:     '#FFFFFF',
};

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
