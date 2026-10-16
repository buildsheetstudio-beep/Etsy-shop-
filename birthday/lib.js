'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = path.join(__dirname, 'tokens.json');
const SECRET_PATH = path.join(__dirname, 'client_secret.json');

function getAuth() {
  const credentials = JSON.parse(fs.readFileSync(SECRET_PATH));
  const { client_id, client_secret, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
  oAuth2Client.setCredentials(token);
  return oAuth2Client;
}

function getSheets() { return google.sheets({ version: 'v4', auth: getAuth() }); }

const CHUNK = 400;
const VCHUNK = 200;

async function batchUpdate(spreadsheetId, requests, label) {
  const sheets = getSheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 300));
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const sheets = getSheets();
  for (let i = 0; i < data.length; i += VCHUNK) {
    const chunk = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk }
    });
    if (i + VCHUNK < data.length) await new Promise(r => setTimeout(r, 300));
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red:r, green:g, blue:b };
}

const C = {
  primary:   '#28566B',   // Deep Peacock Blue
  secondary: '#E79A78',   // Apricot Coral
  bg:        '#FBF7ED',   // Soft Vanilla
  panel:     '#FFFFFF',
  success:   '#91B49A',
  warning:   '#D8B66C',
  attention: '#C27772',
  mutedBlue: '#AFC9D5',
  lavender:  '#C9BEDA',
  mainText:  '#2D3134',
  secText:   '#6D7275',
  border:    '#CEC8BE',
  white:     '#FFFFFF',
  input:     '#FEFAE8',
  formula:   '#EBF2F5',
  altRow:    '#F5F1E8',
};

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
