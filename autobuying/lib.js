'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TOKEN_PATH = path.join(__dirname, 'tokens.json');
const SECRET_PATH = path.join(__dirname, 'client_secret.json');

function getAuth() {
  const credentials = JSON.parse(fs.readFileSync(SECRET_PATH));
  const { client_id, client_secret, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
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
  primary:   '#244C5A',
  secondary: '#C77B4B',
  bg:        '#F3F1EC',
  panel:     '#FFFFFF',
  success:   '#8DA98F',
  warning:   '#D2AD63',
  attention: '#C98B4B',
  rust:      '#B86D66',
  mutedBlue: '#ABC2CF',
  slate:     '#7B8790',
  mainText:  '#2B3235',
  secText:   '#687176',
  border:    '#C7C3BA',
  input:     '#FEFAE8',
  formula:   '#EDF2F5',
  altRow:    '#F0EDE7',
  white:     '#FFFFFF',
};

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
