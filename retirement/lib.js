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

async function batchUpdate(spreadsheetId, requests) {
  const sheets = getSheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 300));
  }
}

async function valuesBatchUpdate(spreadsheetId, data) {
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
  primary:   '#4B536F',   // Deep Slate Indigo
  secondary: '#C58E63',   // Champagne Copper
  bg:        '#F7F3EA',   // Soft Parchment
  panel:     '#FFFFFF',
  success:   '#8FAF93',
  warning:   '#D5B36B',
  attention: '#B87373',
  mutedBlue: '#AFC3D5',
  mutedMauve:'#B8A5B9',
  mainText:  '#2F3138',
  secText:   '#6D7078',
  border:    '#CAC5BC',
  white:     '#FFFFFF',
  input:     '#FEFAE8',
  formula:   '#EBF2F5',
  altRow:    '#F5F1E4',
};

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
