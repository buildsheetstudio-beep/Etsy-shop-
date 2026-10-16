'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const C = {
  deepCharcoal: '#262626',
  burntOrange:  '#C1502E',
  warmCream:    '#FAF6EE',
  mutedSage:    '#6F9172',
  amber:        '#D99A45',
  rustRed:      '#B4432E',
  altRow:       '#F0EBDD',
  darkText:     '#2B2B2B',
  inputBg:      '#FFFDF7',
  formulaBg:    '#E9E7E2',
  white:        '#FFFFFF',
  lightSage:    '#D6EBD8',
  lightAmber:   '#FDF3DB',
  lightRed:     '#FAE0E0',
  lightBlueGray:'#E3E9F0',
  lightGray:    '#F5F5F5',
  gray:         '#8A9B9A',
  border:       '#C1502E',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const { client_id, client_secret, redirect_uris } = secret.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const token = JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json')));
  oAuth2Client.setCredentials(token);
  return oAuth2Client;
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (data.length > VCHUNK) console.log(`  ${label} value chunk ${Math.floor(i/VCHUNK)+1}`);
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { C, hex, getAuth, batchUpdate, valuesBatchUpdate, gridRange };
