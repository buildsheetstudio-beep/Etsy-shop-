'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const C = {
  deepPlum:   '#5B2A4D',
  roseGold:   '#C08D6B',
  mutedGold:  '#C9A667',
  sageGreen:  '#8FA888',
  warmSand:   '#EAD9C4',
  mutedRust:  '#B5573F',
  ivoryCream: '#FAF6F0',
  altRow:     '#F3EBE6',
  white:      '#FFFFFF',
  bodyText:   '#3A1A2E',
  border:     '#D4BFC8',
  lightGold:  '#FFF8E1',
  lightGreen: '#E8F2E8',
  lightRust:  '#FAE4E0',
  paleGold:   '#FFFDF0',
  palePlum:   '#F0E8EE',
  lightGray:  '#F5F5F5',
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
