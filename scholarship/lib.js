'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const CHUNK  = 400;
const VCHUNK = 200;

const C = {
  deepBlue:   '#1A3A8B',
  royalBlue:  '#1A56C4',
  gold:       '#C9A800',
  paleGold:   '#FFFDF0',
  inputBg:    '#FFFFFF',
  formulaBg:  '#EEF3FF',
  green:      '#1A6B3A',
  red:        '#8B1A1A',
  purple:     '#6A5A8B',
  bodyText:   '#0D1B3E',
  slateBlue:  '#4A5A8B',
  white:      '#FFFFFF',
  border:     '#C5D0E8',
  paleBlue:   '#F0F4FF',
  lightGold:  '#FFF8D4',
  lightGreen: '#D4EDDA',
  lightRed:   '#F8D7DA',
  lightPurple:'#EDE8F8',
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

module.exports = { C, hex, getAuth, batchUpdate, valuesBatchUpdate, gridRange, google };
