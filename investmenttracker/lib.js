'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

function getAuth() {
  const { client_secret, client_id, redirect_uris } =
    JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

const sheets = google.sheets({ version: 'v4', auth: getAuth() });

// Deep Blue Spruce / Muted Orchid Clay / Cool Porcelain theme
const C = {
  text:        '#2D3436',
  secText:     '#6A7375',
  bg:          '#F4F5F2',   // Cool Porcelain
  panel:       '#FFFFFF',
  border:      '#C8CBC6',
  altRow:      '#E8ECEA',
  input:       '#FDFCF8',   // pale warm yellow tint for inputs
  formula:     '#EDF1F3',   // pale blue-gray for formula cells
  white:       '#FFFFFF',

  primary:     '#2F5860',   // Deep Blue Spruce
  primaryText: '#FFFFFF',
  secondary:   '#A56F82',   // Muted Orchid Clay
  secondaryDk: '#8A5B6C',

  highlight:   '#DDE8E7',   // Pale Glacier
  success:     '#8EAA92',
  warning:     '#D2B16A',
  attention:   '#B97972',
  info:        '#AFC4D4',

  hdrA:        '#2F5860',   // primary dark
  hdrB:        '#3D6E78',   // primary mid
  hdrC:        '#547F87',   // primary light

  // Asset-class colors
  usStocks:    '#496A88',
  intlStocks:  '#7A6680',
  emerging:    '#C98474',
  bonds:       '#8FA98C',
  realestate:  '#C6A15B',
  cash:        '#8FB9B7',
  commodities: '#D0B36A',
  crypto:      '#74758F',
  alts:        '#A59482',
  other:       '#A9A9A6',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

async function batchUpdate(id, requests, label) {
  const CHUNK = 400;
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: chunk } });
    if (label) console.log(`  [${label}] batch ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 300));
  }
}

async function valuesBatchUpdate(id, data, label) {
  const CHUNK = 200;
  for (let i = 0; i < data.length; i += CHUNK) {
    const chunk = data.slice(i, i + CHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: id,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (label) console.log(`  [${label}] values batch ${Math.floor(i/CHUNK)+1}/${Math.ceil(data.length/CHUNK)}`);
    if (i + CHUNK < data.length) await new Promise(r => setTimeout(r, 300));
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { getAuth, sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange };
