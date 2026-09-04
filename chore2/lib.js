'use strict';
const { google } = require('googleapis');
const fs = require('fs');

const CHUNK = 400;
const VCHUNK = 200;

function getAuth() {
  const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = secret.installed;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const tokens = JSON.parse(fs.readFileSync(__dirname + '/tokens.json'));
  auth.setCredentials(tokens);
  return auth;
}

async function batchUpdate(spreadsheetId, requests, label) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < data.length; i += VCHUNK) {
    const chunk = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk }
    });
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

const C = {
  chrome:    '#EDEDEA',  // neutral header/label bg
  charcoal:  '#3A3A38',  // dark text
  white:     '#FFFFFF',
  sage:      '#7A9B7A',  // success in reward tab
  rust:      '#C0392B',  // negative balance
  // Per-slot pastel colors (10 slots)
  slot: ['#F5D9C0','#F5D0DC','#D0E4F5','#E0D4F0','#D4F0E0','#F5EFD0','#F5DCD4','#D4EDF0','#E8D4F0','#DCE8D4'],
  // Lighter alternating tints
  slotAlt: ['#FAF0E6','#FAE8F0','#EAF3FA','#F0EAF8','#EAF8F0','#FAF7E8','#FAF0EC','#EAF7FA','#F4EAF8','#EEF4EA'],
};

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
