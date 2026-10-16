'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const C = {
  primary:    '#5C536E',  // Deep Academic Aubergine
  secondary:  '#7FA69A',  // Muted Eucalyptus
  accent:     '#C39A74',  // Warm Copper Sand
  bg:         '#F7F4F0',  // Soft Chalk
  white:      '#FFFFFF',
  aubergTint: '#E3DFE8',  // Light Aubergine Tint
  eucalTint:  '#DCE8E3',  // Light Eucalyptus Tint
  copperTint: '#EBDDD0',  // Light Copper Tint
  text:       '#303236',
  secText:    '#747579',
  border:     '#CAC6C0',
  success:    '#8FA990',
  warning:    '#D0AE69',
  attention:  '#B77A72',
  info:       '#A9BDD1',
  input:      '#FEFCE8',  // Pale warm yellow — input cells
  formula:    '#EAF0F4',  // Pale blue-gray — formula cells
  altRow:     '#F3F1EE',
  hdrDark:    '#5C536E',  // = primary
  hdrLight:   '#E3DFE8',  // = aubergTint
  amber:      '#D0AE69',
  gray:       '#DDDAD6',
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

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  const sheets = await getSheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  const sheets = await getSheets();
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

module.exports = { C, hex, getAuth, getSheets, batchUpdate, valuesBatchUpdate, gridRange };
