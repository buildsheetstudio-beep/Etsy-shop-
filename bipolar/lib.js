'use strict';
const fs = require('fs');
const { google } = require('googleapis');

const SECRET = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const creds = SECRET.installed || SECRET.web;

function getAuth() {
  const oAuth2 = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
  return oAuth2;
}

const sheets = () => google.sheets({ version: 'v4', auth: getAuth() });

const CHUNK = 400;
const VCHUNK = 200;

async function batchUpdate(spreadsheetId, requests, tag) {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets().spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 400));
  }
}

async function valuesBatchUpdate(spreadsheetId, data, tag) {
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets().spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (i + VCHUNK < data.length) await new Promise(r => setTimeout(r, 400));
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function hex(h) {
  const c = h.replace('#', '');
  const r = parseInt(c.substring(0,2),16)/255;
  const g = parseInt(c.substring(2,4),16)/255;
  const b = parseInt(c.substring(4,6),16)/255;
  return { red: r, green: g, blue: b };
}

const C = {
  periwinkle:  '#6B7FA3',   // primary header bg
  taupe:       '#A68B6F',   // secondary accent
  fog:         '#F5F3F0',   // background
  heatMin:     '#DCE4F0',   // heatmap minimal
  heatMid:     '#8FA0C2',   // heatmap moderate
  heatMax:     '#3A4566',   // heatmap severe
  sage:        '#7A9B7A',   // medication taken
  missedGray:  '#C9C5BE',   // medication missed (not red)
  altRow:      '#EDEAE5',   // alternating row
  white:       '#FFFFFF',
  darkText:    '#2E3244',
  inputBg:     '#FFFFFF',   // input cells white bg
  formulaBg:   '#E7EAF1',   // formula/read-only pale periwinkle
  amber:       '#D9A548',   // gentle therapy nudge
  lightAmber:  '#FBF0D8',
};

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
