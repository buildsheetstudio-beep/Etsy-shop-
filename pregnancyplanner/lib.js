'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const { client_id, client_secret, redirect_uris } = secret.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
const tokens = JSON.parse(fs.readFileSync(__dirname + '/tokens.json'));
oAuth2Client.setCredentials(tokens);
const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

const C = {
  deepLavender:   '#6B5B8B',
  softLavender:   '#8B7DA8',
  taupeGold:      '#B8A06A',
  lavenderMist:   '#F5F2F9',
  ivory:          '#FFFDF7',
  paleLavender:   '#F0EDF5',
  warmCream:      '#FAF7F2',
  sage:           '#5A7A6A',
  dustyRose:      '#9B5A6A',
  deepPlum:       '#2A2235',
  warmGrey:       '#8C8C8C',
  white:          '#FFFFFF',
  lavenderBorder: '#DDD5E8',
  whisperLav:     '#EDE8F5',
  errorBg:        '#F8D7DA',
  doneBg:         '#D4EDDA',
  doneFg:         '#5A7A6A',
  warnBg:         '#FFF3CD',
  warnFg:         '#B8A06A',
  moodLowBg:      '#F5E8ED',
};

function hex(h) {
  const v = h.replace('#', '');
  return {
    red:   parseInt(v.slice(0,2),16)/255,
    green: parseInt(v.slice(2,4),16)/255,
    blue:  parseInt(v.slice(4,6),16)/255,
  };
}

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

function colLetter(n) {
  let s = '';
  for (let x = n + 1; x > 0; x = Math.floor((x - 1) / 26))
    s = String.fromCharCode(((x - 1) % 26) + 65) + s;
  return s;
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (data.length > VCHUNK) console.log(`  ${label} vchunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)}`);
  }
}

module.exports = { sheets, batchUpdate, valuesBatchUpdate, hex, C, gridRange, colLetter };
