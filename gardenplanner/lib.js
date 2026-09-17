'use strict';
const { google } = require('googleapis');
const fs = require('fs');

// ── OAuth ─────────────────────────────────────────────────────────────────────
const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const { client_id, client_secret, redirect_uris } = secret.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
const tokens = JSON.parse(fs.readFileSync(__dirname + '/tokens.json'));
oAuth2Client.setCredentials(tokens);
const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

// ── Colour Palette ────────────────────────────────────────────────────────────
const C = {
  // Primary
  terracotta:     '#C4622D',  // deep warm orange-red
  terracottaLight:'#E8956D',  // lighter terracotta
  paleTerracotta: '#F5D5C3',  // very light terracotta/peach
  cream:          '#FAF3E0',  // warm off-white
  cream2:         '#FFF8EE',  // lighter cream
  parchment:      '#F0E6CE',  // warm parchment
  // Greens
  olive:          '#6B7C45',  // muted olive green
  oliveLight:     '#9AAD6A',  // lighter olive
  paleSage:       '#D4DDB8',  // very light sage
  sage:           '#8FA67A',  // medium sage
  // Earthy accents
  soil:           '#5C3D1E',  // dark brown (soil)
  soilLight:      '#8B6347',  // medium brown
  wheat:          '#D4A84B',  // warm gold/wheat
  wheatLight:     '#F0D080',  // light gold
  paleWheat:      '#FBF0C8',  // very light gold
  // Neutrals
  charcoal:       '#3D3028',  // dark warm gray
  warmGray:       '#8C7B6B',  // warm gray
  lightGray:      '#E8E0D5',  // light warm gray
  white:          '#FFFFFF',
  // Status colors
  growingGreen:   '#4A7C59',  // dark green for growing/active
  growingBg:      '#D6EADE',  // light green bg
  harvestGold:    '#B8860B',  // gold for harvest ready
  harvestBg:      '#FFF3CD',  // light gold bg
  notPlantedGray: '#7A7A7A',  // gray for not planted
  notPlantedBg:   '#F0F0F0',  // light gray bg
  warningRed:     '#C0392B',  // red for pests/issues
  warningBg:      '#FADBD8',  // light red bg
  completeBg:     '#D6EADE',
  complete:       '#1E7A3C',
  inProgress:     '#2E6DA4',
  inProgressBg:   '#D6E8F7',
};

function hex(h) {
  const c = h.replace('#','');
  return {
    red:   parseInt(c.slice(0,2),16)/255,
    green: parseInt(c.slice(2,4),16)/255,
    blue:  parseInt(c.slice(4,6),16)/255,
  };
}

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

function colLetter(n) {
  let s = '';
  n++;
  while (n > 0) { const m = (n-1) % 26; s = String.fromCharCode(65+m) + s; n = (n-m-1)/26; }
  return s;
}

// ── API helpers ───────────────────────────────────────────────────────────────
const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const batch = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: batch } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  for (let i = 0; i < data.length; i += VCHUNK) {
    const batch = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: batch },
    });
    if (data.length > VCHUNK) console.log(`  ${label} values chunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)}`);
  }
}

module.exports = { sheets, batchUpdate, valuesBatchUpdate, hex, C, gridRange, colLetter };
