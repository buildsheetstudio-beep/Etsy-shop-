'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

function getAuth() {
  const creds = JSON.parse(fs.readFileSync(SECRET_PATH));
  const { client_id, client_secret, redirect_uris } = creds.installed || creds.web;
  const oAuth2 = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2;
}

function sheets() { return google.sheets({ version: 'v4', auth: getAuth() }); }

// ── Color palette ─────────────────────────────────────────────────────────────
const C = {
  cobalt:      '#2454A6',   // primary header bg
  limeGreen:   '#7CB342',   // secondary accent
  offWhite:    '#F7F9FB',   // background
  sage:        '#6B9B6E',   // success / completed
  amber:       '#DFA23C',   // warning / planned not done
  brick:       '#B3382E',   // danger / missed / broken streak
  altRow:      '#EAF0EF',   // alternating row
  navyDark:    '#233047',   // dark text
  inputBg:     '#FFFFFC',   // input cell bg
  formulaBg:   '#E3EAF5',   // formula/read-only cell bg
  white:       '#FFFFFF',
  lightSage:   '#D4EACF',   // light sage for CF
  lightAmber:  '#FBF0D8',   // light amber for CF
  lightBrick:  '#F9DCDA',   // light brick for CF
  gold:        '#F0C040',   // PR current best highlight
  lightGold:   '#FFF8D6',   // light gold for CF
};

// Muscle group / workout type color coding — identical across ALL tabs
const MG_COLORS = {
  'Chest':      '#FFE0D0',
  'Back':       '#CCE8F7',
  'Shoulders':  '#E8D5F7',
  'Arms':       '#FDE8CC',
  'Legs':       '#FFF5CC',
  'Glutes':     '#FFD6E8',
  'Core':       '#D4F0E4',
  'Cardio':     '#FFD6D6',
  'Rest':       '#E8E8E8',
};

// Zone colors — cool to warm gradient (not good/bad)
const ZONE_COLORS = {
  'Zone 1': '#D6EAF8',  // light blue
  'Zone 2': '#D1F2EB',  // light teal/green
  'Zone 3': '#FEF9E7',  // light yellow
  'Zone 4': '#FDEBD0',  // light orange
  'Zone 5': '#FADBD8',  // light red
};

function hex(h) {
  if (!h || h.length < 7) return { red: 0, green: 0, blue: 0 };
  const r = parseInt(h.slice(1, 3), 16) / 255;
  const g = parseInt(h.slice(3, 5), 16) / 255;
  const b = parseInt(h.slice(5, 7), 16) / 255;
  return { red: r, green: g, blue: b };
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

const CHUNK  = 400;
const VCHUNK = 200;

async function batchUpdate(spreadsheetId, requests, label) {
  const s = sheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await s.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (requests.length > CHUNK) console.log(`  ${label}: chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const s = sheets();
  for (let i = 0; i < data.length; i += VCHUNK) {
    const chunk = data.slice(i, i + VCHUNK);
    await s.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (data.length > VCHUNK) console.log(`  ${label}: chunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)}`);
  }
}

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MG_COLORS, ZONE_COLORS };
