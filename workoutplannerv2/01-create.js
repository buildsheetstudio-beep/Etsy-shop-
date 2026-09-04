'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SECRET = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
const { client_id, client_secret, redirect_uris } = SECRET.installed || SECRET.web;
const oAuth2 = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
oAuth2.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));

const TABS = [
  { title: '📋 Reference Data',                     color: '#9E9E9E' },
  { title: '📊 Dashboard',                           color: '#2454A6' },
  { title: '💪 Workout Log',                         color: '#7CB342' },
  { title: '📅 Monthly Calendar View',               color: '#3F8CB5' },
  { title: '📝 Weekly Detail Log',                   color: '#5C6BC0' },
  { title: '⚡ Superset & Circuit Builder',          color: '#F57C00' },
  { title: '🏆 Personal Records (PR) Tracker',      color: '#C9A227' },
  { title: '❤️ Cardio & Heart Rate Zone Tracker',   color: '#B3382E' },
  { title: '📏 Body Measurements & Progress Photos', color: '#7B5EA7' },
  { title: '🔥 Streak & Consistency Tracker',       color: '#E65100' },
];

(async () => {
  const s = google.sheets({ version: 'v4', auth: oAuth2 });
  const res = await s.spreadsheets.create({
    requestBody: {
      properties: { title: '🏋️ Ultimate Workout Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColorStyle: {
            rgbColor: {
              red:   parseInt(t.color.slice(1,3), 16) / 255,
              green: parseInt(t.color.slice(3,5), 16) / 255,
              blue:  parseInt(t.color.slice(5,7), 16) / 255,
            },
          },
        },
      })),
    },
  });
  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });
  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', id);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
