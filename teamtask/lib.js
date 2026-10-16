'use strict';
const { google } = require('googleapis');
const fs = require('fs');

const CHUNK  = 400;
const VCHUNK = 200;

function getAuth() {
  const s = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = s.installed;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
  return auth;
}

async function batchUpdate(spreadsheetId, requests, label) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < requests.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId, requestBody: { requests: requests.slice(i, i + CHUNK) }
    });
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < data.length; i += VCHUNK) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: data.slice(i, i + VCHUNK) }
    });
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function hex(h) {
  return { red: parseInt(h.slice(1,3),16)/255, green: parseInt(h.slice(3,5),16)/255, blue: parseInt(h.slice(5,7),16)/255 };
}

// Team Task Tracker palette
const C = {
  primary:    '#533B59',  // Deep Aubergine
  secondary:  '#2F8F89',  // Electric Teal
  bg:         '#F7F5F2',  // Soft Porcelain
  panel:      '#FFFFFF',  // Light Panel
  success:    '#8FAF92',  // sage green
  warning:    '#D7B56D',  // amber
  attention:  '#B97373',  // muted rust
  mutedBlue:  '#AFC6D8',
  lavender:   '#C6B7D5',  // review lavender
  mainText:   '#2C2930',
  secText:    '#6D6870',
  border:     '#C9C4CB',
  white:      '#FFFFFF',
  input:      '#FEFBE8',  // pale warm yellow
  formula:    '#EBF2F5',  // pale blue-gray
  altRow:     '#F0EDE8',
};

const STATUSES = ['Not Started','In Progress','Waiting','Blocked','In Review','Completed','Cancelled'];
const PRIORITIES = ['Low','Medium','High','Urgent'];
const DEPARTMENTS = ['Leadership','Operations','Marketing','Sales','Finance','Customer Service','Human Resources','Product','Creative','Administration','General Team'];
const PROJECTS = ['Website Redesign','Product Launch','Client Onboarding','Monthly Reporting','Social Media Campaign','Team Operations','Event Planning','Process Improvement','Training Program','General Tasks'];
const RECURRENCE = ['None','Daily','Weekdays','Weekly','Every 2 Weeks','Monthly','Quarterly','Yearly'];
const TASK_TYPES = ['Task','Milestone','Meeting Action','Follow-Up','Review','Approval','Administrative','Recurring Responsibility'];
const EFFORTS = ['Quick','Small','Medium','Large','Major'];
const YESNO = ['Yes','No'];
const WEEKDAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
const PROJ_HEALTH = ['On Track','At Risk','Delayed','Complete','Not Started'];
const TEAM_MEMBERS = ['Alex Morgan','Jamie Lee','Taylor Brooks','Jordan Patel','Casey Rivera','Morgan Chen','Riley Davis','Sam Thompson','Avery Walker','Quinn Martinez','Member 11','Member 12','Member 13','Member 14','Member 15'];
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

module.exports = {
  batchUpdate, valuesBatchUpdate, gridRange, hex, C,
  STATUSES, PRIORITIES, DEPARTMENTS, PROJECTS, RECURRENCE, TASK_TYPES, EFFORTS,
  YESNO, WEEKDAYS, PROJ_HEALTH, TEAM_MEMBERS, MONTHS,
};
