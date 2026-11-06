'use strict';
// Google connector. Talks to the Apps Script web app in apps-script/Code.gs.
// Every call that creates or changes something in Google is blocked unless
// GOOGLE_EXTERNAL_ACTIONS_APPROVED=yes, so staging can never create forms by accident.

function configured() {
  return !!(process.env.APPS_SCRIPT_URL && process.env.APPS_SCRIPT_TOKEN);
}

function writesApproved() {
  return process.env.GOOGLE_EXTERNAL_ACTIONS_APPROVED === 'yes';
}

async function call(action, payload) {
  if (!configured()) {
    const e = new Error('Google is not connected. Set APPS_SCRIPT_URL and APPS_SCRIPT_TOKEN (see README, "Connecting Google").');
    e.status = 503;
    throw e;
  }
  let res;
  try {
    res = await fetch(process.env.APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ token: process.env.APPS_SCRIPT_TOKEN, action, payload }),
      redirect: 'follow',
    });
  } catch (err) {
    const e = new Error(`Could not reach the Google Apps Script web app: ${err.message}`);
    e.status = 502;
    throw e;
  }
  const text = await res.text();
  let body;
  try {
    body = JSON.parse(text);
  } catch (_) {
    const e = new Error(`Google Apps Script returned a non-JSON response (HTTP ${res.status}). Check the deployment URL and that it runs as the owner.`);
    e.status = 502;
    throw e;
  }
  if (!body.ok) {
    const e = new Error(`Google Apps Script error: ${body.error || 'unknown error'}`);
    e.status = 502;
    throw e;
  }
  return body.data;
}

function requireWriteApproval(what) {
  if (!configured()) return; // call() reports "not connected" first
  if (!writesApproved()) {
    const e = new Error(`${what} is blocked in staging. A trainer must approve external Google actions by setting GOOGLE_EXTERNAL_ACTIONS_APPROVED=yes.`);
    e.status = 403;
    throw e;
  }
}

module.exports = {
  configured,
  writesApproved,
  ping: () => call('ping', {}),
  createForm: (quiz) => {
    requireWriteApproval('Creating a Google Form');
    return call('createForm', { quiz });
  },
  fetchResponses: (quiz) => call('fetchResponses', { quizId: quiz.id, form: quiz.form }),
  syncSheets: (tables) => {
    requireWriteApproval('Writing to Google Sheets');
    return call('syncSheets', { tables });
  },
};
