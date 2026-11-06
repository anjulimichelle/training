/**
 * Ticket Assessment Hub: Google connector (Apps Script web app).
 *
 * The hub server calls this web app to create Google Forms, read form responses,
 * and write the six database tabs in Google Sheets. It never sends email, never
 * shares files, and never publishes a form. The trainer does those steps by hand
 * after approving the prototype.
 *
 * Setup (once, by a trainer with access to the target Google account):
 *   1. Create a new Apps Script project, paste this file and appsscript.json.
 *   2. Run setupDatabase() once. It creates the database spreadsheet with the six
 *      tabs and stores its ID and a random API token in Script Properties.
 *   3. Run showToken() and copy the token into the hub server's APPS_SCRIPT_TOKEN.
 *   4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *      Copy the /exec URL into the hub server's APPS_SCRIPT_URL.
 * The token check below is what keeps the web app private.
 */

var TABS = ['Quiz Registry', 'Raw Responses', 'Evaluation Results', 'Score Breakdown', 'Areas for Improvement', 'Cohort Performance'];

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
    if (!token || !req.token || !constantTimeEquals_(String(req.token), token)) return json_({ ok: false, error: 'Unauthorized' });
    var p = req.payload || {};
    switch (req.action) {
      case 'ping':
        return json_({ ok: true, data: ping_() });
      case 'createForm':
        return json_({ ok: true, data: createForm_(p.quiz) });
      case 'fetchResponses':
        return json_({ ok: true, data: fetchResponses_(p.form) });
      case 'syncSheets':
        return json_({ ok: true, data: syncSheets_(p.tables) });
      default:
        return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return json_({ ok: false, error: 'Use POST.' });
}

// ---------- setup helpers (run by hand from the editor) ----------

function setupDatabase() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DB_SPREADSHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('Ticket Assessment Hub DB (staging)');
  TABS.forEach(function (name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  props.setProperty('DB_SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('APP_TOKEN')) props.setProperty('APP_TOKEN', Utilities.getUuid() + Utilities.getUuid());
  Logger.log('Database spreadsheet: ' + ss.getUrl());
  Logger.log('Run showToken() to copy the API token into the hub server.');
}

function showToken() {
  Logger.log(PropertiesService.getScriptProperties().getProperty('APP_TOKEN'));
}

function rotateToken() {
  PropertiesService.getScriptProperties().setProperty('APP_TOKEN', Utilities.getUuid() + Utilities.getUuid());
  showToken();
}

// ---------- actions ----------

function ping_() {
  var ss = db_();
  return { spreadsheetUrl: ss.getUrl(), tabs: ss.getSheets().map(function (s) { return s.getName(); }) };
}

function createForm_(quiz) {
  if (!quiz || !quiz.tickets || !quiz.tickets.length) throw new Error('Quiz has no tickets.');
  var ss = db_();
  var form = FormApp.create(quiz.title);
  form.setDescription(
    'Cohort: ' + quiz.cohort +
    (quiz.trainingDay ? '\nTraining day: ' + quiz.trainingDay : '') +
    (quiz.assessmentType ? '\nAssessment type: ' + quiz.assessmentType : '') +
    (quiz.deadline ? '\nSubmission deadline: ' + quiz.deadline : '') +
    '\nQuiz ID: ' + quiz.id
  );
  try { form.setCollectEmail(true); } catch (e) { /* newer API: email collection set in UI */ }
  form.setProgressBar(true);
  form.setAllowResponseEdits(false);
  form.setShowLinkToRespondAgain(false);

  var map = { name: null, cohort: null, tickets: [] };
  map.name = form.addTextItem().setTitle('Full name').setRequired(true).getId();
  map.cohort = form.addListItem().setTitle('Cohort').setChoiceValues([quiz.cohort]).setRequired(true).getId();

  quiz.tickets.forEach(function (t) {
    var help = 'Ticket reference: ' + t.url;
    if (t.description) help += '\n\n' + t.description;
    if (t.context) help += '\n\n' + t.context;
    (t.attachments || []).forEach(function (a) {
      if (/^https?:\/\//i.test(a.url)) help += '\nAttachment: ' + (a.name ? a.name + ' ' : '') + a.url;
    });
    form.addPageBreakItem().setTitle('Ticket ' + t.number).setHelpText(help);
    var ids = {};
    quiz.questions.forEach(function (q) {
      ids[q.key] = form.addParagraphTextItem().setTitle(q.label).setRequired(true).getId();
    });
    map.tickets.push({ ticketId: t.id, number: t.number, url: t.url, items: ids });
  });

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  // Staging: keep the form closed and unpublished. A trainer opens it by hand after approval.
  form.setAcceptingResponses(false);
  try { if (form.setPublished) form.setPublished(false); } catch (e) { /* not available on older runtimes */ }

  return {
    formId: form.getId(),
    publishedUrl: form.getPublishedUrl(),
    editUrl: form.getEditUrl(),
    responseSheetUrl: ss.getUrl(),
    itemMap: map,
    acceptingResponses: false,
  };
}

function fetchResponses_(formRef) {
  if (!formRef || !formRef.formId || !formRef.itemMap) throw new Error('Quiz has no linked Google Form.');
  var form = FormApp.openById(formRef.formId);
  var map = formRef.itemMap;
  return form.getResponses().map(function (r) {
    var byId = {};
    r.getItemResponses().forEach(function (ir) { byId[ir.getItem().getId()] = String(ir.getResponse() || ''); });
    var email = '';
    try { email = r.getRespondentEmail() || ''; } catch (e) { email = ''; }
    return {
      responseId: r.getId(),
      submittedAt: r.getTimestamp().toISOString(),
      traineeEmail: email,
      traineeName: byId[map.name] || '',
      cohort: byId[map.cohort] || '',
      answers: map.tickets.map(function (t) {
        return { ticketId: t.ticketId, q1: byId[t.items.q1] || '', q2: byId[t.items.q2] || '', q3: byId[t.items.q3] || '', q4: byId[t.items.q4] || '' };
      }),
    };
  });
}

// Rewrites only the six hub tabs. Google's own "Form Responses" tabs are never touched.
function syncSheets_(tables) {
  var ss = db_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var counts = {};
    TABS.forEach(function (name) {
      var rows = (tables && tables[name]) || [];
      var sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clearContents();
      if (!rows.length) {
        sh.getRange(1, 1).setValue('No rows yet');
        counts[name] = 0;
        return;
      }
      var headers = [];
      rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (headers.indexOf(k) < 0) headers.push(k); }); });
      var values = [headers].concat(rows.map(function (r) {
        return headers.map(function (h) {
          var v = r[h];
          if (v === null || v === undefined) return '';
          var s = String(v);
          return /^[=+\-@]/.test(s) ? "'" + s : v; // no formula injection
        });
      }));
      sh.getRange(1, 1, values.length, headers.length).setValues(values);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      counts[name] = rows.length;
    });
    return { spreadsheetUrl: ss.getUrl(), rowsWritten: counts, syncedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}

// ---------- utils ----------

function db_() {
  var id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('Database spreadsheet is not set up. Run setupDatabase() in the Apps Script editor.');
  return SpreadsheetApp.openById(id);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function constantTimeEquals_(a, b) {
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}
