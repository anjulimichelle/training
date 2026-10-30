/**
 * Daily sync from the Care Transformation Comms Kit into the Macro Finder.
 *
 * Install: paste into the Macro Finder spreadsheet's Apps Script project, run
 * syncFromCommsKit once to authorize, then add a time-driven trigger (daily).
 *
 * Rules (from the UPDATE PROCESS tab):
 *  - Never overwrite an Approved macro. A changed source becomes a Review row.
 *  - A Pending Review macro takes the new wording directly, with a Review row noting it.
 *  - New Comms Kit macros are added to the Master as Pending Review, uncategorized.
 *  - Macros that disappear from the Comms Kit get a Review row; nothing is deleted.
 */

const COMMS_KIT_ID = '1T2Q7uBOMW6zwBAG8EZ3_e8RMmjqF9qeRgDe4tPB9i-Y';
const COMMS_KIT_TAB = 'All Macros — New taxonomy';
const SYNC_SHEETS = {
  master: '📚 MASTER MACRO LIBRARY',
  review: '⚠️ MACRO REVIEW',
  process: '🛠 UPDATE PROCESS',
};
const MODEL_TO_TICKET = { 'DHJ': 'DHJ', 'FCF': 'FCF', 'DHJ, FCF': 'Both' };

function syncFromCommsKit() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(SYNC_SHEETS.master);
  const review = ss.getSheetByName(SYNC_SHEETS.review);

  // Comms Kit columns: A Transformation, B Audience, C Category, D Subcategory, E Title, F Subject,
  // G Body, H Key Points / Agent Instructions, I Membership Model.
  const source = SpreadsheetApp.openById(COMMS_KIT_ID).getSheetByName(COMMS_KIT_TAB).getDataRange().getValues();
  // Two macros can share a title, so each is keyed by title plus its occurrence number.
  const incoming = {};
  const nth = {};
  source.slice(1).forEach((r, i) => {
    if (r[1] !== 'Customer' || !r[4]) return;
    nth[r[4]] = (nth[r[4]] || 0) + 1;
    incoming[r[4] + '#' + nth[r[4]]] = { title: r[4], row: i + 2, body: r[6], notes: r[7], model: r[8], path: r[2] + ' › ' + r[3] };
  });

  const values = master.getDataRange().getValues();
  const head = values[0];
  const c = name => head.indexOf(name);
  const C = { id: c('Macro ID'), name: c('Macro Name'), body: c('Full Macro Response'), notes: c('Required CRM Actions'),
              status: c('Review Status'), source: c('Wording Source'), row: c('Source Row') };

  const openSync = new Set(review.getDataRange().getValues().slice(1)
    .filter(r => r[7] === 'Pending Review').map(r => String(r[5])).filter(e => e.indexOf('sync:') === 0));
  const reviewRows = [];
  const flag = (kind, id, count, details, key, action) => {
    const evidence = 'sync:' + key;
    if (!openSync.has(evidence)) reviewRows.push(['', kind, id, count, details, evidence, action, 'Pending Review', '', '', '']);
  };

  const seen = new Set();
  const masterNth = {};
  const masterKey = {};
  values.slice(1).map((r, i) => ({ r: r, i: i }))
    .filter(x => String(x.r[C.source]).indexOf('Comms Kit') === 0)
    .sort((a, b) => a.r[C.row] - b.r[C.row])
    .forEach(x => {
      const t = x.r[C.name];
      masterNth[t] = (masterNth[t] || 0) + 1;
      masterKey[x.i] = t + '#' + masterNth[t];
    });
  let changed = 0, added = 0, removed = 0;
  values.slice(1).forEach((r, i) => {
    if (!(i in masterKey)) return;
    const title = r[C.name];
    const key = masterKey[i];
    const src = incoming[key];
    if (!src) {
      removed++;
      flag('Removed from Comms Kit', r[C.id], 1, `"${title}" is no longer in the Comms Kit tab.`, 'removed:' + r[C.id],
           'Confirm, then set the macro to Archived.');
      return;
    }
    seen.add(key);
    if (src.body === r[C.body] && src.notes === r[C.notes]) return;
    changed++;
    if (r[C.status] === 'Approved') {
      flag('Source changed', r[C.id], 1, `Comms Kit wording changed for an Approved macro. New response:\n${src.body}`,
           'changed:' + r[C.id] + ':' + digest_(src.body + src.notes),
           'Review the new wording. If approved, copy it into the Master and keep the status Approved.');
    } else {
      master.getRange(i + 2, C.body + 1).setValue(src.body);
      master.getRange(i + 2, C.notes + 1).setValue(src.notes);
      flag('Updated from source', r[C.id], 1, 'A Pending Review macro took the new Comms Kit wording.',
           'updated:' + r[C.id] + ':' + digest_(src.body + src.notes), 'Review as part of approval.');
    }
  });

  let nextId = Math.max(0, ...values.slice(1).map(r => parseInt(String(r[C.id]).replace('MF-', ''), 10) || 0)) + 1;
  Object.keys(incoming).filter(k => !seen.has(k)).forEach(key => {
    const src = incoming[key];
    const title = src.title;
    const id = 'MF-' + String(nextId++).padStart(3, '0');
    const row = new Array(head.length).fill('');
    row[C.id] = id; row[C.name] = title; row[C.body] = src.body; row[C.notes] = src.notes;
    row[c('Ticket Type')] = MODEL_TO_TICKET[src.model] || 'Other / Undefined';
    row[c('Main Category')] = 'Other'; row[c('Subcategory')] = 'General Responses';
    row[C.status] = 'Pending Review'; row[C.source] = 'Comms Kit › ' + COMMS_KIT_TAB; row[C.row] = src.row;
    row[c('Comms Kit Category Path')] = src.path;
    master.appendRow(row);
    added++;
    flag('New macro in Comms Kit', id, 1, `"${title}" was added from the Comms Kit and needs a category, scenario and keywords.`,
         'new:' + key, 'Categorize, then approve.');
  });

  if (reviewRows.length) {
    const start = review.getLastRow() + 1;
    reviewRows.forEach((r, i) => (r[0] = 'RV-S' + Utilities.formatDate(new Date(), 'UTC', 'yyMMdd') + '-' + (i + 1)));
    review.getRange(start, 1, reviewRows.length, reviewRows[0].length).setValues(reviewRows);
  }
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE MMM d yyyy HH:mm');
  ss.getSheetByName(SYNC_SHEETS.process).getRange('B2')
    .setValue(`${stamp}: automatic sync. ${added} new, ${changed} changed, ${removed} removed; ${reviewRows.length} review rows added.`);
}

function digest_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s))
    .map(b => (b & 0xff).toString(16).padStart(2, '0')).join('').slice(0, 10);
}
