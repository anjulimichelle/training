/**
 * Macro Finder web app (Option A on the UPDATE PROCESS tab).
 *
 * Paste into Extensions › Apps Script of the Macro Finder spreadsheet together with
 * Index.html, then Deploy › New deployment › Web app:
 *   Execute as: Me    Who has access: Anyone in your organization
 * Each agent gets a private search page. It reads the Master Macro Library as the
 * owner, so agents need no access to the sheet and can't overwrite each other.
 */

const SHEETS = {
  master: '📚 MASTER MACRO LIBRARY',
  categories: '🗂 CATEGORY MAP',
  synonyms: '🔤 SEARCH SYNONYMS',
};

// Only Approved macros reach agents. Flip to true for a reviewer-only preview deployment.
const INCLUDE_PENDING = false;
const MAX_RESULTS = 25;

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Macro Finder')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getFilterOptions() {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEETS.categories).getDataRange().getValues().slice(1);
  const subsByMain = {};
  rows.filter(r => r[0] && r[1]).forEach(r => (subsByMain[r[0]] = subsByMain[r[0]] || []).push(r[1]));
  return {
    ticketTypes: ['DHJ', 'FCF', 'Both', 'Other / Undefined'],
    subsByMain: subsByMain,
    scenarios: rows.map(r => r[8]).filter(String),
  };
}

function loadMacros_() {
  const values = SpreadsheetApp.getActive().getSheetByName(SHEETS.master).getDataRange().getValues();
  const head = values[0];
  const col = name => head.indexOf(name);
  const c = {
    id: col('Macro ID'), name: col('Macro Name'), tt: col('Ticket Type'), main: col('Main Category'),
    sub: col('Subcategory'), scen: col('Customer Scenario'), kw: col('Search Keywords'),
    phrases: col('Common Customer Phrases'), when: col('When to Use'), body: col('Full Macro Response'),
    crm: col('Required CRM Actions'), policy: col('Policy Reminders'), status: col('Review Status'),
    kbLink: col('KB Link'), kbArticle: col('KB Article(s)'), src: col('Source Row'),
  };
  return values.slice(1).filter(r => r[c.id]).map(r => ({
    id: r[c.id], name: r[c.name], ticketType: r[c.tt], main: r[c.main], sub: r[c.sub], scenario: r[c.scen],
    when: r[c.when], body: r[c.body], crm: r[c.crm], policy: r[c.policy], status: r[c.status],
    kbLink: r[c.kbLink], kbLabel: r[c.kbArticle] ? 'KB: ' + r[c.kbArticle] : 'Comms Kit row ' + r[c.src],
    _name: norm_(r[c.name]),
    _meta: norm_([r[c.tt], r[c.main], r[c.sub], r[c.scen], r[c.kw], r[c.phrases], r[c.when], r[c.kbArticle]].join(' ')),
    _body: norm_(r[c.body]),
  }));
}

function norm_(s) {
  return String(s || '').toLowerCase().replace(/['’]/g, '');
}

// Same scoring as the sheet. Each typed word scores 14 when it appears in the name (13 if
// only a synonym does), 12 / 11.5 in keywords, phrases, category or scenario, and 11 / 10.5
// in the response text, plus 1 when the exact word also appears outside the name.
function buildPatterns_(query) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.synonyms);
  const rows = sheet.getDataRange().getValues().slice(1);
  const triggers = rows.filter(r => r[0]).map(r => ({ re: new RegExp(r[0]), expand: r[1] }));
  const stop = new Set(rows.map(r => String(r[4]).toLowerCase()).filter(String));
  const words = norm_(query).replace(/[^a-z0-9&]+/g, ' ').trim().split(' ')
    .filter(w => w.length >= 2 && !stop.has(w));
  return words.map(w => {
    const extra = triggers.filter(t => t.re.test(w)).map(t => t.expand);
    return { exact: new RegExp(w), any: new RegExp([w].concat(extra).join('|')) };
  });
}

function search(query, filters) {
  filters = filters || {};
  const patterns = buildPatterns_(query || '');
  const allowed = INCLUDE_PENDING ? ['Approved', 'Pending Review'] : ['Approved'];
  const tt = filters.ticketType;
  return loadMacros_()
    .filter(m => allowed.indexOf(m.status) !== -1)
    .filter(m => !tt || m.ticketType === tt || (m.ticketType === 'Both' && (tt === 'DHJ' || tt === 'FCF')))
    .filter(m => !filters.main || m.main === filters.main)
    .filter(m => !filters.sub || m.sub === filters.sub)
    .filter(m => !filters.scenario || m.scenario === filters.scenario)
    .map(m => {
      const score = patterns.length === 0 ? 1 : patterns.reduce((acc, p) => acc +
        (p.exact.test(m._name) ? 14 : p.any.test(m._name) ? 13 : p.exact.test(m._meta) ? 12 :
         p.any.test(m._meta) ? 11.5 : p.exact.test(m._body) ? 11 : p.any.test(m._body) ? 10.5 : 0) +
        (p.exact.test(m._meta + ' ' + m._body) ? 1 : 0), 0);
      return { m: m, score: score };
    })
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, MAX_RESULTS)
    .map(x => {
      const m = Object.assign({}, x.m);
      delete m._name; delete m._meta; delete m._body;
      return m;
    });
}
