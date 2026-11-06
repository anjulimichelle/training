/**
 * Ticket Quiz Form Maker
 *
 * A tiny Google Apps Script web app. The Ticket Assessment Hub opens it with
 * the quiz details in the link; the trainer clicks "Create Google Form" and the
 * form is built in the trainer's own Google Drive. It does nothing else: it does
 * not publish or share the form, does not create response sheets and does not
 * send anything.
 *
 * One-time setup (about two minutes):
 *  1. Go to https://script.google.com and click New project.
 *  2. Replace everything in Code.gs with this file. Name the project
 *     "Ticket Quiz Form Maker" and click Save.
 *  3. Click Deploy > New deployment > gear icon > Web app.
 *       Execute as: User accessing the web app
 *       Who has access: Anyone within your organization
 *     Click Deploy, allow access when Google asks, and copy the Web app URL
 *     (it ends in /exec).
 *  4. In the hub, open Setup and paste that URL into "Form maker link".
 *
 * Because it runs as the person who clicks it, each trainer's forms go into
 * their own Drive. Each trainer allows access the first time they use it.
 */

function doGet(e) {
  var data;
  var error = '';
  try {
    data = readQuiz_(e && e.parameter && e.parameter.q);
  } catch (err) {
    error = String(err && err.message ? err.message : err);
  }
  var t = HtmlService.createTemplate(PAGE_);
  t.error = error;
  t.quiz = data || { title: '', tickets: [] };
  t.payload = data ? JSON.stringify(data).replace(/</g, '\\u003c') : 'null';
  return t.evaluate().setTitle('Ticket Quiz Form Maker').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Called from the page's button. Builds the form and returns its links.
function createQuizForm(quiz) {
  quiz = validate_(quiz);
  var form = FormApp.create(quiz.formTitle);
  form.setTitle(quiz.formTitle);
  form.setDescription('Instructions:\n' + quiz.instructions.map(function (t, i) { return (i + 1) + '. ' + t; }).join('\n'));
  try { form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED); }
  catch (e) { try { form.setCollectEmail(true); } catch (e2) {} }
  form.setProgressBar(true);
  form.setAllowResponseEdits(false);
  form.addTextItem().setTitle('Full name').setRequired(true);
  quiz.tickets.forEach(function (url, i) {
    form.addPageBreakItem().setTitle('Ticket ' + (i + 1)).setHelpText('Ticket link: ' + url);
    quiz.questions.forEach(function (label) {
      form.addParagraphTextItem().setTitle(label).setRequired(true);
    });
  });
  return { id: form.getId(), editUrl: form.getEditUrl(), title: quiz.formTitle };
}

function readQuiz_(q) {
  if (!q) throw new Error('This link has no quiz in it. Open it from the hub with Generate quiz form.');
  var json = Utilities.newBlob(Utilities.base64DecodeWebSafe(q)).getDataAsString('UTF-8');
  return validate_(JSON.parse(json));
}

function validate_(x) {
  if (!x || typeof x !== 'object') throw new Error('The quiz details could not be read.');
  var str = function (v, max) { return String(v == null ? '' : v).slice(0, max); };
  var out = {
    formTitle: str(x.formTitle, 200).trim(),
    instructions: (x.instructions || []).slice(0, 10).map(function (v) { return str(v, 1000); }),
    questions: (x.questions || []).slice(0, 10).map(function (v) { return str(v, 300); }),
    tickets: (x.tickets || []).slice(0, 60).map(function (v) { return str(v, 2000).trim(); }),
  };
  if (!out.formTitle) throw new Error('The quiz has no title.');
  if (!out.tickets.length) throw new Error('The quiz has no ticket links.');
  if (!out.questions.length) throw new Error('The quiz has no questions.');
  out.tickets.forEach(function (u) { if (!/^https?:\/\/\S+$/i.test(u)) throw new Error('Not a link: ' + u); });
  return out;
}

var PAGE_ = [
  '<!doctype html><html><head><base target="_top"><style>',
  'body{font-family:Arial,Helvetica,sans-serif;background:#f4f8f7;color:#1d1d1f;margin:0;padding:24px 16px}',
  '.card{max-width:640px;margin:0 auto;background:#fff;border:1px solid #d8e2e0;border-radius:12px;padding:20px;display:grid;gap:12px}',
  'h1{font-size:1.3rem;margin:0}.eyebrow{font-size:.72rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#0b7a6f}',
  'ol{margin:0;padding-left:22px;font-size:.9rem;overflow-wrap:anywhere}.muted{color:#5f6b69;font-size:.9rem}',
  'button,a.btn{background:#0b7a6f;color:#fff;border:0;border-radius:8px;padding:10px 16px;font-weight:700;font-size:1rem;cursor:pointer;text-decoration:none;display:inline-block}',
  'button:disabled{opacity:.6;cursor:default}.err{background:#fcebea;color:#b3261e;border-radius:8px;padding:10px 12px}.ok{background:#e4f2e1;color:#2e7d32;border-radius:8px;padding:10px 12px}',
  '</style></head><body><div class="card"><span class="eyebrow">Ticket Assessment Hub</span>',
  '<? if (error) { ?><h1>Can\'t make this form</h1><div class="err"><?= error ?></div>',
  '<? } else { ?><h1><?= quiz.formTitle ?></h1>',
  '<p class="muted">Creates a Google Form in your Drive with the instructions at the top, a section for each ticket below, and the four questions in each section. It is not published or shared.</p>',
  '<ol><? for (var i = 0; i < quiz.tickets.length; i++) { ?><li><?= quiz.tickets[i] ?></li><? } ?></ol>',
  '<div><button id="go" onclick="make()">Create Google Form</button></div><div id="out"></div>',
  '<script>var QUIZ = <?!= payload ?>;',
  'function make(){var b=document.getElementById("go");b.disabled=true;b.textContent="Creating…";',
  'google.script.run.withSuccessHandler(function(r){b.style.display="none";var o=document.getElementById("out");o.className="ok";',
  'o.innerHTML="Form created in your Drive. ";var a=document.createElement("a");a.className="btn";a.href=r.editUrl;a.textContent="Open the form";o.appendChild(a);',
  'var p=document.createElement("p");p.className="muted";p.textContent="Next: move it to the right folder, link a response sheet, publish it, and share the link. Then go back to the hub and click Find my form.";o.appendChild(p);})',
  '.withFailureHandler(function(e){b.disabled=false;b.textContent="Create Google Form";var o=document.getElementById("out");o.className="err";o.textContent="Couldn\'t create the form: "+(e&&e.message?e.message:e);})',
  '.createQuizForm(QUIZ);}</script><? } ?></div></body></html>'
].join('\n');
