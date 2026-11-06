'use strict';
/*
 * Ticket Assessment & Scoring Hub: staging server.
 * Zero-framework Node HTTP server. Serves the trainer app, the staging trainee form,
 * and a JSON API. Secrets (Claude key, Apps Script token) stay in server env vars.
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const core = require('./core');
const { Store } = require('./store');
const evaluator = require('./evaluator');
const google = require('./google');
const demo = require('./demo');
const exporter = require('./exporter');

const ROOT = path.join(__dirname, '..');
const PUBLIC = path.join(ROOT, 'public');
const PORT = Number(process.env.PORT || 4317);
const HOST = process.env.HOST || '127.0.0.1';
const DATA_FILE = process.env.DATA_FILE || path.join(ROOT, 'data', 'staging-db.json');
const UPLOAD_DIR = path.join(path.dirname(DATA_FILE), 'uploads');
const PASSCODE = process.env.TRAINER_PASSCODE || '';

const store = new Store(DATA_FILE);
fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// ---------- helpers ----------
class HttpError extends Error {
  constructor(status, message, extra) {
    super(message);
    this.status = status;
    this.extra = extra;
  }
}

function send(res, status, body, headers) {
  const isBuf = Buffer.isBuffer(body);
  const payload = isBuf || typeof body === 'string' ? body : JSON.stringify(body);
  res.writeHead(status, {
    'Content-Type': isBuf || typeof body === 'string' ? 'text/plain; charset=utf-8' : 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
    ...(headers || {}),
  });
  res.end(payload);
}

function readBody(req, limit = 12 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) {
        reject(new HttpError(413, 'Request is too large.'));
        req.destroy();
      } else chunks.push(c);
    });
    req.on('end', () => {
      if (!chunks.length) return resolve({});
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch (_) {
        reject(new HttpError(400, 'Request body is not valid JSON.'));
      }
    });
    req.on('error', reject);
  });
}

function safeEqual(a, b) {
  const ab = Buffer.from(String(a));
  const bb = Buffer.from(String(b));
  return ab.length === bb.length && crypto.timingSafeEqual(ab, bb);
}

function requireTrainer(req) {
  if (!PASSCODE) return;
  const got = req.headers['x-trainer-passcode'] || '';
  if (!safeEqual(got, PASSCODE)) throw new HttpError(401, 'Trainer passcode required.');
}

function str(v, max = 20000) {
  return String(v == null ? '' : v).slice(0, max);
}

function hashAnswers(answers) {
  return crypto.createHash('sha256').update(JSON.stringify(answers)).digest('hex');
}

const clone = (o) => JSON.parse(JSON.stringify(o));

// Accept only known fields from the browser, so the client cannot write form/audit data.
function sanitizeQuiz(input, existing) {
  const q = existing ? clone(existing) : { id: core.uid('quiz'), createdAt: new Date().toISOString(), form: null, status: 'Draft' };
  q.title = str(input.title, 300).trim();
  q.cohort = str(input.cohort, 200).trim();
  q.trainingDay = str(input.trainingDay, 100);
  q.assessmentType = str(input.assessmentType, 100);
  q.assignedTrainees = Array.isArray(input.assignedTrainees) ? input.assignedTrainees.map((x) => str(x, 100)) : [];
  q.deadline = str(input.deadline, 40);
  q.passingThreshold = input.passingThreshold === '' || input.passingThreshold == null ? null : Number(input.passingThreshold);
  q.tickets = (Array.isArray(input.tickets) ? input.tickets : []).slice(0, 50).map((t) => ({
    id: str(t.id, 100) || core.uid('tkt'),
    url: str(t.url, 2000).trim(),
    description: str(t.description),
    context: str(t.context),
    attachments: (Array.isArray(t.attachments) ? t.attachments : []).slice(0, 20).map((a) => ({ name: str(a.name, 200), url: str(a.url, 2000) })),
    rubric: {
      totalPoints: Number(t.rubric && t.rubric.totalPoints) || 0,
      expectedIntent: str(t.rubric && t.rubric.expectedIntent),
      requiredCrmActions: str(t.rubric && t.rubric.requiredCrmActions),
      requiredResolution: str(t.rubric && t.rubric.requiredResolution),
      expectedC: str(t.rubric && t.rubric.expectedC),
      expectedCP: str(t.rubric && t.rubric.expectedCP),
      criticalErrors: str(t.rubric && t.rubric.criticalErrors),
      acceptableAlternatives: str(t.rubric && t.rubric.acceptableAlternatives),
      autoFailConditions: str(t.rubric && t.rubric.autoFailConditions),
      criteria: ((t.rubric && t.rubric.criteria) || []).slice(0, 40).map((c) => ({
        id: str(c.id, 100) || core.uid('crit'),
        name: str(c.name, 200),
        category: core.CRITERION_CATEGORIES.some((x) => x.key === c.category) ? c.category : 'other',
        maxPoints: Math.max(0, Number(c.maxPoints) || 0),
        guidance: str(c.guidance),
        deductions: (c.deductions || []).slice(0, 20).map((d) => ({ description: str(d.description, 500), points: Math.max(0, Number(d.points) || 0) })),
      })),
    },
    references: {
      policyContext: str(t.references && t.references.policyContext),
      kbReferences: str(t.references && t.references.kbReferences),
      expectedResolution: str(t.references && t.references.expectedResolution),
      sampleC: str(t.references && t.references.sampleC),
      sampleCP: str(t.references && t.references.sampleCP),
      evaluatorInstructions: str(t.references && t.references.evaluatorInstructions),
    },
  }));
  q.updatedAt = new Date().toISOString();
  // If the trainee-facing part changed after a form exists, mark the form out of date.
  if (existing && existing.form) {
    const before = JSON.stringify(core.traineeView(existing));
    const after = JSON.stringify(core.traineeView(q));
    if (before !== after) q.form.outOfDate = true;
  }
  return q;
}

function config() {
  return {
    mode: 'staging',
    evaluator: evaluator.evaluatorName(),
    claudeConfigured: evaluator.claudeConfigured(),
    claudeModel: evaluator.claudeConfigured() ? evaluator.MODEL() : null,
    googleConfigured: google.configured(),
    googleWritesApproved: google.writesApproved(),
    passcodeRequired: !!PASSCODE,
    publicBaseUrl: process.env.PUBLIC_BASE_URL || '',
  };
}

function fullState(db) {
  return {
    config: config(),
    trainees: db.trainees,
    quizzes: db.quizzes,
    submissions: db.submissions,
    evaluations: db.evaluations,
    rows: core.assignmentRows(db),
  };
}

function findQuiz(db, id) {
  const q = db.quizzes.find((x) => x.id === id);
  if (!q) throw new HttpError(404, 'Quiz not found.');
  return q;
}

function recordSubmission(db, quiz, input, source, externalId) {
  if (externalId) {
    const dup = db.submissions.find((s) => s.quizId === quiz.id && s.externalResponseId === externalId);
    if (dup) return { submission: dup, duplicate: true };
  }
  const email = str(input.traineeEmail, 300).trim().toLowerCase();
  const name = str(input.traineeName, 200).trim();
  if (!name || !email) throw new HttpError(400, 'Name and email are required.');
  const trainee = db.trainees.find((t) => t.email.toLowerCase() === email);
  const byTicket = new Map((input.answers || []).map((a) => [a.ticketId, a]));
  const answers = quiz.tickets.map((t, i) => {
    const a = byTicket.get(t.id) || {};
    return { ticketId: t.id, ticketNumber: i + 1, ticketUrl: t.url, q1: str(a.q1), q2: str(a.q2), q3: str(a.q3), q4: str(a.q4) };
  });
  const sub = {
    id: core.uid('sub'),
    quizId: quiz.id,
    traineeId: trainee ? trainee.id : null,
    traineeName: name,
    traineeEmail: email,
    cohort: str(input.cohort, 200) || quiz.cohort,
    submittedAt: input.submittedAt || new Date().toISOString(),
    receivedAt: new Date().toISOString(),
    source,
    externalResponseId: externalId || null,
    answers,
    rawHash: hashAnswers(answers),
    evaluationState: 'none',
    lastError: null,
  };
  db.submissions.push(sub);
  store.audit(db, 'submission.received', { submissionId: sub.id, quizId: quiz.id, source });
  return { submission: sub, duplicate: false };
}

// ---------- routes ----------
const routes = [];
function route(method, pattern, handler, opts) {
  const keys = [];
  const re = new RegExp('^' + pattern.replace(/:(\w+)/g, (_, k) => (keys.push(k), '([^/]+)')) + '$');
  routes.push({ method, re, keys, handler, public: opts && opts.public });
}

route('GET', '/api/health', async () => ({ ok: true, config: config() }), { public: true });

route('GET', '/api/state', async () => fullState(store.db));

// Trainee roster
route('POST', '/api/trainees', async (req, p, body) =>
  store.mutate((db) => {
    const name = str(body.name, 200).trim();
    const email = str(body.email, 300).trim().toLowerCase();
    const cohort = str(body.cohort, 200).trim();
    if (!name || !email || !cohort) throw new HttpError(400, 'Name, email and cohort are required.');
    if (db.trainees.some((t) => t.email.toLowerCase() === email)) throw new HttpError(409, 'A trainee with that email already exists.');
    const t = { id: core.uid('trn'), name, email, cohort, createdAt: new Date().toISOString() };
    db.trainees.push(t);
    return t;
  })
);
route('DELETE', '/api/trainees/:id', async (req, p) =>
  store.mutate((db) => {
    if (db.submissions.some((s) => s.traineeId === p.id)) throw new HttpError(409, 'This trainee has submissions. Keep the record for the audit trail.');
    db.trainees = db.trainees.filter((t) => t.id !== p.id);
    db.quizzes.forEach((q) => (q.assignedTrainees = q.assignedTrainees.filter((x) => x !== p.id)));
    return { ok: true };
  })
);

// Quizzes
route('POST', '/api/quizzes', async (req, p, body) =>
  store.mutate((db) => {
    const existing = body.id ? db.quizzes.find((q) => q.id === body.id) : null;
    const q = sanitizeQuiz(body, existing);
    if (existing) db.quizzes[db.quizzes.indexOf(existing)] = q;
    else db.quizzes.push(q);
    store.audit(db, existing ? 'quiz.updated' : 'quiz.created', { quizId: q.id });
    return { quiz: q, validation: core.validateQuiz(q) };
  })
);
route('DELETE', '/api/quizzes/:id', async (req, p) =>
  store.mutate((db) => {
    if (db.submissions.some((s) => s.quizId === p.id)) throw new HttpError(409, 'This quiz has submissions and cannot be deleted. Set its status to Closed instead.');
    findQuiz(db, p.id);
    db.quizzes = db.quizzes.filter((q) => q.id !== p.id);
    store.audit(db, 'quiz.deleted', { quizId: p.id });
    return { ok: true };
  })
);
route('POST', '/api/quizzes/:id/status', async (req, p, body) =>
  store.mutate((db) => {
    const q = findQuiz(db, p.id);
    if (!core.QUIZ_STATUSES.includes(body.status)) throw new HttpError(400, 'Unknown status.');
    if ((body.status === 'Ready' || body.status === 'Sent') && !q.form) throw new HttpError(400, 'Generate the form before marking the quiz Ready or Sent.');
    q.status = body.status;
    q.updatedAt = new Date().toISOString();
    store.audit(db, 'quiz.status', { quizId: q.id, status: body.status, note: 'Status label only. Nothing was sent to trainees.' });
    return q;
  })
);

// Generate the trainee form, only after the trainer approves the preview.
route('POST', '/api/quizzes/:id/generate-form', async (req, p, body) => {
  const q = findQuiz(store.db, p.id);
  if (!body.approved) throw new HttpError(400, 'Review the preview and confirm approval before generating the form.');
  const v = core.validateQuiz(q);
  if (v.errors.length) throw new HttpError(422, 'Fix these before generating the form.', { errors: v.errors });
  const target = body.target === 'google' ? 'google' : 'staging';
  let form;
  if (target === 'google') {
    const out = await google.createForm({ ...core.traineeView(q), assignedTraineeEmails: q.assignedTrainees.map((id) => (store.db.trainees.find((t) => t.id === id) || {}).email).filter(Boolean) });
    form = { mode: 'google', formId: out.formId, formUrl: out.publishedUrl, editUrl: out.editUrl, responseSheetUrl: out.responseSheetUrl, itemMap: out.itemMap, acceptingResponses: out.acceptingResponses };
  } else {
    form = { mode: 'staging', formUrl: `/form/${q.id}`, editUrl: null, responseSheetUrl: null };
  }
  return store.mutate((db) => {
    const qq = findQuiz(db, p.id);
    qq.form = { ...form, createdAt: new Date().toISOString(), approvedAt: new Date().toISOString(), approvedBy: str(body.approvedBy, 200) || 'trainer', outOfDate: false };
    if (qq.status === 'Draft') qq.status = 'Ready';
    store.audit(db, 'form.generated', { quizId: qq.id, target });
    return qq;
  });
});

// Pull responses from the linked Google Form. Deduplicated by Google response ID.
route('POST', '/api/quizzes/:id/sync-responses', async (req, p) => {
  const q = findQuiz(store.db, p.id);
  if (!q.form || q.form.mode !== 'google') throw new HttpError(400, 'This quiz uses the staging form. Its responses are already here.');
  const responses = await google.fetchResponses(q);
  return store.mutate((db) => {
    const qq = findQuiz(db, p.id);
    let added = 0;
    responses.forEach((r) => {
      const out = recordSubmission(db, qq, r, 'google-form', r.responseId);
      if (!out.duplicate) added++;
    });
    return { fetched: responses.length, added, duplicatesSkipped: responses.length - added };
  });
});

// Staging-only helper: add fictional sample submissions for a quiz.
route('POST', '/api/quizzes/:id/sample-submissions', async (req, p) =>
  store.mutate((db) => {
    const q = findQuiz(db, p.id);
    if (!q.form) throw new HttpError(400, 'Generate the staging form first.');
    const sets = q.tickets.length >= 2 && q.tickets[0].url.includes('fictional-tickets') ? demo.sampleAnswers(q) : null;
    const assigned = q.assignedTrainees.map((id) => db.trainees.find((t) => t.id === id)).filter(Boolean).slice(0, 2);
    const made = assigned.map((t, i) => {
      const answers = sets ? sets[i === 0 ? 'strong' : 'weak'] : q.tickets.map((tk) => ({ ticketId: tk.id, q1: 'Sample answer', q2: 'Sample CRM action', q3: 'Sample reply to C', q4: 'N/A' }));
      return recordSubmission(db, q, { traineeName: t.name, traineeEmail: t.email, cohort: t.cohort, answers }, 'staging-sample').submission;
    });
    return { added: made.length };
  })
);

// Public trainee endpoints (staging form). Only the trainee-safe view leaves the server.
route(
  'GET',
  '/api/public/quiz/:id',
  async (req, p) => {
    const q = store.db.quizzes.find((x) => x.id === p.id);
    if (!q || !q.form || q.form.mode !== 'staging') throw new HttpError(404, 'This form is not available.');
    return { quiz: core.traineeView(q), closed: q.status === 'Closed' };
  },
  { public: true }
);
route(
  'POST',
  '/api/public/quiz/:id/submit',
  async (req, p, body) =>
    store.mutate((db) => {
      const q = db.quizzes.find((x) => x.id === p.id);
      if (!q || !q.form || q.form.mode !== 'staging') throw new HttpError(404, 'This form is not available.');
      if (q.status === 'Closed') throw new HttpError(409, 'This quiz is closed.');
      const { submission } = recordSubmission(db, q, body, 'staging-form');
      return { ok: true, submissionId: submission.id, submittedAt: submission.submittedAt };
    }),
  { public: true }
);

// Evaluation
route('POST', '/api/submissions/:id/evaluate', async (req, p, body) => {
  const prep = await store.mutate((db) => {
    const sub = db.submissions.find((s) => s.id === p.id);
    if (!sub) throw new HttpError(404, 'Submission not found.');
    const current = db.evaluations.find((e) => e.submissionId === sub.id && !e.supersededAt);
    if (current && !body.reevaluate) throw new HttpError(409, 'This submission already has an evaluation. Choose "Request reevaluation" to create a new version.');
    if (sub.evaluationState === 'running') throw new HttpError(409, 'An evaluation is already running for this submission.');
    sub.evaluationState = 'running';
    sub.lastError = null;
    return { sub: clone(sub), quiz: clone(findQuiz(db, sub.quizId)) };
  });
  let result;
  try {
    result = await evaluator.evaluateSubmission(prep.quiz, prep.sub);
  } catch (err) {
    await store.mutate((db) => {
      const sub = db.submissions.find((s) => s.id === p.id);
      sub.evaluationState = 'failed';
      sub.lastError = { at: new Date().toISOString(), message: err.message || String(err) };
      store.audit(db, 'evaluation.failed', { submissionId: sub.id, error: sub.lastError.message });
    });
    throw new HttpError(502, `Evaluation failed: ${err.message || err}. The trainee's answers are unchanged; you can retry.`);
  }
  return store.mutate((db) => {
    const sub = db.submissions.find((s) => s.id === p.id);
    if (hashAnswers(sub.answers) !== sub.rawHash) throw new HttpError(500, 'Raw submission integrity check failed. Evaluation not saved.');
    const prior = db.evaluations.filter((e) => e.submissionId === sub.id);
    const current = prior.find((e) => !e.supersededAt);
    const now = new Date().toISOString();
    if (current) {
      current.supersededAt = now;
      current.supersededReason = str(body.reason, 1000) || 'Trainer requested reevaluation';
    }
    const ev = { id: core.uid('eval'), submissionId: sub.id, quizId: sub.quizId, version: prior.length + 1, createdAt: now, ai: result, trainer: null, history: [] };
    db.evaluations.push(ev);
    sub.evaluationState = 'done';
    store.audit(db, current ? 'evaluation.reevaluated' : 'evaluation.created', { submissionId: sub.id, evaluationId: ev.id, evaluator: result.evaluator });
    return ev;
  });
});

// Trainer review. Never touches ev.ai; edits live in ev.trainer.
route('POST', '/api/evaluations/:id/review', async (req, p, body) =>
  store.mutate((db) => {
    const ev = db.evaluations.find((e) => e.id === p.id);
    if (!ev) throw new HttpError(404, 'Evaluation not found.');
    if (ev.supersededAt) throw new HttpError(409, 'This evaluation was superseded by a newer version.');
    const quiz = findQuiz(db, ev.quizId);
    const now = new Date().toISOString();
    const reviewer = str(body.reviewer, 200) || 'trainer';
    const before = ev.trainer ? clone(ev.trainer) : null;
    const tr = ev.trainer ? ev.trainer : { tickets: clone(ev.ai.tickets), comments: '', overrideResult: '', createdAt: now };
    if (Array.isArray(body.tickets)) {
      body.tickets.forEach((bt, ti) => {
        const t = tr.tickets[ti];
        if (!t || !bt) return;
        (bt.criteria || []).forEach((bc, ci) => {
          const c = t.criteria[ci];
          if (!c || !bc) return;
          if (bc.awarded !== undefined) c.awarded = Number(bc.awarded);
          if (bc.reason !== undefined) c.reason = str(bc.reason, 2000);
          c.trainerEdited = true;
        });
        if (Array.isArray(bt.improvements)) t.improvements = bt.improvements.map((i) => ({ category: core.IMPROVEMENT_CATEGORIES.includes(i.category) ? i.category : 'Other', detail: str(i.detail, 2000) })).filter((i) => i.detail);
        if (bt.trainerNotes !== undefined) t.trainerNotes = str(bt.trainerNotes, 4000);
        core.recomputeTicket(t);
      });
    }
    if (body.comments !== undefined) tr.comments = str(body.comments, 8000);
    if (body.overrideResult !== undefined) {
      if (!['', 'Pass', 'Fail', 'Needs retake'].includes(body.overrideResult)) throw new HttpError(400, 'Unknown result override.');
      tr.overrideResult = body.overrideResult;
    }
    tr.overall = core.summarize(tr.tickets, quiz);
    tr.updatedAt = now;
    tr.updatedBy = reviewer;
    if (body.approve) {
      tr.approvedAt = now;
      tr.approvedBy = reviewer;
    } else if (body.unapprove) {
      tr.approvedAt = null;
    }
    ev.trainer = tr;
    ev.history.push({ at: now, by: reviewer, action: body.approve ? 'approved' : 'edited', before: before ? { earned: before.overall && before.overall.earned, approvedAt: before.approvedAt || null } : null, after: { earned: tr.overall.earned } });
    store.audit(db, body.approve ? 'evaluation.approved' : 'evaluation.edited', { evaluationId: ev.id, by: reviewer });
    return ev;
  })
);

// Uploads (staging: stored on disk next to the data file, served by random name)
route('POST', '/api/uploads', async (req, p, body) => {
  const m = /^data:([\w/+.-]+);base64,(.+)$/.exec(String(body.dataUrl || ''));
  if (!m) throw new HttpError(400, 'Upload must be a data URL.');
  const allowed = { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/gif': '.gif', 'image/webp': '.webp', 'application/pdf': '.pdf' };
  if (!allowed[m[1]]) throw new HttpError(400, 'Only PNG, JPEG, GIF, WebP and PDF files are allowed.');
  const buf = Buffer.from(m[2], 'base64');
  if (buf.length > 8 * 1024 * 1024) throw new HttpError(413, 'Files must be 8 MB or smaller.');
  const name = crypto.randomBytes(12).toString('hex') + allowed[m[1]];
  fs.writeFileSync(path.join(UPLOAD_DIR, name), buf);
  return { name: str(body.name, 200) || name, url: `/uploads/${name}` };
});

// Exports
route('GET', '/api/export', async (req, p, body, url) => {
  const kind = url.searchParams.get('kind') || 'all';
  const format = url.searchParams.get('format') === 'csv' ? 'csv' : 'xlsx';
  const filter = { traineeId: url.searchParams.get('traineeId') || '', cohort: url.searchParams.get('cohort') || '', quizId: url.searchParams.get('quizId') || '' };
  const out = await exporter.build(store.db, kind, format, filter);
  store.mutate((db) => store.audit(db, 'export', { kind, format, filter }));
  return { __file: out };
});

// Google Sheets sync (blocked until approved)
route('POST', '/api/google/sync-sheets', async () => {
  const tables = core.exportTables(store.db);
  const out = await google.syncSheets(tables);
  await store.mutate((db) => store.audit(db, 'sheets.synced', {}));
  return out;
});
route('POST', '/api/google/ping', async () => google.ping());

// Demo data
route('POST', '/api/demo/seed', async () =>
  store.mutate((db) => {
    const now = new Date().toISOString();
    demo.TRAINEES.forEach((t) => {
      if (!db.trainees.some((x) => x.email === t.email)) db.trainees.push({ id: core.uid('trn'), ...t, createdAt: now });
    });
    const q = demo.sampleQuiz(db.trainees);
    db.quizzes.push(q);
    store.audit(db, 'demo.seeded', { quizId: q.id });
    return { quizId: q.id };
  })
);
route('POST', '/api/demo/reset', async (req, p, body) => {
  if (body.confirm !== 'RESET STAGING DATA') throw new HttpError(400, 'Type the confirmation phrase to reset staging data.');
  await store.reset();
  return { ok: true };
});

// ---------- static ----------
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.jpg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp', '.pdf': 'application/pdf', '.svg': 'image/svg+xml' };
function serveFile(res, file) {
  fs.readFile(file, (err, buf) => {
    if (err) return send(res, 404, 'Not found');
    res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream', 'X-Content-Type-Options': 'nosniff', 'Cache-Control': 'no-cache' });
    res.end(buf);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const pathname = decodeURIComponent(url.pathname);
  try {
    if (pathname.startsWith('/api/')) {
      const r = routes.find((x) => x.method === req.method && x.re.test(pathname));
      if (!r) throw new HttpError(404, 'Unknown API route.');
      if (!r.public) requireTrainer(req);
      const m = r.re.exec(pathname);
      const params = {};
      r.keys.forEach((k, i) => (params[k] = m[i + 1]));
      const body = req.method === 'GET' ? {} : await readBody(req);
      const out = await r.handler(req, params, body, url);
      if (out && out.__file) {
        const f = out.__file;
        res.writeHead(200, { 'Content-Type': f.contentType, 'Content-Disposition': `attachment; filename="${f.filename}"`, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
        return res.end(f.buffer);
      }
      return send(res, 200, out);
    }
    if (pathname === '/core.js') return serveFile(res, path.join(__dirname, 'core.js'));
    if (/^\/form\/[\w-]+$/.test(pathname)) return serveFile(res, path.join(PUBLIC, 'form.html'));
    if (pathname.startsWith('/uploads/')) {
      const name = path.basename(pathname);
      if (!/^[a-f0-9]{24}\.\w+$/.test(name)) return send(res, 404, 'Not found');
      return serveFile(res, path.join(UPLOAD_DIR, name));
    }
    const rel = pathname === '/' ? 'index.html' : pathname.slice(1);
    const file = path.normalize(path.join(PUBLIC, rel));
    if (!file.startsWith(PUBLIC)) return send(res, 404, 'Not found');
    return serveFile(res, file);
  } catch (err) {
    const status = err.status || 500;
    if (!err.status) console.error(err);
    send(res, status, { error: err.message || 'Server error', ...(err.extra || {}) });
  }
});

if (require.main === module) {
  server.listen(PORT, HOST, () => {
    const c = config();
    console.log(`Ticket Assessment Hub (STAGING) on http://${HOST}:${PORT}`);
    console.log(`  Evaluator: ${c.evaluator === 'claude' ? `Claude (${c.claudeModel})` : 'offline heuristic (set ANTHROPIC_API_KEY for Claude)'}`);
    console.log(`  Google: ${c.googleConfigured ? 'connected' : 'not connected'}${c.googleConfigured ? (c.googleWritesApproved ? ', external writes APPROVED' : ', external writes blocked') : ''}`);
    console.log(`  Data file: ${DATA_FILE}`);
  });
}

module.exports = { server, store };
