(() => {
  'use strict';
  const C = window.HubCore;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const safeUrl = (u) => (/^https?:\/\//i.test(String(u || '')) ? esc(u) : '#');
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const COLLS = ['trainees', 'quizzes', 'submissions', 'evaluations'];
  const SHEETS = 'Google Sheets';

  // ---------- state ----------
  const store = { trainees: [], quizzes: [], submissions: [], evaluations: [] };
  const caps = { db: undefined, sample: undefined, downloads: undefined, user: undefined, mcp: undefined };
  const loaded = new Set();
  let me = null;
  let readOnly = false;
  let dbError = null;
  const names = new Map();
  const S = {
    tab: lsGet('tab') || 'builder',
    draft: null,
    dirty: false,
    respFilter: { quizId: '', cohort: '', status: '' },
    selectedSub: null,
    evalFilter: 'open',
    dash: { cohort: '', traineeId: '', quizId: '', trainingDay: '', status: '', from: '', to: '' },
    busy: {},
    failed: {},
    evalAllBusy: false,
    reviewDraft: null,
  };

  function lsGet(k) { try { return localStorage.getItem('tah.' + k); } catch (_) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem('tah.' + k, v); } catch (_) { /* storage unavailable */ } }

  // ---------- formatting ----------
  const fmtDate = (iso) => {
    if (!iso) return '';
    const d = new Date(iso.length === 10 ? iso + 'T00:00:00' : iso);
    return isNaN(d) ? esc(iso) : d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
  };
  const fmtDT = (iso) => {
    if (!iso) return '';
    const d = new Date(iso);
    return isNaN(d) ? esc(iso) : d.toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
  };
  const pctTxt = (p) => (p == null ? '—' : p + '%');
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;

  // ---------- toast and dialogs (alert/confirm/prompt don't run in artifacts) ----------
  let toastT;
  function toast(msg, bad) {
    $('toast-root').innerHTML = `<div class="toast ${bad ? 'bad' : ''}" role="status">${esc(msg)}</div>`;
    clearTimeout(toastT);
    toastT = setTimeout(() => ($('toast-root').innerHTML = ''), bad ? 6500 : 3000);
  }

  let modalResolve = null;
  function openModal(title, body, foot, opts) {
    opts = opts || {};
    $('modal-root').innerHTML = `<div class="backdrop" data-act="modal-bg"><div class="modal ${opts.narrow ? 'narrow' : ''}" role="dialog" aria-modal="true" aria-labelledby="m-title">
      <div class="modal-head"><h2 id="m-title">${esc(title)}</h2><button class="icon" type="button" data-act="close-modal" aria-label="Close">✕</button></div>
      <div class="modal-body">${body}</div>${foot ? `<div class="modal-foot">${foot}</div>` : ''}</div></div>`;
    const f = $('modal-root').querySelector('[autofocus]');
    if (f) f.focus();
  }
  function closeModal(value) {
    $('modal-root').innerHTML = '';
    if (modalResolve) {
      const r = modalResolve;
      modalResolve = null;
      r(value === undefined ? null : value);
    }
    if (pendingRender) flushRender();
  }
  function ask(o) {
    if (modalResolve) closeModal(null);
    return new Promise((resolve) => {
      openModal(
        o.title,
        `<p style="margin:0">${o.body || ''}</p>${o.input ? `<label class="field">${esc(o.input)}<input type="text" id="ask-input" autofocus placeholder="${esc(o.placeholder || '')}"></label>` : ''}`,
        `<button class="btn" type="button" data-act="ask-cancel">Cancel</button><button class="btn ${o.danger ? 'danger' : 'primary'}" type="button" data-act="ask-ok">${esc(o.ok || 'Confirm')}</button>`,
        { narrow: true }
      );
      modalResolve = resolve;
    });
  }

  // ---------- data layer ----------
  function dbView() {
    // Overlay local run state so statuses show "Pending Evaluation" while grading runs or after a failure.
    return {
      trainees: store.trainees,
      quizzes: store.quizzes,
      evaluations: store.evaluations,
      submissions: store.submissions.map((s) => (S.busy[s.id] || S.failed[s.id] ? { ...s, evaluationState: S.busy[s.id] ? 'running' : 'failed' } : s)),
    };
  }
  const byId = (coll, id) => store[coll].find((x) => x.id === id);
  const currentEval = (subId) => store.evaluations.find((e) => e.submissionId === subId && !e.supersededAt);

  function upsertLocal(coll, obj) {
    const i = store[coll].findIndex((x) => x.id === obj.id);
    if (i >= 0) store[coll][i] = obj;
    else store[coll].push(obj);
  }
  async function put(coll, obj) {
    if (readOnly) throw new Error('You can view these records but not change them. Ask the page owner for Contributor access.');
    const clean = clone(obj);
    if (caps.db) {
      try {
        await caps.db.collection(coll).doc(clean.id).set(clean);
      } catch (e) {
        if (e && e.code === 'invalid_argument' && !/256|size|deep/i.test(e.message || '')) {
          readOnly = true;
          renderStatus();
          throw new Error('You can view these records but not change them. Ask the page owner for Contributor access.');
        }
        if (e && e.code === 'quota_exceeded') throw new Error('The shared database is full. Delete old staging data in Setup, then try again.');
        throw new Error(e && e.message ? e.message : 'Saving failed. Try again.');
      }
    }
    upsertLocal(coll, clean);
    return clean;
  }
  async function del(coll, id) {
    if (caps.db) await caps.db.collection(coll).doc(id).delete();
    store[coll] = store[coll].filter((x) => x.id !== id);
  }

  async function sha256(text) {
    try {
      const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
      return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
    } catch (_) {
      return '';
    }
  }

  // ---------- render scheduling: never re-render under someone's cursor ----------
  let pendingRender = false;
  const typing = () => {
    const a = document.activeElement;
    return a && a.matches('input[type=text],input[type=url],input[type=email],input[type=number],input[type=date],textarea') && a.closest('#app,#modal-root');
  };
  function scheduleRender() {
    if (typing() || $('modal-root').innerHTML) {
      pendingRender = true;
      renderStatus();
      renderTabs();
      return;
    }
    render();
  }
  function flushRender() {
    if (!typing() && !$('modal-root').innerHTML) {
      pendingRender = false;
      render();
    }
  }
  document.addEventListener('focusout', () => pendingRender && setTimeout(flushRender, 0));

  // ---------- names (store ids only; resolve at render) ----------
  async function resolveNames(ids) {
    const need = [...new Set(ids.filter((x) => x && !names.has(x)))];
    if (!need.length || !caps.user) return;
    need.forEach((x) => names.set(x, ''));
    try {
      const ps = await caps.user.profiles(need);
      need.forEach((x) => names.set(x, (ps[x] && ps[x].name) || ''));
      scheduleRender();
    } catch (_) { /* keep blanks */ }
  }
  const nameOf = (id) => (id && names.get(id)) || 'a trainer';

  // =====================================================================
  // RENDER: status, tabs, router
  // =====================================================================
  function renderStatus() {
    const pills = [];
    pills.push('<span class="pill warn">Staging</span>');
    if (caps.db === undefined) pills.push('<span class="pill busy">Connecting</span>');
    else if (!caps.db) pills.push('<span class="pill bad">Not saved</span>');
    else if (dbError) pills.push('<span class="pill bad">Database error</span>');
    else if (COLLS.every((c) => loaded.has(c))) pills.push(`<span class="pill ok">Shared records ${readOnly ? '(view only)' : 'up to date'}</span>`);
    else pills.push('<span class="pill busy">Loading records</span>');
    pills.push(caps.sample ? '<span class="pill ok">Claude grading ready</span>' : caps.sample === undefined ? '<span class="pill busy">Checking Claude</span>' : '<span class="pill bad">Claude grading unavailable</span>');
    let msg;
    if (caps.db === null) msg = 'This view cannot reach the shared database, so changes last only until you reload. Open the page in claude.ai to save records.';
    else if (dbError) msg = `The shared database stopped responding (${esc(dbError.code)}). Reload the page.`;
    else if (readOnly) msg = 'You can view these records but not change them. Ask the page owner for Contributor access.';
    else msg = `${plural(store.quizzes.length, 'quiz')} · ${plural(store.submissions.length, 'submission')} · ${plural(store.evaluations.filter((e) => !e.supersededAt).length, 'evaluation')}. Nothing on this page sends messages or publishes forms.`;
    $('status').innerHTML = pills.join('') + `<span class="sync-msg">${msg}</span>`;
  }

  function renderTabs() {
    const rows = C.assignmentRows(dbView());
    const review = rows.filter((r) => ['Submitted', 'Pending Evaluation', 'Needs Trainer Review'].includes(r.status)).length;
    const tabs = [
      ['builder', 'Quiz Builder', store.quizzes.length],
      ['responses', 'Quiz Responses', store.submissions.length],
      ['evaluation', 'Evaluation & Scoring', review],
      ['dashboard', 'Performance Dashboard', ''],
    ];
    $('tabs').innerHTML = tabs.map(([k, l, n]) => `<button class="tab" role="tab" type="button" data-tab="${k}" aria-selected="${S.tab === k}">${esc(l)}${n !== '' ? `<span class="n">${n}</span>` : ''}</button>`).join('');
  }

  function render() {
    pendingRender = false;
    renderStatus();
    renderTabs();
    const app = $('app');
    if (caps.db && !COLLS.every((c) => loaded.has(c)) && !dbError) {
      app.innerHTML = `<div class="panel"><div class="empty">Loading quizzes, submissions and evaluations…</div></div>`;
      return;
    }
    const y = window.scrollY;
    const open = [...app.querySelectorAll('details[open][data-k]')].map((d) => d.dataset.k);
    if (S.tab === 'builder') app.innerHTML = renderBuilder();
    else if (S.tab === 'responses') app.innerHTML = renderResponses();
    else if (S.tab === 'evaluation') app.innerHTML = renderEvaluation();
    else app.innerHTML = renderDashboard();
    open.forEach((k) => {
      const d = app.querySelector(`details[data-k="${CSS.escape(k)}"]`);
      if (d) d.open = true;
    });
    window.scrollTo(0, y);
    resolveNames([
      ...store.evaluations.flatMap((e) => [e.trainer && e.trainer.approvedBy, e.trainer && e.trainer.updatedBy, ...(e.history || []).map((h) => h.by)]),
      ...store.quizzes.map((q) => q.form && q.form.approvedBy),
    ]);
  }

  const cohorts = () => [...new Set([...store.trainees.map((t) => t.cohort), ...store.quizzes.map((q) => q.cohort)].filter(Boolean))].sort();
  function statusTag(s) {
    const m = { 'Awaiting Submission': 'plain', Submitted: '', 'Pending Evaluation': 'warn', 'Needs Trainer Review': 'warn', Evaluated: 'ok', Draft: 'plain', Ready: '', Sent: 'ok', Closed: 'plain' };
    return `<span class="tag ${m[s] || ''}">${esc(s)}</span>`;
  }
  function resultTag(r) {
    if (r === 'Pass') return '<span class="tag ok">Pass</span>';
    if (r === 'Fail') return '<span class="tag bad">Fail</span>';
    if (r === 'Needs retake') return '<span class="tag warn">Needs retake</span>';
    return '<span class="tag plain">No pass mark set</span>';
  }

  // =====================================================================
  // TAB 1: QUIZ BUILDER
  // =====================================================================
  function startDraft(q) {
    S.draft = q
      ? clone(q)
      : { id: null, title: '', cohort: cohorts()[0] || '', trainingDay: '', assessmentType: 'Ticket handling', assignedTrainees: [], deadline: '', passingThreshold: '', status: 'Draft', tickets: [C.newTicket()], form: null };
    if (S.draft.passingThreshold == null) S.draft.passingThreshold = '';
    S.dirty = false;
  }

  function renderBuilder() {
    const qs = store.quizzes.slice().sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
    if (!S.draft && qs.length) startDraft(qs[0]);
    if (S.draft && S.draft.id && !S.dirty) {
      const live = byId('quizzes', S.draft.id);
      if (live && live.updatedAt !== S.draft.updatedAt) startDraft(live);
      if (!live) S.draft = qs.length ? clone(qs[0]) : null;
    }
    const list = qs.length
      ? qs.map((q) => `<li><button class="li" type="button" data-act="pick-quiz" data-id="${q.id}" aria-current="${S.draft && S.draft.id === q.id}">
          <span class="t">${esc(q.title || 'Untitled quiz')}</span><span class="s">${esc(q.cohort)} · ${plural(q.tickets.length, 'ticket')}${q.trainingDay ? ' · ' + esc(q.trainingDay) : ''}</span>
          <span class="tags">${statusTag(q.status)}${q.form ? '<span class="tag">Form approved</span>' : ''}${q.form && q.form.google && q.form.google.formUrl ? '<span class="tag ok">Google Form linked</span>' : ''}</span></button></li>`).join('')
      : `<li class="empty">No quizzes yet. Create one, or load the fictional demo in Setup.</li>`;
    return `<div class="split">
      <aside class="side">
        <section class="panel"><div class="panel-head"><h2>Quizzes</h2><button class="btn primary small" type="button" data-act="new-quiz">New quiz</button></div><ul class="list">${list}</ul></section>
        ${renderRoster()}
      </aside>
      <div class="stack">${S.draft ? renderEditor() : '<section class="panel"><div class="empty">Select a quiz or create a new one.</div></section>'}${renderRegistry()}</div>
    </div>`;
  }

  function renderRoster() {
    const rows = store.trainees
      .slice()
      .sort((a, b) => (a.cohort + a.name).localeCompare(b.cohort + b.name))
      .map((t) => `<tr><td>${esc(t.name)}<div class="small muted">${esc(t.email)}</div></td><td class="small">${esc(t.cohort)}</td><td><button class="icon" type="button" data-act="del-trainee" data-id="${t.id}" aria-label="Remove ${esc(t.name)}">✕</button></td></tr>`)
      .join('');
    return `<details class="block" data-k="roster"><summary>Trainee roster <span class="count">${store.trainees.length}</span></summary><div class="body">
      <p class="small muted" style="margin:0">Staging uses fictional trainees only.</p>
      <div class="tw"><table><tbody>${rows || '<tr><td class="muted small">No trainees yet.</td></tr>'}</tbody></table></div>
      <div class="grid"><input type="text" id="tr-name" placeholder="Full name" aria-label="Trainee name"><input type="email" id="tr-email" placeholder="Email" aria-label="Trainee email">
      <input type="text" id="tr-cohort" placeholder="Cohort" list="cohort-dl" aria-label="Cohort"><button class="btn small" type="button" data-act="add-trainee">Add trainee</button></div></div></details>`;
  }

  function renderAssign(d) {
    const mine = store.trainees.filter((t) => !d.cohort || t.cohort === d.cohort);
    const others = store.trainees.filter((t) => d.cohort && t.cohort !== d.cohort);
    const check = (t) => `<label><input type="checkbox" data-act="toggle-trainee" data-id="${t.id}" ${d.assignedTrainees.includes(t.id) ? 'checked' : ''}> ${esc(t.name)}${t.cohort !== d.cohort ? ` <span class="muted small">(${esc(t.cohort)})</span>` : ''}</label>`;
    return `<div class="actions" style="justify-content:space-between"><span class="field">Assigned trainees (${d.assignedTrainees.length})</span>
      <span class="actions"><button class="btn small" type="button" data-act="assign-cohort">Assign whole cohort</button><button class="btn small ghost" type="button" data-act="clear-assign">Clear</button></span></div>
      <div class="checks">${mine.map(check).join('') || '<span class="muted small">No trainees in this cohort yet. Add them in the roster.</span>'}</div>
      ${others.length ? `<details class="block" data-k="other-cohorts"><summary class="small">Trainees from other cohorts (${others.length})</summary><div class="body"><div class="checks">${others.map(check).join('')}</div></div></details>` : ''}`;
  }

  function renderEditor() {
    const d = S.draft;
    const v = C.validateQuiz(d);
    return `<section class="panel">
      <div class="panel-head"><div><h2>${esc(d.title || 'New quiz')}</h2><p class="small muted">${d.id ? 'Last saved ' + fmtDT(d.updatedAt) : 'Not saved yet'}${S.dirty ? ' · <b style="color:var(--warn)">Unsaved changes</b>' : ''}</p></div>
        <div class="actions">${d.id ? '<button class="btn ghost small" type="button" data-act="delete-quiz">Delete</button>' : ''}
          <button class="btn" type="button" data-act="save-draft">Save draft</button><button class="btn primary" type="button" data-act="preview">Preview and approve form</button></div></div>
      ${d.form ? renderFormPanel(d) : ''}
      <h4>A · Quiz details</h4>
      <div class="grid g3">
        <label class="field span2">Quiz title<input type="text" id="q-title" data-bind="title" value="${esc(d.title)}" placeholder="Day 3 Ticket Assessment"></label>
        <label class="field">Quiz status<select id="q-status" data-act="set-status">${C.QUIZ_STATUSES.map((s) => `<option ${d.status === s ? 'selected' : ''} ${!d.form && (s === 'Ready' || s === 'Sent') ? 'disabled' : ''}>${s}</option>`).join('')}</select></label>
        <label class="field">Cohort<input type="text" id="q-cohort" data-bind="cohort" data-assign="1" value="${esc(d.cohort)}" list="cohort-dl"></label>
        <label class="field">Training day<input type="text" id="q-day" data-bind="trainingDay" value="${esc(d.trainingDay)}" placeholder="Day 3"></label>
        <label class="field">Assessment type<select id="q-type" data-bind="assessmentType">${['Ticket handling', 'Knowledge check', 'Final assessment', 'Practice'].map((x) => `<option ${d.assessmentType === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
        <label class="field">Submission deadline<input type="date" id="q-deadline" data-bind="deadline" value="${esc(d.deadline)}"></label>
        <label class="field">Passing threshold (%) <span class="opt">optional</span><input type="number" id="q-pass" min="1" max="100" data-bind="passingThreshold" data-type="number" value="${esc(d.passingThreshold)}" placeholder="Blank means no pass or fail"></label>
        <label class="field">Number of tickets<input type="text" value="${d.tickets.length}" disabled></label>
      </div>
      <datalist id="cohort-dl">${cohorts().map((c) => `<option value="${esc(c)}">`).join('')}</datalist>
      <div id="assign-box" class="grid" style="gap:8px">${renderAssign(d)}</div>
    </section>
    <section class="panel">
      <div class="panel-head"><div><h4>B · Tickets</h4><p class="small muted">Each ticket gets its own form section, its own rubric and its own reference material.</p></div><button class="btn primary small" type="button" data-act="add-ticket">Add ticket</button></div>
      ${d.tickets.map(renderTicket).join('')}
      <details class="block" data-k="questions"><summary>C · Standard trainee questions <span class="muted small" style="font-weight:400">same four for every ticket</span></summary><div class="body">
        <ol style="margin:0;padding-left:20px">${C.QUESTIONS.map((q) => `<li>${esc(q.label)} <span class="muted small">(long answer)</span></li>`).join('')}</ol>
        <p class="small muted" style="margin:0">The form also collects the trainee's name and email, and records the cohort, quiz ID and submission time.</p></div></details>
    </section>
    ${renderValidation(v)}`;
  }

  function renderValidation(v) {
    return (v.errors.length ? `<div class="notice bad"><b>Fix before approving the form:</b><ul>${v.errors.map((e) => `<li>${esc(e)}</li>`).join('')}</ul></div>` : '') + (v.warnings.length ? `<div class="notice warn"><b>Check these:</b><ul>${v.warnings.map((e) => `<li>${esc(e)}</li>`).join('')}</ul></div>` : '');
  }

  function renderFormPanel(d) {
    const f = d.form;
    const g = f.google || {};
    return `<div class="notice ${f.outOfDate ? 'warn' : 'ok'}" style="display:grid;gap:8px">
      <div><b>Form approved</b> ${fmtDT(f.approvedAt)} by ${esc(nameOf(f.approvedBy))}.${f.outOfDate ? ' The tickets changed after approval. Preview and approve again, then regenerate the Google Form.' : ''}</div>
      <div class="actions">
        <button class="btn small" type="button" data-act="open-staging-form" data-id="${d.id}">Open staging form</button>
        <button class="btn small" type="button" data-act="apps-script" data-id="${d.id}">Get Google Form script</button>
        <button class="btn small" type="button" data-act="link-google" data-id="${d.id}">${g.formUrl ? 'Edit Google links' : 'Link Google Form'}</button>
        ${g.formUrl ? `<a class="btn small" href="${safeUrl(g.formUrl)}" target="_blank" rel="noopener">Open Google Form</a><button class="btn small" type="button" data-act="copy" data-text="${esc(g.formUrl)}">Copy link</button><button class="btn small" type="button" data-act="send-dialog" data-id="${d.id}">Send…</button>` : ''}
      </div></div>`;
  }

  function renderTicket(t, i) {
    const d = S.draft;
    const p = `tickets.${i}`;
    const r = t.rubric;
    const sum = C.rubricMax(r);
    const ok = Number(r.totalPoints) > 0 && sum === Number(r.totalPoints);
    const ta = (path, val, ph, rows) => `<textarea id="${t.id}-${path.replace(/\./g, '-')}" data-bind="${p}.${path}" rows="${rows || 3}" placeholder="${esc(ph || '')}">${esc(val)}</textarea>`;
    const crit = r.criteria
      .map((c, ci) => `<tr>
        <td style="min-width:150px"><input type="text" id="${c.id}-name" data-bind="${p}.rubric.criteria.${ci}.name" value="${esc(c.name)}" placeholder="Criterion" aria-label="Criterion name"></td>
        <td style="min-width:150px"><select id="${c.id}-cat" data-bind="${p}.rubric.criteria.${ci}.category" aria-label="Assesses">${C.CRITERION_CATEGORIES.map((k) => `<option value="${k.key}" ${c.category === k.key ? 'selected' : ''}>${esc(k.label)}</option>`).join('')}</select></td>
        <td style="width:84px"><input type="number" id="${c.id}-max" min="0" step="0.5" data-bind="${p}.rubric.criteria.${ci}.maxPoints" data-type="number" data-sum="${i}" value="${esc(c.maxPoints)}" aria-label="Max points"></td>
        <td style="min-width:220px"><textarea id="${c.id}-g" rows="1" data-bind="${p}.rubric.criteria.${ci}.guidance" placeholder="What earns full points" aria-label="Guidance">${esc(c.guidance)}</textarea>
          ${(c.deductions || []).map((dd, di) => `<div class="actions" style="margin-top:4px;flex-wrap:nowrap"><input type="text" id="${c.id}-d${di}" data-bind="${p}.rubric.criteria.${ci}.deductions.${di}.description" value="${esc(dd.description)}" placeholder="Deduction reason" aria-label="Deduction reason"><input type="number" id="${c.id}-dp${di}" min="0" step="0.5" style="width:74px" data-bind="${p}.rubric.criteria.${ci}.deductions.${di}.points" data-type="number" value="${esc(dd.points)}" aria-label="Points deducted"><button class="icon" type="button" data-act="del-deduction" data-t="${i}" data-c="${ci}" data-d="${di}" aria-label="Remove deduction">✕</button></div>`).join('')}
          <button class="linkish small" type="button" data-act="add-deduction" data-t="${i}" data-c="${ci}">+ Point deduction</button></td>
        <td><button class="icon" type="button" data-act="del-criterion" data-t="${i}" data-c="${ci}" aria-label="Remove criterion">✕</button></td></tr>`)
      .join('');
    return `<div class="ticket" id="ticket-${i}">
      <div class="ticket-head"><span class="ticket-name">Ticket ${i + 1}</span><span class="ticket-url">${esc(t.url || 'No ticket link yet')}</span>
        <span class="actions" style="gap:2px"><button class="icon" type="button" data-act="move-ticket" data-t="${i}" data-dir="-1" ${i === 0 ? 'disabled' : ''} aria-label="Move up">↑</button>
        <button class="icon" type="button" data-act="move-ticket" data-t="${i}" data-dir="1" ${i === d.tickets.length - 1 ? 'disabled' : ''} aria-label="Move down">↓</button>
        <button class="btn small ghost" type="button" data-act="dup-ticket" data-t="${i}">Duplicate</button><button class="btn small danger" type="button" data-act="del-ticket" data-t="${i}">Remove</button></span></div>
      <details class="block" data-k="${t.id}-ref" open><summary>Ticket reference <span class="muted small" style="font-weight:400">trainees see this</span></summary><div class="body grid g2">
        <label class="field span2">Ticket link<input type="url" id="${t.id}-url" data-bind="${p}.url" value="${esc(t.url)}" placeholder="https://…"></label>
        <label class="field">Ticket description <span class="opt">optional</span>${ta('description', t.description, 'Short summary shown under the link', 2)}</label>
        <label class="field">Extra context or instructions <span class="opt">optional</span>${ta('context', t.context, 'Anything the trainee should know', 2)}</label>
        <div class="span2 grid" style="gap:6px"><span class="field">Screenshots and attachments <span class="opt">optional, as links</span></span>
          ${(t.attachments || []).map((a, ai) => `<div class="actions small"><a href="${safeUrl(a.url)}" target="_blank" rel="noopener">${esc(a.name || a.url)}</a><button class="icon" type="button" data-act="del-attach" data-t="${i}" data-a="${ai}" aria-label="Remove attachment">✕</button></div>`).join('')}
          <div class="actions" style="flex-wrap:nowrap"><input type="url" id="att-${t.id}" placeholder="Link to a screenshot in Google Drive" aria-label="Attachment link"><button class="btn small" type="button" data-act="add-attach" data-t="${i}">Add</button></div></div>
      </div></details>
      <details class="block" data-k="${t.id}-rubric"><summary>Scoring rubric <span class="tag ${ok ? 'ok' : 'warn'}" id="sum-${i}">${sum} of ${Number(r.totalPoints) || 0} points allocated</span></summary><div class="body">
        <div class="private">Trainer only. Never shown to trainees or written into the form.</div>
        <div class="grid g2">
          <label class="field">Total available points<input type="number" id="${t.id}-total" min="0" step="0.5" data-bind="${p}.rubric.totalPoints" data-type="number" data-sum="${i}" value="${esc(r.totalPoints)}"></label><span></span>
          <label class="field">Expected customer intent${ta('rubric.expectedIntent', r.expectedIntent, 'What C actually wants')}</label>
          <label class="field">Required CRM actions <span class="opt">one per line</span>${ta('rubric.requiredCrmActions', r.requiredCrmActions, 'Add internal note\nIssue partial refund')}</label>
          <label class="field">Required customer resolution${ta('rubric.requiredResolution', r.requiredResolution)}</label>
          <label class="field">Critical errors${ta('rubric.criticalErrors', r.criticalErrors, 'Mistakes Claude must flag separately')}</label>
          <label class="field">Expected communication to C${ta('rubric.expectedC', r.expectedC)}</label>
          <label class="field">Expected communication to CP${ta('rubric.expectedCP', r.expectedCP, 'Write N/A if no CP message is expected')}</label>
          <label class="field">Acceptable alternative responses${ta('rubric.acceptableAlternatives', r.acceptableAlternatives)}</label>
          <label class="field">Automatic fail conditions <span class="opt">optional</span>${ta('rubric.autoFailConditions', r.autoFailConditions)}</label>
        </div>
        <div class="panel-head"><b>Criteria and points</b><span class="actions"><button class="btn small" type="button" data-act="suggest" data-t="${i}">Insert suggested criteria</button><button class="btn small" type="button" data-act="add-criterion" data-t="${i}">Add criterion</button></span></div>
        <div class="tw"><table class="crit"><thead><tr><th>Criterion</th><th>Assesses</th><th>Max</th><th>Guidance and deductions</th><th></th></tr></thead><tbody>${crit || '<tr><td colspan="5" class="muted small">No criteria yet. Insert the suggested five or add your own.</td></tr>'}</tbody></table></div>
      </div></details>
      <details class="block" data-k="${t.id}-refs"><summary>Reference material <span class="muted small" style="font-weight:400">for Claude's grading only</span></summary><div class="body">
        <div class="private">Trainer only. Claude grades with it; trainees never see it.</div>
        <div class="grid g2">
          <label class="field">Policy context${ta('references.policyContext', t.references.policyContext)}</label>
          <label class="field">Knowledge Base references${ta('references.kbReferences', t.references.kbReferences)}</label>
          <label class="field">Expected resolution${ta('references.expectedResolution', t.references.expectedResolution)}</label>
          <label class="field">Extra instructions for grading${ta('references.evaluatorInstructions', t.references.evaluatorInstructions)}</label>
          <label class="field">Sample response to C${ta('references.sampleC', t.references.sampleC)}</label>
          <label class="field">Sample message to CP${ta('references.sampleCP', t.references.sampleCP)}</label>
        </div></div></details>
    </div>`;
  }

  function renderRegistry() {
    const rows = store.quizzes
      .slice()
      .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
      .map((q) => `<tr><td>${esc(q.title)}</td><td>${q.form && q.form.google && q.form.google.formUrl ? `<a href="${safeUrl(q.form.google.formUrl)}" target="_blank" rel="noopener">Google Form</a>` : q.form ? 'Staging form' : '<span class="muted">Not approved</span>'}</td><td>${esc(q.cohort)}</td><td class="small">${q.assignedTrainees.map((id) => esc((byId('trainees', id) || {}).name || 'Removed trainee')).join(', ')}</td><td>${fmtDate(q.createdAt)}</td><td>${fmtDate(q.deadline)}</td><td>${statusTag(q.status)}</td></tr>`)
      .join('');
    return `<details class="block" data-k="registry"><summary>Quiz registry <span class="count">${store.quizzes.length}</span></summary><div class="body"><div class="tw">
      <table><thead><tr><th>Quiz title</th><th>Form</th><th>Cohort</th><th>Assigned trainees</th><th>Created</th><th>Deadline</th><th>Status</th></tr></thead><tbody>${rows || '<tr><td colspan="7" class="muted">No quizzes yet.</td></tr>'}</tbody></table></div></div></details>`;
  }

  function traineeFormHtml(q, interactive) {
    const tv = C.traineeView(q);
    const roster = q.assignedTrainees.map((id) => byId('trainees', id)).filter(Boolean);
    const ident = interactive
      ? `<div class="gcard"><label class="field">Full name *<input type="text" id="sf-name" list="sf-roster" placeholder="Fictional trainee name"></label><datalist id="sf-roster">${roster.map((t) => `<option value="${esc(t.name)}">`).join('')}</datalist></div>
         <div class="gcard"><label class="field">Email *<input type="email" id="sf-email" placeholder="name@example.com"></label><span class="small muted">Pick a roster name and the email fills in.</span></div>`
      : `<div class="gcard"><b>Full name *</b><span class="muted small">Short answer</span></div><div class="gcard"><b>Email *</b><span class="muted small">Short answer</span></div>`;
    return `<div class="gform">
      <div class="gcard head"><h2 style="font-family:var(--body);font-weight:400;font-size:1.4rem">${esc(tv.title)}</h2>
        <span class="small muted">Cohort: ${esc(tv.cohort)}${tv.trainingDay ? ' · ' + esc(tv.trainingDay) : ''}${tv.assessmentType ? ' · ' + esc(tv.assessmentType) : ''}${tv.deadline ? ' · Due ' + fmtDate(tv.deadline) : ''}<br>Quiz ID: ${esc(tv.id || 'assigned when saved')}</span></div>
      ${ident}
      <div class="gcard"><b>Cohort *</b><span class="small">${esc(tv.cohort)}</span></div>
      ${tv.tickets.map((t) => `<span class="gsec">Section ${t.number + 1} of ${tv.tickets.length + 1}</span>
        <div class="gcard head" style="border-top-left-radius:0"><h3 style="font-family:var(--body);font-weight:400;font-size:1.2rem">Ticket ${t.number}</h3>
          <div>Ticket reference: <a href="${safeUrl(t.url)}" target="_blank" rel="noopener">${esc(t.url || '(missing link)')}</a></div>
          ${t.description ? `<div class="small">${esc(t.description)}</div>` : ''}${t.context ? `<div class="small muted">${esc(t.context)}</div>` : ''}
          ${t.attachments.map((a) => `<div class="small">Attachment: <a href="${safeUrl(a.url)}" target="_blank" rel="noopener">${esc(a.name || a.url)}</a></div>`).join('')}</div>
        ${tv.questions.map((qq) => `<div class="gcard">${interactive ? `<label class="field" style="color:var(--ink)">${esc(qq.label)} *<textarea id="sf-${t.id}-${qq.key}" rows="3" placeholder="Your answer"></textarea></label>` : `<b>${esc(qq.label)} *</b><span class="muted small">Long answer</span>`}</div>`).join('')}`).join('')}
    </div>`;
  }

  function openPreview() {
    const d = S.draft;
    const v = C.validateQuiz(d);
    const rubric = d.tickets
      .map((t, i) => `<details class="block"><summary>Ticket ${i + 1} rubric <span class="muted small" style="font-weight:400">${C.rubricMax(t.rubric)} of ${t.rubric.totalPoints} points</span></summary><div class="body">
        <div class="tw"><table><thead><tr><th>Criterion</th><th class="num">Points</th><th>Deductions</th></tr></thead><tbody>${t.rubric.criteria.map((c) => `<tr><td>${esc(c.name)}</td><td class="num">${esc(c.maxPoints)}</td><td class="small">${(c.deductions || []).map((x) => `${esc(x.description)} (−${esc(x.points)})`).join('<br>')}</td></tr>`).join('')}</tbody></table></div>
        <div class="small"><b>Intent:</b> ${esc(t.rubric.expectedIntent) || '—'}<br><b>Resolution:</b> ${esc(t.rubric.requiredResolution) || '—'}<br><b>Critical errors:</b> ${esc(t.rubric.criticalErrors) || '—'}<br><b>Auto-fail:</b> ${esc(t.rubric.autoFailConditions) || '—'}</div></div></details>`)
      .join('');
    openModal(
      'Preview and approve',
      `${renderValidation(v)}
      <div class="grid g2" style="align-items:start">
        <div class="grid" style="gap:8px"><h3>What trainees will see</h3>${traineeFormHtml(d, false)}</div>
        <div class="grid" style="gap:8px"><h3>Trainer-only rubric</h3><div class="private">Not included in the form.</div>${rubric}
          <h3>Ticket links</h3><ol style="margin:0;padding-left:20px">${d.tickets.map((t) => `<li class="small"><a href="${safeUrl(t.url)}" target="_blank" rel="noopener">${esc(t.url || '(missing)')}</a></li>`).join('')}</ol></div>
      </div>`,
      `<label class="small" style="margin-right:auto;display:flex;gap:8px;align-items:center"><input type="checkbox" id="approve-check" ${v.errors.length ? 'disabled' : ''}> I reviewed the links, questions and rubric. Approve this form.</label>
       <button class="btn" type="button" data-act="close-modal">Keep editing</button>
       <button class="btn primary" type="button" data-act="approve-form" ${v.errors.length ? 'disabled' : ''}>Approve form</button>`
    );
  }

  async function saveDraft(silent) {
    const d = S.draft;
    const now = new Date().toISOString();
    const existing = d.id ? byId('quizzes', d.id) : null;
    const q = clone(d);
    if (!q.id) {
      q.id = C.uid('quiz');
      q.createdAt = now;
      q.createdBy = me && me.id;
    }
    q.passingThreshold = q.passingThreshold === '' || q.passingThreshold == null ? null : Number(q.passingThreshold);
    q.updatedAt = now;
    if (existing && existing.form && JSON.stringify(C.traineeView(existing)) !== JSON.stringify(C.traineeView(q))) q.form = { ...q.form, outOfDate: true };
    const saved = await put('quizzes', q);
    startDraft(saved);
    if (!silent) toast('Draft saved.');
    render();
    return saved;
  }

  // =====================================================================
  // Google Form: Apps Script the trainer runs themselves
  // =====================================================================
  function appsScriptFor(q) {
    const tv = C.traineeView(q);
    const data = JSON.stringify({ title: tv.title, cohort: tv.cohort, trainingDay: tv.trainingDay, assessmentType: tv.assessmentType, deadline: tv.deadline, quizId: tv.id, questions: tv.questions.map((x) => x.label), tickets: tv.tickets.map((t) => ({ number: t.number, url: t.url, description: t.description, context: t.context, attachments: t.attachments })) }, null, 2);
    return `/**
 * Creates the Google Form for "${tv.title.replace(/\*\//g, '')}".
 * Generated by the Ticket Assessment Hub on ${new Date().toISOString().slice(0, 10)}.
 *
 * 1. Go to script.google.com, click New project, paste this file, save.
 * 2. Run createQuizForm and allow access when Google asks.
 * 3. Open View > Logs (or Execution log). Copy the form link and the
 *    response sheet link into the hub ("Link Google Form").
 *
 * The form is created CLOSED and is not shared with anyone. Open it to
 * responses yourself (Responses tab > Accepting responses) when ready.
 * Trainer-only rubrics and expected answers are not in this script.
 */
var QUIZ = ${data};

function createQuizForm() {
  var form = FormApp.create(QUIZ.title);
  form.setDescription(
    'Cohort: ' + QUIZ.cohort +
    (QUIZ.trainingDay ? '\\nTraining day: ' + QUIZ.trainingDay : '') +
    (QUIZ.assessmentType ? '\\nAssessment type: ' + QUIZ.assessmentType : '') +
    (QUIZ.deadline ? '\\nSubmission deadline: ' + QUIZ.deadline : '') +
    '\\nQuiz ID: ' + QUIZ.quizId);
  try { form.setCollectEmail(true); } catch (e) {}
  form.setProgressBar(true);
  form.setAllowResponseEdits(false);
  form.addTextItem().setTitle('Full name').setRequired(true);
  form.addListItem().setTitle('Cohort').setChoiceValues([QUIZ.cohort]).setRequired(true);
  QUIZ.tickets.forEach(function (t) {
    var help = 'Ticket reference: ' + t.url;
    if (t.description) help += '\\n\\n' + t.description;
    if (t.context) help += '\\n\\n' + t.context;
    (t.attachments || []).forEach(function (a) { help += '\\nAttachment: ' + (a.name ? a.name + ' ' : '') + a.url; });
    form.addPageBreakItem().setTitle('Ticket ' + t.number).setHelpText(help);
    QUIZ.questions.forEach(function (label) {
      form.addParagraphTextItem().setTitle(label).setRequired(true);
    });
  });
  var sheet = SpreadsheetApp.create(QUIZ.title + ' (responses)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  form.setAcceptingResponses(false);
  Logger.log('Form link (send to trainees later): ' + form.getPublishedUrl());
  Logger.log('Form editor: ' + form.getEditUrl());
  Logger.log('Response sheet: ' + sheet.getUrl());
}
`;
  }

  // =====================================================================
  // Submissions and imports
  // =====================================================================
  async function recordSubmission(q, input, source, externalId) {
    if (externalId) {
      const dup = store.submissions.find((s) => s.quizId === q.id && s.externalResponseId === externalId);
      if (dup) return { duplicate: true, submission: dup };
    }
    const email = String(input.traineeEmail || '').trim().toLowerCase();
    const name = String(input.traineeName || '').trim();
    if (!name || !email) throw new Error('Each response needs a name and an email.');
    const tr = store.trainees.find((t) => t.email.toLowerCase() === email);
    const map = new Map((input.answers || []).map((a) => [a.ticketId, a]));
    const answers = q.tickets.map((t, i) => {
      const a = map.get(t.id) || {};
      return { ticketId: t.id, ticketNumber: i + 1, ticketUrl: t.url, q1: String(a.q1 || ''), q2: String(a.q2 || ''), q3: String(a.q3 || ''), q4: String(a.q4 || '') };
    });
    const sub = {
      id: C.uid('sub'),
      quizId: q.id,
      traineeId: tr ? tr.id : null,
      traineeName: name,
      traineeEmail: email,
      cohort: input.cohort || q.cohort,
      submittedAt: input.submittedAt || new Date().toISOString(),
      receivedAt: new Date().toISOString(),
      source,
      externalResponseId: externalId || null,
      answers,
      rawHash: await sha256(JSON.stringify(answers)),
      evaluationState: 'none',
    };
    await put('submissions', sub);
    return { duplicate: false, submission: sub };
  }

  function parseCSV(text) {
    const rows = [];
    let row = [];
    let f = '';
    let q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"' && text[i + 1] === '"') { f += '"'; i++; }
        else if (c === '"') q = false;
        else f += c;
      } else if (c === '"') q = true;
      else if (c === ',') { row.push(f); f = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(f); rows.push(row); row = []; f = '';
      } else f += c;
    }
    if (f || row.length) { row.push(f); rows.push(row); }
    return rows.filter((r) => r.some((x) => String(x).trim()));
  }

  // Rows from a Google Form response sheet (header + one row per response).
  function rowsToResponses(q, rows) {
    if (!rows.length) throw new Error('The sheet has no rows.');
    const norm = (s) => String(s || '').replace(/^﻿/, '').replace(/\s+/g, ' ').trim().toLowerCase();
    const h = rows[0].map(norm);
    const ts = h.findIndex((x) => x === 'timestamp');
    const em = h.findIndex((x) => /email/.test(x));
    const nm = h.findIndex((x) => x === 'full name' || x === 'name');
    const co = h.findIndex((x) => x === 'cohort');
    const first = h.indexOf(norm(C.QUESTIONS[0].label));
    if (first < 0 || em < 0 || nm < 0) throw new Error('Could not find the name, email and question columns. Use the response sheet of a form made with this hub\'s script.');
    const need = first + q.tickets.length * 4;
    if (h.length < need) throw new Error(`The sheet has question columns for ${Math.floor((h.length - first) / 4)} tickets, but this quiz has ${q.tickets.length}.`);
    return rows.slice(1).map((r) => {
      const raw = ts >= 0 ? String(r[ts] || '') : '';
      const d = raw ? new Date(raw) : null;
      return {
        externalId: `${raw}|${String(r[em] || '').trim().toLowerCase()}`,
        input: {
          traineeName: r[nm],
          traineeEmail: r[em],
          cohort: co >= 0 ? r[co] : q.cohort,
          submittedAt: d && !isNaN(d) ? d.toISOString() : new Date().toISOString(),
          answers: q.tickets.map((t, i) => {
            const b = first + i * 4;
            return { ticketId: t.id, q1: r[b] || '', q2: r[b + 1] || '', q3: r[b + 2] || '', q4: r[b + 3] || '' };
          }),
        },
      };
    });
  }

  async function importRows(q, rows, source) {
    const parsed = rowsToResponses(q, rows);
    let added = 0;
    let dup = 0;
    for (const p of parsed) {
      const out = await recordSubmission(q, p.input, source, p.externalId);
      out.duplicate ? dup++ : added++;
    }
    return { added, dup, total: parsed.length };
  }

  function findValues(p) {
    if (!p || typeof p !== 'object') return null;
    if (Array.isArray(p.values)) return p.values;
    if (Array.isArray(p) && p.every(Array.isArray)) return p;
    for (const v of Object.values(p)) {
      const got = findValues(v);
      if (got) return got;
    }
    return null;
  }

  async function importFromSheet(q, sheetRef, tab) {
    if (!caps.mcp) throw new Error('Google Sheets is not available in this view. Import a CSV instead.');
    const id = (/\/d\/([a-zA-Z0-9_-]{20,})/.exec(sheetRef) || [])[1] || sheetRef.trim();
    if (!/^[a-zA-Z0-9_-]{20,}$/.test(id)) throw new Error('Paste the response sheet link or its ID.');
    let res;
    try {
      res = await caps.mcp.callTool(SHEETS, 'get_values', { spreadsheetId: id, range: `'${(tab || 'Form Responses 1').replace(/'/g, "''")}'!A1:ZZ3000` }, { cache: false });
    } catch (e) {
      const code = e && e.code;
      if (code === 'server_not_connected') throw new Error('Google Sheets isn\'t connected. Add it in claude.ai Settings › Connectors, or import a CSV.');
      if (code === 'needs_reauth') throw new Error('Your Google Sheets connection has expired. Reconnect it in claude.ai Settings › Connectors.');
      if (code === 'not_in_manifest' || code === 'not_granted') throw new Error('Google Sheets is turned off for this page. Allow it in the page\'s permissions.');
      throw new Error('Google Sheets could not read that sheet: ' + (e && e.message ? String(e.message).slice(0, 160) : 'unknown error'));
    }
    let p = res && res.payload;
    if (typeof p === 'string') { try { p = JSON.parse(p); } catch (_) { /* keep */ } }
    const values = findValues(p);
    if (!values) throw new Error('Google Sheets answered, but no rows were found in that tab. Check the tab name.');
    return importRows(q, values, 'google-form');
  }

  // =====================================================================
  // Evaluation with Claude
  // =====================================================================
  function schemaText() {
    return JSON.stringify(C.TICKET_EVAL_SCHEMA);
  }

  async function gradeTicket(t, n, ans) {
    const prompt = `${C.EVALUATOR_SYSTEM_PROMPT}\n\n${C.buildTicketPrompt(t, n, ans)}\n\nReply with only one JSON object that matches this JSON Schema. Use the exact criterionId values given above.\n${schemaText()}`;
    try {
      return await caps.sample.json(prompt, { modelTier: 'default', cache: false });
    } catch (e) {
      const code = e && e.code;
      const msgs = {
        not_granted: 'Claude grading is not allowed for this page. Allow it in the page\'s permissions, then try again.',
        sampling_disabled: 'Claude is not available for this account.',
        rate_limited: 'Too many grading requests right now, or your Claude usage limit was reached. Try again in a few minutes.',
        refused: `Claude declined to grade Ticket ${n}. Score it by hand.`,
        invalid_json: `Claude's grade for Ticket ${n} could not be read. Try again.`,
        prompt_too_large: `Ticket ${n}'s rubric and answers are too long to grade at once. Shorten the reference material.`,
        session_expired: 'Your claude.ai session expired. Sign in again.',
      };
      const err = new Error(msgs[code] || `Grading Ticket ${n} failed (${code || 'error'}). Try again.`);
      err.code = code;
      throw err;
    }
  }

  async function evaluate(subId, reevaluate, reason) {
    if (!caps.sample) throw new Error('Claude grading is not available in this view. Open the page in claude.ai.');
    const sub = byId('submissions', subId);
    const q = byId('quizzes', sub.quizId);
    if (!q) throw new Error('The quiz for this submission was deleted.');
    const cur = currentEval(subId);
    if (cur && !reevaluate) throw new Error('This submission already has an evaluation. Use Reevaluate to make a new version.');
    if (S.busy[subId]) throw new Error('Grading is already running for this submission.');
    S.busy[subId] = true;
    delete S.failed[subId];
    scheduleRender();
    try {
      const ansBy = new Map(sub.answers.map((a) => [a.ticketId, a]));
      const settled = await Promise.allSettled(q.tickets.map((t, i) => gradeTicket(t, i + 1, ansBy.get(t.id) || {})));
      const bad = settled.find((r) => r.status === 'rejected');
      if (bad) throw bad.reason;
      const tickets = settled.map((r, i) => C.normalizeTicketEval(q.tickets[i], r.value || {}, i + 1));
      const now = new Date().toISOString();
      const prior = store.evaluations.filter((e) => e.submissionId === subId);
      const ev = {
        id: C.uid('eval'),
        submissionId: subId,
        quizId: q.id,
        version: prior.length + 1,
        createdAt: now,
        createdBy: me && me.id,
        ai: { evaluator: 'claude', model: 'Claude (default tier, via claude.ai)', tickets, overall: C.summarize(tickets, q), rubricSnapshot: q.tickets.map((t) => ({ ticketId: t.id, rubric: t.rubric, references: t.references })) },
        trainer: null,
        history: [],
      };
      await put('evaluations', ev);
      if (cur) await put('evaluations', { ...clone(cur), supersededAt: now, supersededReason: reason || 'Trainer requested reevaluation' });
      return ev;
    } catch (e) {
      S.failed[subId] = e.message;
      throw e;
    } finally {
      delete S.busy[subId];
      S.reviewDraft = null;
      scheduleRender();
    }
  }

  async function saveReview(approve) {
    const rd = S.reviewDraft;
    const ev = clone(byId('evaluations', rd.evalId));
    const q = byId('quizzes', ev.quizId);
    const now = new Date().toISOString();
    const before = ev.trainer ? { earned: ev.trainer.overall.earned, approvedAt: ev.trainer.approvedAt || null } : null;
    const tr = ev.trainer || { tickets: clone(ev.ai.tickets), createdAt: now };
    rd.tickets.forEach((t, ti) => {
      const tt = tr.tickets[ti];
      t.criteria.forEach((c, ci) => {
        const tc = tt.criteria[ci];
        if (tc.awarded !== Number(c.awarded) || tc.reason !== c.reason) tc.trainerEdited = true;
        tc.awarded = Number(c.awarded) || 0;
        tc.reason = String(c.reason || '');
      });
      tt.improvements = t.improvements.filter((i) => i.detail);
      tt.trainerNotes = t.trainerNotes || '';
      C.recomputeTicket(tt);
    });
    tr.comments = rd.comments;
    tr.overrideResult = rd.overrideResult;
    tr.overall = C.summarize(tr.tickets, q);
    tr.updatedAt = now;
    tr.updatedBy = me && me.id;
    if (approve) {
      tr.approvedAt = now;
      tr.approvedBy = me && me.id;
      tr.approvedByName = me && me.id ? '' : rd.reviewer;
    }
    ev.trainer = tr;
    ev.history = [...(ev.history || []), { at: now, by: me && me.id, action: approve ? 'approved' : 'edited', before, after: { earned: tr.overall.earned } }];
    await put('evaluations', ev);
    S.reviewDraft = null;
    toast(approve ? 'Evaluation approved as final.' : 'Edits saved.');
    render();
  }

  // =====================================================================
  // TAB 2: QUIZ RESPONSES
  // =====================================================================
  function renderResponses() {
    const f = S.respFilter;
    const rows = C.assignmentRows(dbView()).filter((r) => (!f.quizId || r.quizId === f.quizId) && (!f.cohort || r.cohort === f.cohort) && (!f.status || r.status === f.status));
    const approved = store.quizzes.filter((q) => q.form);
    const body = rows
      .map((r) => `<tr ${r.submissionId ? `class="click" data-act="open-sub" data-id="${r.submissionId}"` : ''}>
        <td>${r.submissionId ? `<button class="linkish" type="button" data-act="open-sub" data-id="${r.submissionId}">${esc(r.traineeName)}</button>` : esc(r.traineeName)}${r.unassigned ? ' <span class="tag warn">Not assigned</span>' : ''}<div class="small muted">${esc(r.traineeEmail)}</div></td>
        <td>${esc(r.cohort)}</td><td>${esc(r.quizTitle)}</td><td>${r.submittedAt ? fmtDT(r.submittedAt) : '<span class="muted">—</span>'}</td>
        <td class="num">${r.submissionId ? `${r.ticketsAnswered} of ${r.ticketsTotal}` : '—'}</td><td>${statusTag(r.status)}</td></tr>`)
      .join('');
    return `<section class="panel">
      <div class="panel-head"><div><h2>Quiz responses</h2><p class="small muted">Answers are stored exactly as received. Grading and review never change them.</p></div>
        <div class="actions"><button class="btn small" type="button" data-act="import" ${approved.length ? '' : 'disabled'}>Import responses</button>
        <button class="btn small primary" type="button" data-act="test-response" ${approved.length ? '' : 'disabled'}>Submit a test response</button></div></div>
      ${approved.length ? '' : '<div class="notice info">Approve a quiz form in Quiz Builder first. Then submit test responses here, or import real ones from the Google Form\'s response sheet.</div>'}
      <div class="grid g3">
        <select id="rf-quiz" data-filter="respFilter.quizId" aria-label="Quiz"><option value="">All quizzes</option>${store.quizzes.map((q) => `<option value="${q.id}" ${f.quizId === q.id ? 'selected' : ''}>${esc(q.title)}</option>`).join('')}</select>
        <select id="rf-cohort" data-filter="respFilter.cohort" aria-label="Cohort"><option value="">All cohorts</option>${cohorts().map((c) => `<option ${f.cohort === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>
        <select id="rf-status" data-filter="respFilter.status" aria-label="Status"><option value="">All statuses</option>${C.EVAL_STATUSES.map((s) => `<option ${f.status === s ? 'selected' : ''}>${s}</option>`).join('')}</select>
      </div>
      <div class="tw"><table><thead><tr><th>Trainee</th><th>Cohort</th><th>Quiz</th><th>Submitted</th><th class="num">Tickets answered</th><th>Evaluation status</th></tr></thead>
      <tbody>${body || '<tr><td colspan="6"><div class="empty">No responses match these filters.</div></td></tr>'}</tbody></table></div></section>`;
  }

  function answersHtml(sub) {
    return sub.answers
      .map((a) => `<details class="block" open><summary>Ticket ${a.ticketNumber} <a class="small" href="${safeUrl(a.ticketUrl)}" target="_blank" rel="noopener" style="font-weight:400">${esc(a.ticketUrl)}</a></summary><div class="body" style="gap:6px">
        ${C.QUESTIONS.map((qq) => `<div class="qlabel">${esc(qq.label)}</div><div class="answer">${esc(a[qq.key]) || '<span class="muted">(blank)</span>'}</div>`).join('')}</div></details>`)
      .join('');
  }

  function openSubmission(id) {
    const s = byId('submissions', id);
    if (!s) return;
    const q = byId('quizzes', s.quizId) || {};
    openModal(
      `${s.traineeName} · ${q.title || 'Deleted quiz'}`,
      `<div class="small" style="display:grid;gap:2px"><span><b>Email</b> ${esc(s.traineeEmail)}</span><span><b>Cohort</b> ${esc(s.cohort)}</span><span><b>Submitted</b> ${fmtDT(s.submittedAt)}</span><span><b>Source</b> ${esc(s.source === 'google-form' ? 'Google Form import' : s.source === 'staging-test' ? 'Staging test response' : s.source)}</span><span class="muted">Integrity check ${esc((s.rawHash || '').slice(0, 12))}</span></div>
      ${S.failed[s.id] ? `<div class="notice bad">Last grading attempt failed: ${esc(S.failed[s.id])}</div>` : ''}${answersHtml(s)}`,
      `<button class="btn" type="button" data-act="close-modal">Close</button><button class="btn primary" type="button" data-act="go-eval" data-id="${s.id}">${currentEval(s.id) ? 'Open evaluation' : 'Grade with Claude'}</button>`
    );
  }

  function openStagingForm(qid) {
    const approved = store.quizzes.filter((q) => q.form);
    const q = byId('quizzes', qid) || approved[0];
    openModal(
      'Staging form',
      `<div class="notice warn">Test form. Use fictional answers only. Nothing is sent anywhere; the response is saved here as a staging test.</div>
      ${approved.length > 1 ? `<label class="field">Quiz<select id="sf-quiz" data-act="sf-quiz">${approved.map((x) => `<option value="${x.id}" ${x.id === q.id ? 'selected' : ''}>${esc(x.title)}</option>`).join('')}</select></label>` : ''}
      <div id="sf-body" data-quiz="${q.id}">${traineeFormHtml(q, true)}</div>`,
      `<span class="small muted" style="margin-right:auto">Every question is required. Type N/A where it does not apply.</span><button class="btn" type="button" data-act="close-modal">Cancel</button><button class="btn primary" type="button" data-act="submit-staging">Submit test response</button>`
    );
  }

  function openImport() {
    const approved = store.quizzes.filter((q) => q.form);
    const sheetOk = !!caps.mcp;
    openModal(
      'Import responses',
      `<p class="small" style="margin:0">Import from the Google Sheet that the quiz's Google Form writes to. Responses already imported are skipped, so importing again is safe.</p>
      <label class="field">Quiz<select id="im-quiz">${approved.map((q) => `<option value="${q.id}">${esc(q.title)}</option>`).join('')}</select></label>
      <details class="block" open><summary>From Google Sheets <span class="muted small" style="font-weight:400">read only, through your connector</span></summary><div class="body">
        ${sheetOk ? '' : '<div class="notice warn">Google Sheets is not available in this view. Use a CSV instead.</div>'}
        <label class="field">Response sheet link or ID<input type="text" id="im-sheet" placeholder="https://docs.google.com/spreadsheets/d/…"></label>
        <label class="field">Tab name<input type="text" id="im-tab" value="Form Responses 1"></label>
        <div class="actions"><button class="btn primary small" type="button" data-act="import-sheet" ${sheetOk ? '' : 'disabled'}>Import from sheet</button></div></div></details>
      <details class="block"><summary>From a CSV file</summary><div class="body">
        <p class="small muted" style="margin:0">In the response sheet, choose File › Download › Comma separated values, then pick the file here.</p>
        <input type="file" id="im-file" accept=".csv,text/csv" data-act="import-csv" aria-label="CSV file"></div></details>
      <div id="im-status" class="small" role="status"></div>`,
      `<button class="btn" type="button" data-act="close-modal">Close</button>`
    );
  }

  // =====================================================================
  // TAB 3: EVALUATION & SCORING
  // =====================================================================
  function renderEvaluation() {
    const rows = C.assignmentRows(dbView()).filter((r) => r.submissionId);
    const f = S.evalFilter;
    const shown = rows.filter((r) => f === 'all' || (f === 'open' ? r.status !== 'Evaluated' : r.status === 'Evaluated'));
    if ((!S.selectedSub || !byId('submissions', S.selectedSub)) && shown.length) S.selectedSub = shown[0].submissionId;
    const pending = rows.filter((r) => r.status === 'Submitted' || (r.status === 'Pending Evaluation' && !S.busy[r.submissionId])).length;
    const list = shown
      .map((r) => `<li><button class="li" type="button" data-act="pick-sub" data-id="${r.submissionId}" aria-current="${S.selectedSub === r.submissionId}">
        <span class="t">${esc(r.traineeName)}</span><span class="s">${esc(r.quizTitle)}</span>
        <span class="tags">${S.busy[r.submissionId] ? '<span class="tag warn"><span class="spin"></span> Grading</span>' : statusTag(r.status)}${r.percent != null ? `<span class="tag plain">${r.percent}%${r.scoreSource === 'ai' ? ' Claude' : ' final'}</span>` : ''}</span></button></li>`)
      .join('');
    return `<div class="split"><aside class="side"><section class="panel">
      <div class="panel-head"><h2>Submissions</h2></div>
      <select id="ev-filter" data-filter="evalFilter" aria-label="Show"><option value="open" ${f === 'open' ? 'selected' : ''}>Not final yet</option><option value="done" ${f === 'done' ? 'selected' : ''}>Final</option><option value="all" ${f === 'all' ? 'selected' : ''}>All</option></select>
      <button class="btn primary small" type="button" data-act="eval-all" ${pending && !S.evalAllBusy && caps.sample ? '' : 'disabled'}>${S.evalAllBusy ? '<span class="spin"></span> Grading…' : `Grade all new submissions (${pending})`}</button>
      <ul class="list">${list || '<li class="empty">Nothing here yet.</li>'}</ul></section></aside>
      <div class="stack">${S.selectedSub && byId('submissions', S.selectedSub) ? renderEvalDetail(S.selectedSub) : '<section class="panel"><div class="empty">Select a submission.</div></section>'}</div></div>`;
  }

  function scoreBox(label, o, sub, id) {
    return `<div class="score" ${id ? `id="${id}"` : ''}><span class="small muted"><b>${label}</b></span>${o ? `<span class="big">${o.earned} / ${o.possible}</span><span>${pctTxt(o.percent)}</span><div class="bar"><span style="width:${Math.max(0, Math.min(100, o.percent || 0))}%"></span></div>` : '<span class="big muted">—</span>'}<span class="small muted">${sub || ''}</span></div>`;
  }

  function makeReviewDraft(ev) {
    const src = ev.trainer || ev.ai;
    return { evalId: ev.id, tickets: clone(src.tickets).map((t) => ({ ...t, trainerNotes: t.trainerNotes || '' })), comments: ev.trainer ? ev.trainer.comments || '' : '', overrideResult: ev.trainer ? ev.trainer.overrideResult || '' : '', reviewer: '' };
  }

  function renderEvalDetail(subId) {
    const s = byId('submissions', subId);
    const q = byId('quizzes', s.quizId) || { title: 'Deleted quiz', tickets: [] };
    const ev = currentEval(s.id);
    const versions = store.evaluations.filter((e) => e.submissionId === s.id).sort((a, b) => a.version - b.version);
    const head = `<div class="panel-head"><div><h2>${esc(s.traineeName)}</h2><p class="small muted">${esc(q.title)} · ${esc(s.cohort)} · submitted ${fmtDT(s.submittedAt)}</p></div>
      <div class="actions"><button class="btn small" type="button" data-act="open-sub" data-id="${s.id}">Raw answers</button>
      ${S.busy[s.id] ? '<span class="pill busy">Claude is grading</span>' : ev ? `<button class="btn small" type="button" data-act="reevaluate" data-id="${s.id}">Reevaluate</button>` : `<button class="btn primary" type="button" data-act="evaluate" data-id="${s.id}" ${caps.sample ? '' : 'disabled'}>Grade with Claude</button>`}</div></div>`;
    if (!ev) {
      return `<section class="panel">${head}
        ${S.failed[s.id] ? `<div class="notice bad">${esc(S.failed[s.id])} The trainee's answers are unchanged.</div>` : ''}
        ${S.busy[s.id] ? '<div class="notice info">Claude is grading each ticket against its rubric. This usually takes under a minute.</div>' : `<div class="notice info">${caps.sample ? 'Claude grades each ticket against only the rubric and reference material you wrote for it. The first time, claude.ai asks you to allow this page to use Claude.' : 'Claude grading is not available in this view. Open the page in claude.ai.'}</div>`}
        ${answersHtml(s)}</section>`;
    }
    const rd = S.reviewDraft && S.reviewDraft.evalId === ev.id ? S.reviewDraft : (S.reviewDraft = makeReviewDraft(ev));
    const ai = ev.ai;
    const fin = ev.trainer && ev.trainer.approvedAt ? ev.trainer : null;
    const draftOverall = C.summarize(rd.tickets.map((t) => C.recomputeTicket(t)), q);
    const res = fin ? fin.overrideResult || fin.overall.result : ai.overall.result;
    const out = [];
    out.push(`<section class="panel">${head}
      <div class="scores">
        ${scoreBox('Claude grade (original)', ai.overall, `Version ${ev.version} · ${fmtDT(ev.createdAt)}`)}
        ${scoreBox('Trainer final', fin ? fin.overall : null, fin ? `Approved by ${esc(fin.approvedByName || nameOf(fin.approvedBy))} · ${fmtDT(fin.approvedAt)}` : 'Not approved yet')}
        ${scoreBox('Your edits', draftOverall, 'Save or approve to keep', 'draft-box')}
      </div>
      <div class="grid g3">
        <div><h4>Result</h4><div class="tags" style="margin-top:6px">${resultTag(res)}${ai.overall.autoFail ? '<span class="tag bad">Auto-fail condition met</span>' : ''}</div>${ai.overall.passingThreshold ? `<p class="small muted" style="margin:4px 0 0">Pass mark ${ai.overall.passingThreshold}%</p>` : ''}</div>
        <div><h4>Main strengths</h4>${ai.overall.strengths.length ? `<ul class="plain">${ai.overall.strengths.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : '<p class="small muted">None identified.</p>'}</div>
        <div><h4>Recurring weaknesses</h4>${ai.overall.recurringWeaknesses.length ? `<ul class="plain">${ai.overall.recurringWeaknesses.map((x) => `<li>${esc(x.category)} <span class="muted">(${plural(x.count, 'ticket')})</span></li>`).join('')}</ul>` : '<p class="small muted">None identified.</p>'}
          ${ai.overall.coachingFocus.length ? `<p class="small" style="margin:4px 0 0"><b>Coaching focus:</b> ${ai.overall.coachingFocus.map(esc).join(', ')}</p>` : ''}</div>
      </div></section>`);
    ai.tickets.forEach((t, ti) => {
      const rt = rd.tickets[ti];
      const ans = s.answers[ti] || {};
      out.push(`<section class="panel"><div class="panel-head"><div><h3>Ticket ${t.ticketNumber} <a class="small" style="font-family:var(--body);font-weight:400" href="${safeUrl(t.ticketUrl)}" target="_blank" rel="noopener">${esc(t.ticketUrl)}</a></h3>
        <p class="small muted">Claude ${t.earned} / ${t.possible} (${pctTxt(t.percent)}) · Your edits <b id="tsum-${ti}">${rt.earned} / ${rt.possible}</b></p></div>
        <div class="tags">${t.needsTrainerReview ? '<span class="tag warn">Flagged for trainer review</span>' : '<span class="tag ok">No flags</span>'}${t.autoFail ? '<span class="tag bad">Auto-fail</span>' : ''}</div></div>
        ${t.criticalErrors.length ? `<div class="notice bad"><b>Critical errors</b><ul>${t.criticalErrors.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''}
        ${t.autoFail ? `<div class="notice bad">Automatic fail condition met: ${esc(t.autoFailReason)}</div>` : ''}
        ${t.reviewNotes ? `<p class="small muted" style="margin:0">Grader's note: ${esc(t.reviewNotes)}</p>` : ''}
        <div class="tw"><table><thead><tr><th>Criterion</th><th class="num">Max</th><th class="num">Claude</th><th>Reason for deduction</th><th style="width:86px">Trainer</th><th>Trainer reason</th></tr></thead><tbody>
        ${t.criteria.map((c, ci) => `<tr><td>${esc(c.name)}${c.needsReview ? '<div class="flag">Needs your judgment</div>' : ''}</td><td class="num">${c.maxPoints}</td><td class="num">${c.awarded}</td><td class="small">${esc(c.reason) || '<span class="muted">No deduction</span>'}</td>
          <td><input type="number" id="rv-${ev.id}-${ti}-${ci}" min="0" max="${c.maxPoints}" step="0.5" data-review="${ti}.${ci}.awarded" value="${esc(rt.criteria[ci].awarded)}" aria-label="Trainer points for ${esc(c.name)}"></td>
          <td><textarea id="rr-${ev.id}-${ti}-${ci}" rows="1" data-review="${ti}.${ci}.reason" placeholder="Optional" aria-label="Trainer reason for ${esc(c.name)}">${esc(rt.criteria[ci].reason)}</textarea></td></tr>`).join('')}
        </tbody></table></div>
        <div class="grid g3">
          <div><h4>What went well</h4>${t.strengths.length ? `<ul class="plain">${t.strengths.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : '<p class="small muted">None identified.</p>'}</div>
          <div><h4>Areas for improvement</h4>${rt.improvements.length ? `<ul class="plain">${rt.improvements.map((x, ii) => `<li><span class="tag plain">${esc(x.category)}</span> ${esc(x.detail)} <button class="linkish small" type="button" data-act="del-imp" data-t="${ti}" data-i="${ii}">remove</button></li>`).join('')}</ul>` : '<p class="small muted">None.</p>'}
            <details class="block" data-k="imp-${ev.id}-${ti}"><summary class="small">Add one</summary><div class="body"><select id="imp-cat-${ti}" aria-label="Category">${C.IMPROVEMENT_CATEGORIES.map((x) => `<option>${esc(x)}</option>`).join('')}</select><input type="text" id="imp-det-${ti}" placeholder="What to improve" aria-label="Detail"><button class="btn small" type="button" data-act="add-imp" data-t="${ti}">Add</button></div></details></div>
          <div><h4>Suggested coaching</h4>${t.coaching.length ? `<ul class="plain">${t.coaching.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : '<p class="small muted">None.</p>'}</div>
        </div>
        <label class="field">Trainer notes for this ticket<textarea id="rn-${ev.id}-${ti}" rows="2" data-review="${ti}.notes">${esc(rt.trainerNotes || '')}</textarea></label>
        <details class="block" data-k="ans-${ev.id}-${ti}"><summary class="small">Trainee's answers for Ticket ${t.ticketNumber}</summary><div class="body" style="gap:6px">${C.QUESTIONS.map((qq) => `<div class="qlabel">${esc(qq.label)}</div><div class="answer">${esc(ans[qq.key]) || '<span class="muted">(blank)</span>'}</div>`).join('')}</div></details>
      </section>`);
    });
    out.push(`<section class="panel"><h3>Trainer review</h3>
      <div class="grid g2"><label class="field">Comments for the trainee's record<textarea id="rc-${ev.id}" rows="3" data-review="comments">${esc(rd.comments)}</textarea></label>
        <div class="grid"><label class="field">Override the overall result<select id="ro-${ev.id}" data-review="overrideResult">${['', 'Pass', 'Fail', 'Needs retake'].map((x) => `<option value="${x}" ${rd.overrideResult === x ? 'selected' : ''}>${x || 'No override'}</option>`).join('')}</select></label>
        ${me && me.id ? `<p class="small muted" style="margin:0">Approving records you as the reviewer.</p>` : `<label class="field">Your name<input type="text" id="rname-${ev.id}" data-review="reviewer" value="${esc(rd.reviewer)}"></label>`}</div></div>
      <div class="actions" style="justify-content:flex-end"><button class="btn" type="button" data-act="reset-review">Discard edits</button><button class="btn" type="button" data-act="save-review">Save edits</button><button class="btn primary" type="button" data-act="approve-review">Approve as final</button></div>
      <p class="small muted" style="margin:0">The original Claude grade stays unchanged. Your version is saved beside it with the time and reviewer.</p></section>`);
    if (versions.length > 1 || (ev.history || []).length) {
      out.push(`<details class="block" data-k="hist-${s.id}"><summary>Audit history</summary><div class="body small">
        ${versions.map((v) => `<div>Version ${v.version}, graded ${fmtDT(v.createdAt)}: Claude ${v.ai.overall.earned} / ${v.ai.overall.possible}${v.trainer && v.trainer.approvedAt ? `, trainer final ${v.trainer.overall.earned} / ${v.trainer.overall.possible}` : ''}${v.supersededAt ? `. Replaced ${fmtDT(v.supersededAt)} (${esc(v.supersededReason)}).` : '. Current.'}</div>`).join('')}
        ${(ev.history || []).map((h) => `<div>${fmtDT(h.at)}: ${esc(nameOf(h.by))} ${esc(h.action)} the review (score ${h.after.earned}).</div>`).join('')}</div></details>`);
    }
    return out.join('');
  }

  // =====================================================================
  // TAB 4: PERFORMANCE DASHBOARD
  // =====================================================================
  function dashRows() {
    const f = S.dash;
    return C.assignmentRows(dbView()).filter((r) =>
      (!f.cohort || r.cohort === f.cohort) && (!f.traineeId || r.traineeId === f.traineeId) && (!f.quizId || r.quizId === f.quizId) && (!f.trainingDay || r.trainingDay === f.trainingDay) && (!f.status || r.status === f.status) &&
      (!f.from || (r.submittedAt && r.submittedAt.slice(0, 10) >= f.from)) && (!f.to || (r.submittedAt && r.submittedAt.slice(0, 10) <= f.to)));
  }

  function renderDashboard() {
    const f = S.dash;
    const rows = dashRows();
    const scored = rows.filter((r) => r.percent != null);
    const claudeOnly = scored.filter((r) => r.scoreSource === 'ai').length;
    const avg = scored.length ? C.round1(scored.reduce((a, r) => a + r.percent, 0) / scored.length) : null;
    const pf = scored.filter((r) => r.result === 'Pass' || r.result === 'Fail');
    const pass = pf.length ? C.round1((pf.filter((r) => r.result === 'Pass').length / pf.length) * 100) : null;
    const days = [...new Set(store.quizzes.map((q) => q.trainingDay).filter(Boolean))];
    const metric = (l, v, s) => `<div class="metric"><span class="l">${l}</span><span class="v">${v}</span>${s ? `<span class="s">${s}</span>` : ''}</div>`;
    const imp = new Map();
    let n = 0;
    rows.forEach((r) => {
      const eff = C.effectiveResult(byId('evaluations', r.evaluationId));
      if (!eff) return;
      n++;
      eff.tickets.forEach((t) => t.improvements.forEach((i) => {
        if (!imp.has(i.category)) imp.set(i.category, { count: 0, who: new Map() });
        const m = imp.get(i.category);
        m.count++;
        m.who.set(r.traineeName, (m.who.get(r.traineeName) || 0) + 1);
      }));
    });
    const impRows = [...imp.entries()].sort((a, b) => b[1].count - a[1].count);
    const quizzes = [...new Map(rows.map((r) => [r.quizId, r.quizTitle])).entries()];
    const people = [...new Map(rows.map((r) => [r.traineeId || r.traineeEmail, r.traineeName])).entries()];
    const sel = (id, key, label, opts) => `<label class="field">${label}<select id="${id}" data-filter="dash.${key}"><option value="">All</option>${opts}</select></label>`;
    return `<div class="stack">
      <section class="panel"><div class="grid g4">
        ${sel('df-cohort', 'cohort', 'Cohort', cohorts().map((c) => `<option ${f.cohort === c ? 'selected' : ''}>${esc(c)}</option>`).join(''))}
        ${sel('df-trainee', 'traineeId', 'Trainee', store.trainees.map((t) => `<option value="${t.id}" ${f.traineeId === t.id ? 'selected' : ''}>${esc(t.name)}</option>`).join(''))}
        ${sel('df-quiz', 'quizId', 'Quiz title', store.quizzes.map((q) => `<option value="${q.id}" ${f.quizId === q.id ? 'selected' : ''}>${esc(q.title)}</option>`).join(''))}
        ${sel('df-day', 'trainingDay', 'Training day', days.map((d) => `<option ${f.trainingDay === d ? 'selected' : ''}>${esc(d)}</option>`).join(''))}
        <label class="field">Submitted from<input type="date" id="df-from" data-filter="dash.from" value="${esc(f.from)}"></label>
        <label class="field">Submitted to<input type="date" id="df-to" data-filter="dash.to" value="${esc(f.to)}"></label>
        ${sel('df-status', 'status', 'Evaluation status', C.EVAL_STATUSES.map((s) => `<option ${f.status === s ? 'selected' : ''}>${s}</option>`).join(''))}
        <div style="display:flex;align-items:end"><button class="btn ghost" type="button" data-act="clear-dash">Clear filters</button></div>
      </div></section>
      <div class="metrics">
        ${metric('Assigned assessments', rows.filter((r) => !r.unassigned).length)}
        ${metric('Submissions', rows.filter((r) => r.submissionId).length)}
        ${metric('Pending submissions', rows.filter((r) => r.status === 'Awaiting Submission').length)}
        ${metric('Pending evaluations', rows.filter((r) => ['Submitted', 'Pending Evaluation', 'Needs Trainer Review'].includes(r.status)).length, 'Not approved as final')}
        ${metric('Average score', avg == null ? '—' : avg + '%', scored.length ? `${scored.length} scored${claudeOnly ? `, ${claudeOnly} not yet approved` : ''}` : 'Nothing scored yet')}
        ${metric('Passing rate', pass == null ? '—' : pass + '%', pf.length ? `${pf.length} with a pass mark set` : 'No pass mark set')}
      </div>
      <div class="grid g2" style="align-items:start">
        <section class="panel"><h2>Trainee performance</h2><div class="tw"><table><thead><tr><th>Trainee</th><th>Quiz</th><th class="num">Score</th><th>Status</th></tr></thead><tbody>
          ${rows.map((r) => `<tr><td><button class="linkish" type="button" data-act="history" data-id="${esc(r.traineeId || '')}" data-name="${esc(r.traineeName)}">${esc(r.traineeName)}</button><div class="small muted">${esc(r.cohort)}</div></td><td>${esc(r.quizTitle)}</td><td class="num">${pctTxt(r.percent)}${r.scoreSource === 'ai' ? '<div class="small muted">Claude</div>' : ''}</td><td>${statusTag(r.status)}</td></tr>`).join('') || '<tr><td colspan="4"><div class="empty">No assessments match these filters.</div></td></tr>'}
        </tbody></table></div></section>
        <section class="panel"><div><h2>Areas for improvement</h2><p class="small muted" style="margin:2px 0 0">${n ? `From ${plural(n, 'graded submission')}, using the trainer final where approved and Claude's grade otherwise.` : ''}</p></div>
          ${n === 0 ? '<div class="empty">No graded submissions in this view yet, so there is nothing to summarize.</div>' : impRows.length ? `<div class="tw"><table><thead><tr><th>Issue</th><th class="num">Times</th><th>Who needs coaching</th></tr></thead><tbody>${impRows.map(([c, m]) => `<tr><td>${esc(c)}</td><td class="num">${m.count}</td><td class="small">${[...m.who.entries()].sort((a, b) => b[1] - a[1]).map(([w, k]) => `${esc(w)} (${k})`).join(', ')}</td></tr>`).join('')}</tbody></table></div>` : '<div class="empty">No areas for improvement recorded.</div>'}
        </section>
      </div>
      <section class="panel"><div><h2>Compare across quizzes</h2><p class="small muted" style="margin:2px 0 0">Score per trainee and quiz. Blank means not assigned.</p></div>
        ${quizzes.length ? `<div class="tw"><table><thead><tr><th>Trainee</th>${quizzes.map(([, t]) => `<th>${esc(t)}</th>`).join('')}</tr></thead><tbody>
        ${people.map(([pid, nm]) => `<tr><td>${esc(nm)}</td>${quizzes.map(([qid]) => { const r = rows.find((x) => (x.traineeId || x.traineeEmail) === pid && x.quizId === qid); return `<td class="num">${r ? (r.percent != null ? r.percent + '%' : `<span class="small muted">${esc(r.status)}</span>`) : ''}</td>`; }).join('')}</tr>`).join('')}
        </tbody></table></div>` : '<div class="empty">No data yet.</div>'}</section>
      ${renderExports()}
    </div>`;
  }

  function renderExports() {
    const line = (label, kind) => `<tr><td>${label}</td><td style="text-align:right;white-space:nowrap"><button class="btn small" type="button" data-act="export" data-kind="${kind}" data-fmt="xlsx">Excel</button> <button class="btn small" type="button" data-act="export" data-kind="${kind}" data-fmt="csv">CSV</button></td></tr>`;
    return `<section class="panel"><div class="panel-head"><div><h2>Export</h2><p class="small muted">Uses the cohort and quiz filters above. Raw answers are exported exactly as received.</p></div>
      <button class="btn primary" type="button" data-act="export" data-kind="all" data-fmt="xlsx" ${caps.downloads ? '' : 'disabled'}>Full workbook (Excel, 6 tabs)</button></div>
      ${caps.downloads ? '' : '<div class="notice warn">Downloads are not available in this view. Open the page in claude.ai.</div>'}
      <div class="tw"><table><tbody>${line('Raw trainee responses', 'raw')}${line('Evaluation results', 'evaluations')}${line('Score breakdown', 'breakdown')}${line('Areas for improvement', 'areas')}${line('Cohort performance', 'cohort')}${line('Quiz registry', 'registry')}
      <tr><td><label class="field" style="max-width:320px">Individual trainee report<select id="ex-trainee">${store.trainees.map((t) => `<option value="${t.id}" ${S.dash.traineeId === t.id ? 'selected' : ''}>${esc(t.name)}</option>`).join('')}</select></label></td>
      <td style="text-align:right;white-space:nowrap;vertical-align:bottom"><button class="btn small" type="button" data-act="export" data-kind="trainee" data-fmt="xlsx">Excel</button> <button class="btn small" type="button" data-act="export" data-kind="trainee" data-fmt="csv">CSV</button></td></tr></tbody></table></div></section>`;
  }

  function traineeHistory(id, name) {
    const rows = C.assignmentRows(dbView()).filter((r) => (id ? r.traineeId === id : r.traineeName === name)).sort((a, b) => String(a.submittedAt || '').localeCompare(String(b.submittedAt || '')));
    openModal(
      `Assessment history · ${name}`,
      `<div class="tw"><table><thead><tr><th>Quiz</th><th>Day</th><th>Submitted</th><th class="num">Score</th><th>Result</th><th>Status</th><th></th></tr></thead><tbody>
      ${rows.map((r) => `<tr><td>${esc(r.quizTitle)}</td><td>${esc(r.trainingDay)}</td><td>${fmtDT(r.submittedAt)}</td><td class="num">${r.percent != null ? `${r.earned} / ${r.possible} (${r.percent}%)` : '—'}</td><td>${r.result ? resultTag(r.result) : ''}</td><td>${statusTag(r.status)}</td><td>${r.submissionId ? `<button class="btn small" type="button" data-act="go-eval" data-id="${r.submissionId}">Open</button>` : ''}</td></tr>`).join('')}
      </tbody></table></div>`,
      `<button class="btn" type="button" data-act="close-modal">Close</button>`
    );
  }

  const KINDS = {
    all: ['full-workbook', ['Quiz Registry', 'Raw Responses', 'Evaluation Results', 'Score Breakdown', 'Areas for Improvement', 'Cohort Performance']],
    registry: ['quiz-registry', ['Quiz Registry']],
    raw: ['raw-responses', ['Raw Responses']],
    evaluations: ['evaluation-results', ['Evaluation Results']],
    breakdown: ['score-breakdown', ['Score Breakdown']],
    areas: ['areas-for-improvement', ['Areas for Improvement']],
    cohort: ['cohort-performance', ['Cohort Performance']],
    trainee: ['trainee-report', ['Evaluation Results', 'Score Breakdown', 'Areas for Improvement', 'Raw Responses']],
  };
  const slug = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40);

  async function doExport(kind, fmt) {
    if (!caps.downloads) throw new Error('Downloads are not available in this view.');
    const filter = { cohort: S.dash.cohort, quizId: S.dash.quizId, traineeId: kind === 'trainee' ? $('ex-trainee') && $('ex-trainee').value : '' };
    if (kind === 'trainee' && !filter.traineeId) throw new Error('Add a trainee first.');
    const [label, sheets] = KINDS[kind];
    const tables = C.exportTables(dbView(), filter);
    let name = label;
    if (kind === 'trainee') name += '-' + slug((byId('trainees', filter.traineeId) || {}).name);
    if (filter.cohort) name += '-' + slug(filter.cohort);
    name += '-' + new Date().toISOString().slice(0, 10);
    let data;
    if (fmt === 'csv') {
      const sheet = kind === 'trainee' ? 'Score Breakdown' : sheets[0];
      data = '﻿' + (C.toCSV(tables[sheet]) || 'No rows match this export.\r\n');
      name += '.csv';
    } else {
      if (!window.XLSX) throw new Error('The Excel library did not load. Use CSV, or reload the page.');
      const wb = XLSX.utils.book_new();
      if (kind === 'trainee') {
        const tr = byId('trainees', filter.traineeId) || {};
        const rows = C.assignmentRows(dbView()).filter((r) => r.traineeId === filter.traineeId);
        const aoa = [['Trainee report'], ['Trainee', tr.name || ''], ['Email', tr.email || ''], ['Cohort', tr.cohort || ''], ['Generated', new Date().toISOString()], [], ['Quiz', 'Status', 'Points earned', 'Points possible', 'Score (%)', 'Score source', 'Result']];
        rows.forEach((r) => aoa.push([r.quizTitle, r.status, r.earned ?? '', r.possible ?? '', r.percent ?? '', r.scoreSource === 'trainer' ? 'Trainer final' : r.scoreSource === 'ai' ? 'Claude, not yet approved' : '', r.result ?? '']));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Summary');
      }
      sheets.forEach((s) => {
        const rows = tables[s];
        const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Note: 'No rows match this export.' }]);
        const keys = Object.keys(rows[0] || { Note: '' });
        ws['!cols'] = keys.map((k) => ({ wch: Math.min(60, Math.max(12, k.length + 2)) }));
        XLSX.utils.book_append_sheet(wb, ws, s.slice(0, 31));
      });
      data = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })]);
      name += '.xlsx';
    }
    try {
      await caps.downloads.save({ filename: name, data });
      toast('Saved ' + name);
    } catch (e) {
      if (e && e.code === 'declined') return;
      throw new Error(e && e.code === 'rate_limited' ? 'A save prompt is already open.' : 'The file could not be saved here.');
    }
  }

  // =====================================================================
  // SETUP and demo data
  // =====================================================================
  function openSetup() {
    const row = (label, ok, txt) => `<tr><td><b>${label}</b></td><td>${ok === true ? '<span class="pill ok">Ready</span>' : ok === false ? '<span class="pill bad">Unavailable</span>' : '<span class="pill plain">Optional</span>'}</td><td class="small">${txt}</td></tr>`;
    openModal(
      'Setup and demo data',
      `<div class="tw"><table><tbody>
        ${row('Shared records', !!caps.db, caps.db ? 'Saved in this page\'s database. Contributors and Editors read and change them; Viewers and Commenters cannot see them.' : 'Not reachable in this view. Open the page in claude.ai.')}
        ${row('Claude grading', !!caps.sample, caps.sample ? 'Runs on the account of whoever clicks Grade. claude.ai asks for permission the first time.' : 'Not available in this view.')}
        ${row('Google Sheets import', caps.mcp ? null : false, caps.mcp ? 'Reads a response sheet through your own Google Sheets connector when you import. It never writes.' : 'Not available in this view. CSV import still works.')}
        ${row('Excel and CSV files', !!caps.downloads, caps.downloads ? 'You confirm each file before it saves.' : 'Not available in this view.')}
      </tbody></table></div>
      <div class="notice info">Google Forms are created by a script you run yourself from Quiz Builder (Get Google Form script). The page never creates, opens or shares a form, and never sends messages.</div>
      <h3>Staging data</h3>
      <p class="small" style="margin:0">Load fictional trainees and a two-ticket sample quiz with two sample submissions. No real trainee or customer data is used.</p>
      <div class="actions"><button class="btn primary" type="button" data-act="seed">Load fictional demo data</button><button class="btn danger" type="button" data-act="reset">Delete all staging data…</button></div>`,
      `<button class="btn" type="button" data-act="close-modal">Close</button>`
    );
  }

  function demoData() {
    const T = [
      ['Test Trainee Alpha', 'alpha.trainee@example.com', 'Staging Cohort A'],
      ['Test Trainee Bravo', 'bravo.trainee@example.com', 'Staging Cohort A'],
      ['Test Trainee Charlie', 'charlie.trainee@example.com', 'Staging Cohort A'],
      ['Test Trainee Delta', 'delta.trainee@example.com', 'Staging Cohort B'],
      ['Test Trainee Echo', 'echo.trainee@example.com', 'Staging Cohort B'],
    ];
    const now = new Date().toISOString();
    const trainees = T.filter(([, e]) => !store.trainees.some((x) => x.email === e)).map(([name, email, cohort]) => ({ id: C.uid('trn'), name, email, cohort, createdAt: now }));
    return { trainees, now };
  }

  async function seed() {
    const { trainees, now } = demoData();
    for (const t of trainees) await put('trainees', t);
    const all = store.trainees;
    const q = window.__HUB_DEMO_QUIZ__(all, C);
    q.form = { mode: 'staging', createdAt: now, approvedAt: now, approvedBy: me && me.id, outOfDate: false };
    q.status = 'Ready';
    q.createdBy = me && me.id;
    await put('quizzes', q);
    const ans = window.__HUB_DEMO_ANSWERS__(q);
    const a = all.find((t) => t.email === 'alpha.trainee@example.com');
    const b = all.find((t) => t.email === 'bravo.trainee@example.com');
    await recordSubmission(q, { traineeName: a.name, traineeEmail: a.email, cohort: a.cohort, answers: ans.strong }, 'staging-test');
    await recordSubmission(q, { traineeName: b.name, traineeEmail: b.email, cohort: b.cohort, answers: ans.weak }, 'staging-test');
    S.draft = null;
  }

  // =====================================================================
  // EVENTS
  // =====================================================================
  function setPath(obj, path, value) {
    const parts = path.split('.');
    let o = obj;
    for (let i = 0; i < parts.length - 1; i++) o = o[/^\d+$/.test(parts[i]) ? Number(parts[i]) : parts[i]];
    o[parts[parts.length - 1]] = value;
  }
  async function leaveDraft() {
    if (!S.dirty) return true;
    const ok = await ask({ title: 'Discard unsaved changes?', body: 'This quiz has changes you have not saved.', ok: 'Discard changes', danger: true });
    return ok === true;
  }
  function markDirty(rerender) {
    S.dirty = true;
    if (rerender) render();
  }

  document.addEventListener('click', async (e) => {
    const tab = e.target.closest('[data-tab]');
    if (tab) {
      S.tab = tab.dataset.tab;
      lsSet('tab', S.tab);
      render();
      return;
    }
    if (e.target.id === 'setup-btn') return openSetup();
    const el = e.target.closest('[data-act]');
    if (!el) return;
    const a = el.dataset.act;
    if (a === 'modal-bg' && e.target !== el) return;
    if (el.tagName === 'INPUT' || el.tagName === 'SELECT') {
      if (a !== 'toggle-trainee') return;
    }
    const d = S.draft;
    const ti = el.dataset.t != null ? Number(el.dataset.t) : null;
    try {
      switch (a) {
        case 'modal-bg':
        case 'close-modal':
        case 'ask-cancel':
          closeModal(null);
          break;
        case 'ask-ok': {
          const inp = $('ask-input');
          closeModal(inp ? inp.value : true);
          break;
        }
        case 'new-quiz':
          if (await leaveDraft()) { startDraft(null); render(); }
          break;
        case 'pick-quiz':
          if (await leaveDraft()) { startDraft(byId('quizzes', el.dataset.id)); render(); }
          break;
        case 'save-draft':
          await saveDraft();
          break;
        case 'delete-quiz':
          if (store.submissions.some((s) => s.quizId === d.id)) throw new Error('This quiz has submissions, so it is kept for the record. Set its status to Closed instead.');
          if ((await ask({ title: 'Delete this quiz?', body: `“${esc(d.title)}” and its rubrics will be deleted for everyone.`, ok: 'Delete quiz', danger: true })) === true) {
            await del('quizzes', d.id);
            S.draft = null;
            S.dirty = false;
            render();
            toast('Quiz deleted.');
          }
          break;
        case 'preview':
          openPreview();
          break;
        case 'approve-form': {
          if (!$('approve-check').checked) return toast('Tick the approval box first.', true);
          el.disabled = true;
          const now = new Date().toISOString();
          S.draft.form = { ...(S.draft.form || {}), mode: 'staging', createdAt: (S.draft.form && S.draft.form.createdAt) || now, approvedAt: now, approvedBy: me && me.id, outOfDate: false };
          if (S.draft.status === 'Draft') S.draft.status = 'Ready';
          closeModal();
          await saveDraft(true);
          toast('Form approved. Open the staging form to test it, or get the Google Form script.');
          break;
        }
        case 'open-staging-form':
          openStagingForm(el.dataset.id);
          break;
        case 'apps-script': {
          const q = byId('quizzes', el.dataset.id);
          const code = appsScriptFor(q);
          openModal(
            'Google Form script',
            `<ol class="small" style="margin:0;padding-left:20px"><li>Go to <a href="https://script.google.com/home/projects/create" target="_blank" rel="noopener">script.google.com</a> and start a new project in the Google account that should own the form.</li><li>Paste this script, save, and run <b>createQuizForm</b>. Allow access when Google asks.</li><li>Open the execution log, copy the form link and the response sheet link, and paste them into <b>Link Google Form</b>.</li></ol>
            <div class="notice info">The script makes the form closed and unshared. It holds only what trainees see: no rubric, no expected answers.</div>
            <pre class="code" id="as-code">${esc(code)}</pre>`,
            `<button class="btn" type="button" data-act="close-modal">Close</button>${caps.downloads ? '<button class="btn" type="button" data-act="dl-script">Download .txt</button>' : ''}<button class="btn primary" type="button" data-act="copy-script">Copy script</button>`
          );
          $('modal-root').dataset.code = code;
          $('modal-root').dataset.name = slug(q.title) + '-google-form-script.txt';
          break;
        }
        case 'copy-script':
          copyText($('modal-root').dataset.code, $('as-code'));
          break;
        case 'dl-script':
          await caps.downloads.save({ filename: $('modal-root').dataset.name, data: $('modal-root').dataset.code }).catch((er) => er && er.code !== 'declined' && toast('The file could not be saved here.', true));
          break;
        case 'link-google': {
          const q = byId('quizzes', el.dataset.id);
          const g = (q.form && q.form.google) || {};
          openModal('Link Google Form', `<label class="field">Google Form link (the one trainees open)<input type="url" id="lg-form" value="${esc(g.formUrl || '')}" placeholder="https://docs.google.com/forms/d/e/…/viewform"></label>
            <label class="field">Response sheet link<input type="url" id="lg-sheet" value="${esc(g.sheetUrl || '')}" placeholder="https://docs.google.com/spreadsheets/d/…"></label>`,
            `<button class="btn" type="button" data-act="close-modal">Cancel</button><button class="btn primary" type="button" data-act="save-google" data-id="${q.id}">Save links</button>`, { narrow: true });
          break;
        }
        case 'save-google': {
          const formUrl = $('lg-form').value.trim();
          const sheetUrl = $('lg-sheet').value.trim();
          if (formUrl && !/^https:\/\/(docs\.google\.com\/forms|forms\.gle)\//.test(formUrl)) return toast('That doesn\'t look like a Google Form link.', true);
          if (sheetUrl && !/^https:\/\/docs\.google\.com\/spreadsheets\//.test(sheetUrl)) return toast('That doesn\'t look like a Google Sheets link.', true);
          const q = clone(byId('quizzes', el.dataset.id));
          q.form.google = { formUrl, sheetUrl, linkedAt: new Date().toISOString(), linkedBy: me && me.id };
          q.updatedAt = new Date().toISOString();
          await put('quizzes', q);
          closeModal();
          startDraft(byId('quizzes', q.id));
          render();
          toast('Google links saved.');
          break;
        }
        case 'copy':
          copyText(el.dataset.text);
          break;
        case 'send-dialog': {
          const q = byId('quizzes', el.dataset.id);
          const who = q.assignedTrainees.map((id) => (byId('trainees', id) || {}).name).filter(Boolean);
          const msg = `Hi team, please complete "${q.title}"${q.deadline ? ' by ' + fmtDate(q.deadline) : ''}: ${q.form.google.formUrl}`;
          openModal('Send to trainees', `<div class="notice warn">This page doesn't send anything. Copy the message and send it yourself once the prototype is approved.</div>
            <p class="small" style="margin:0">Assigned (${who.length}): ${who.map(esc).join(', ') || 'nobody yet'}</p><textarea id="send-msg" rows="4">${esc(msg)}</textarea>`,
            `<button class="btn" type="button" data-act="close-modal">Close</button><button class="btn primary" type="button" data-act="copy-msg">Copy message</button>`, { narrow: true });
          break;
        }
        case 'copy-msg':
          copyText($('send-msg').value, $('send-msg'));
          break;
        case 'toggle-trainee': {
          const id = el.dataset.id;
          d.assignedTrainees = el.checked ? [...new Set([...d.assignedTrainees, id])] : d.assignedTrainees.filter((x) => x !== id);
          S.dirty = true;
          $('assign-box').innerHTML = renderAssign(d);
          break;
        }
        case 'assign-cohort':
          d.assignedTrainees = [...new Set([...d.assignedTrainees, ...store.trainees.filter((t) => t.cohort === d.cohort).map((t) => t.id)])];
          S.dirty = true;
          $('assign-box').innerHTML = renderAssign(d);
          break;
        case 'clear-assign':
          d.assignedTrainees = [];
          S.dirty = true;
          $('assign-box').innerHTML = renderAssign(d);
          break;
        case 'add-ticket':
          d.tickets.push(C.newTicket());
          markDirty(true);
          break;
        case 'del-ticket':
          if (d.tickets.length === 1) return toast('A quiz needs at least one ticket.', true);
          if ((await ask({ title: `Remove Ticket ${ti + 1}?`, body: 'Its link, rubric and reference material will be removed from this draft.', ok: 'Remove ticket', danger: true })) === true) {
            d.tickets.splice(ti, 1);
            markDirty(true);
          }
          break;
        case 'dup-ticket': {
          const cp = clone(d.tickets[ti]);
          cp.id = C.uid('tkt');
          cp.url = '';
          cp.rubric.criteria.forEach((c) => (c.id = C.uid('crit')));
          d.tickets.splice(ti + 1, 0, cp);
          markDirty(true);
          toast('Duplicated. Add the new ticket link.');
          break;
        }
        case 'move-ticket': {
          const j = ti + Number(el.dataset.dir);
          [d.tickets[ti], d.tickets[j]] = [d.tickets[j], d.tickets[ti]];
          markDirty(true);
          break;
        }
        case 'add-criterion':
          d.tickets[ti].rubric.criteria.push({ id: C.uid('crit'), name: '', category: 'other', maxPoints: 0, guidance: '', deductions: [] });
          markDirty(true);
          break;
        case 'suggest': {
          const r = d.tickets[ti].rubric;
          if (r.criteria.length && (await ask({ title: 'Replace the criteria?', body: 'The five suggested criteria will replace the ones on this ticket. You can edit them after.', ok: 'Replace' })) !== true) return;
          if (!(Number(r.totalPoints) > 0)) r.totalPoints = 20;
          r.criteria = C.suggestedCriteria(r.totalPoints);
          markDirty(true);
          break;
        }
        case 'del-criterion':
          d.tickets[ti].rubric.criteria.splice(Number(el.dataset.c), 1);
          markDirty(true);
          break;
        case 'add-deduction':
          d.tickets[ti].rubric.criteria[Number(el.dataset.c)].deductions.push({ description: '', points: 1 });
          markDirty(true);
          break;
        case 'del-deduction':
          d.tickets[ti].rubric.criteria[Number(el.dataset.c)].deductions.splice(Number(el.dataset.d), 1);
          markDirty(true);
          break;
        case 'add-attach': {
          const url = $('att-' + d.tickets[ti].id).value.trim();
          if (!/^https?:\/\//i.test(url)) return toast('Paste a link that starts with https://', true);
          d.tickets[ti].attachments.push({ name: decodeURIComponent(url.split('/').filter(Boolean).pop() || url).slice(0, 80), url });
          markDirty(true);
          break;
        }
        case 'del-attach':
          d.tickets[ti].attachments.splice(Number(el.dataset.a), 1);
          markDirty(true);
          break;
        case 'add-trainee': {
          const name = $('tr-name').value.trim();
          const email = $('tr-email').value.trim().toLowerCase();
          const cohort = $('tr-cohort').value.trim();
          if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !cohort) return toast('Enter a name, a valid email and a cohort.', true);
          if (store.trainees.some((t) => t.email.toLowerCase() === email)) return toast('A trainee with that email is already on the roster.', true);
          await put('trainees', { id: C.uid('trn'), name, email, cohort, createdAt: new Date().toISOString() });
          render();
          toast('Trainee added.');
          break;
        }
        case 'del-trainee': {
          const t = byId('trainees', el.dataset.id);
          if (store.submissions.some((s) => s.traineeId === t.id)) return toast('This trainee has submissions, so they stay on the roster for the record.', true);
          if ((await ask({ title: `Remove ${esc(t.name)}?`, body: 'They will be removed from the roster and from any quiz they were assigned to.', ok: 'Remove', danger: true })) !== true) return;
          for (const q of store.quizzes.filter((x) => x.assignedTrainees.includes(t.id))) await put('quizzes', { ...clone(q), assignedTrainees: q.assignedTrainees.filter((x) => x !== t.id) });
          await del('trainees', t.id);
          if (S.draft) S.draft.assignedTrainees = S.draft.assignedTrainees.filter((x) => x !== t.id);
          render();
          break;
        }
        case 'open-sub':
          e.stopPropagation();
          openSubmission(el.dataset.id);
          break;
        case 'go-eval':
          S.selectedSub = el.dataset.id;
          S.evalFilter = 'all';
          S.tab = 'evaluation';
          S.reviewDraft = null;
          lsSet('tab', S.tab);
          closeModal();
          render();
          break;
        case 'test-response':
          openStagingForm();
          break;
        case 'submit-staging': {
          const q = byId('quizzes', $('sf-body').dataset.quiz);
          const name = $('sf-name').value.trim();
          const email = $('sf-email').value.trim();
          if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return toast('Enter a name and a valid email.', true);
          const answers = q.tickets.map((t) => {
            const o = { ticketId: t.id };
            C.QUESTIONS.forEach((qq) => (o[qq.key] = $(`sf-${t.id}-${qq.key}`).value));
            return o;
          });
          if (answers.some((x) => C.QUESTIONS.some((qq) => !x[qq.key].trim()))) return toast('Answer every question. Type N/A where it does not apply.', true);
          el.disabled = true;
          await recordSubmission(q, { traineeName: name, traineeEmail: email, cohort: q.cohort, answers }, 'staging-test');
          closeModal();
          render();
          toast('Test response saved.');
          break;
        }
        case 'import':
          openImport();
          break;
        case 'import-sheet': {
          const q = byId('quizzes', $('im-quiz').value);
          $('im-status').innerHTML = '<span class="pill busy">Reading the sheet</span>';
          el.disabled = true;
          try {
            const r = await importFromSheet(q, $('im-sheet').value, $('im-tab').value.trim());
            $('im-status').innerHTML = `<div class="notice ok">Read ${plural(r.total, 'response')}: ${r.added} new, ${r.dup} already imported.</div>`;
            if (!q.form.google || !q.form.google.sheetUrl) {
              const qq = clone(q);
              qq.form.google = { ...(qq.form.google || {}), sheetUrl: /^https:/.test($('im-sheet').value.trim()) ? $('im-sheet').value.trim() : '' };
              await put('quizzes', qq);
            }
          } catch (er) {
            $('im-status').innerHTML = `<div class="notice bad">${esc(er.message)}</div>`;
          } finally {
            el.disabled = false;
          }
          break;
        }
        case 'pick-sub':
          S.selectedSub = el.dataset.id;
          S.reviewDraft = null;
          render();
          break;
        case 'evaluate':
          await evaluate(el.dataset.id).then(() => toast('Graded. Review the scores, then approve.'));
          break;
        case 'reevaluate': {
          const reason = await ask({ title: 'Grade again?', body: 'Claude grades this submission again as a new version. The current version, and any trainer review on it, is kept in the audit history.', ok: 'Grade again', input: 'Reason (optional)', placeholder: 'Rubric updated' });
          if (reason === null) return;
          await evaluate(el.dataset.id, true, reason || 'Trainer requested reevaluation').then(() => toast('Graded again.'));
          break;
        }
        case 'eval-all': {
          const ids = C.assignmentRows(dbView()).filter((r) => r.status === 'Submitted' || (r.status === 'Pending Evaluation' && !S.busy[r.submissionId])).map((r) => r.submissionId);
          S.evalAllBusy = true;
          render();
          let ok = 0;
          let fail = 0;
          for (const id of ids) {
            try { await evaluate(id); ok++; } catch (er) { fail++; if (er.code === 'not_granted' || er.code === 'rate_limited' || er.code === 'sampling_disabled') break; }
          }
          S.evalAllBusy = false;
          render();
          toast(`${ok} graded${fail ? `, ${fail} failed. Open each one to see why.` : '.'}`, fail > 0);
          break;
        }
        case 'del-imp':
          S.reviewDraft.tickets[ti].improvements.splice(Number(el.dataset.i), 1);
          render();
          break;
        case 'add-imp': {
          const det = $('imp-det-' + ti).value.trim();
          if (!det) return toast('Write what to improve first.', true);
          S.reviewDraft.tickets[ti].improvements.push({ category: $('imp-cat-' + ti).value, detail: det });
          render();
          break;
        }
        case 'reset-review':
          S.reviewDraft = null;
          render();
          break;
        case 'save-review':
          await saveReview(false);
          break;
        case 'approve-review':
          if (!(me && me.id) && !S.reviewDraft.reviewer.trim()) return toast('Add your name before approving.', true);
          await saveReview(true);
          break;
        case 'history':
          traineeHistory(el.dataset.id, el.dataset.name);
          break;
        case 'clear-dash':
          S.dash = { cohort: '', traineeId: '', quizId: '', trainingDay: '', status: '', from: '', to: '' };
          render();
          break;
        case 'export':
          await doExport(el.dataset.kind, el.dataset.fmt);
          break;
        case 'seed':
          el.disabled = true;
          await seed();
          closeModal();
          S.tab = 'builder';
          render();
          toast('Fictional demo data loaded.');
          break;
        case 'reset': {
          const phrase = await ask({ title: 'Delete all staging data?', body: 'Every quiz, trainee, submission and evaluation on this page is deleted for everyone. Type <b>DELETE STAGING DATA</b> to confirm.', ok: 'Delete everything', danger: true, input: 'Confirmation' });
          if (phrase === null) return;
          if (phrase.trim() !== 'DELETE STAGING DATA') return toast('Nothing was deleted. The confirmation text did not match.', true);
          for (const c of COLLS) for (const x of store[c].slice()) await del(c, x.id);
          S.draft = null;
          S.dirty = false;
          S.reviewDraft = null;
          render();
          toast('Staging data deleted.');
          break;
        }
      }
    } catch (err) {
      toast(err && err.message ? err.message : 'Something went wrong.', true);
      if ('disabled' in el) el.disabled = false;
    }
  });

  function copyText(text, selectEl) {
    const fallback = () => {
      if (selectEl) {
        const r = document.createRange();
        r.selectNodeContents(selectEl);
        if (selectEl.select) selectEl.select();
        else { const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      }
      toast('Copy is blocked here. The text is selected; press Ctrl+C or ⌘C.', true);
    };
    try {
      navigator.clipboard.writeText(text).then(() => toast('Copied.'), fallback);
    } catch (_) {
      fallback();
    }
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && $('modal-root').innerHTML) closeModal(null);
    if (e.key === 'Enter' && e.target.id === 'ask-input') closeModal(e.target.value);
  });

  document.addEventListener('input', (e) => {
    const el = e.target;
    if (el.dataset.bind && S.draft) {
      let v = el.value;
      if (el.dataset.type === 'number') v = v === '' ? '' : Number(v);
      setPath(S.draft, el.dataset.bind, v);
      S.dirty = true;
      if (el.dataset.sum != null) {
        const i = Number(el.dataset.sum);
        const r = S.draft.tickets[i].rubric;
        const sum = C.rubricMax(r);
        const b = $('sum-' + i);
        if (b) {
          b.textContent = `${sum} of ${Number(r.totalPoints) || 0} points allocated`;
          b.className = 'tag ' + (Number(r.totalPoints) > 0 && sum === Number(r.totalPoints) ? 'ok' : 'warn');
        }
      }
      return;
    }
    if (el.dataset.review && S.reviewDraft) {
      const p = el.dataset.review.split('.');
      const rd = S.reviewDraft;
      if (p.length === 1) rd[p[0]] = el.value;
      else if (p[1] === 'notes') rd.tickets[Number(p[0])].trainerNotes = el.value;
      else {
        const t = rd.tickets[Number(p[0])];
        const c = t.criteria[Number(p[1])];
        if (p[2] === 'awarded') {
          c.awarded = el.value === '' ? 0 : Number(el.value);
          C.recomputeTicket(t);
          const s = $('tsum-' + p[0]);
          if (s) s.textContent = `${t.earned} / ${t.possible}`;
          const ev = byId('evaluations', rd.evalId);
          const box = $('draft-box');
          if (box && ev) box.outerHTML = scoreBox('Your edits', C.summarize(rd.tickets, byId('quizzes', ev.quizId)), 'Save or approve to keep', 'draft-box');
        } else c.reason = el.value;
      }
    }
    if (el.id === 'sf-name') {
      const t = store.trainees.find((x) => x.name === el.value);
      if (t && $('sf-email') && !$('sf-email').value) $('sf-email').value = t.email;
    }
  });

  document.addEventListener('change', async (e) => {
    const el = e.target;
    if (el.dataset.filter) {
      setPath(S, el.dataset.filter, el.value);
      if (el.dataset.filter === 'evalFilter') S.selectedSub = null;
      render();
      return;
    }
    if (el.dataset.assign && S.draft) {
      $('assign-box').innerHTML = renderAssign(S.draft);
      return;
    }
    if (el.dataset.act === 'set-status') {
      const status = el.value;
      if (!S.draft.id) { S.draft.status = status; return; }
      if (status === 'Sent') {
        const ok = await ask({ title: 'Mark as Sent?', body: 'This only changes the status label. Nothing is sent to trainees from this page.', ok: 'Mark as Sent' });
        if (ok !== true) { el.value = S.draft.status; return; }
      }
      try {
        const q = clone(byId('quizzes', S.draft.id));
        q.status = status;
        q.updatedAt = new Date().toISOString();
        await put('quizzes', q);
        if (S.dirty) { S.draft.status = status; S.draft.updatedAt = q.updatedAt; } else startDraft(byId('quizzes', q.id));
        render();
        toast('Status set to ' + status + '.');
      } catch (er) {
        toast(er.message, true);
        el.value = S.draft.status;
      }
      return;
    }
    if (el.dataset.act === 'sf-quiz') {
      const q = byId('quizzes', el.value);
      $('sf-body').dataset.quiz = q.id;
      $('sf-body').innerHTML = traineeFormHtml(q, true);
      return;
    }
    if (el.dataset.act === 'import-csv') {
      const file = el.files[0];
      if (!file) return;
      const q = byId('quizzes', $('im-quiz').value);
      try {
        const r = await importRows(q, parseCSV(await file.text()), 'google-form');
        $('im-status').innerHTML = `<div class="notice ok">Read ${plural(r.total, 'response')}: ${r.added} new, ${r.dup} already imported.</div>`;
      } catch (er) {
        $('im-status').innerHTML = `<div class="notice bad">${esc(er.message)}</div>`;
      }
    }
  });

  window.addEventListener('beforeunload', (e) => {
    if (S.dirty) { e.preventDefault(); e.returnValue = ''; }
  });

  // =====================================================================
  // BOOT
  // =====================================================================
  render();
  const use = (n) => (window.claude && window.claude.use ? window.claude.use(n).catch(() => null) : Promise.resolve(null));
  use('db').then((db) => {
    caps.db = db;
    if (!db) { COLLS.forEach((c) => loaded.add(c)); scheduleRender(); return; }
    COLLS.forEach((c) => {
      db.collection(c).onSnapshot(
        (snap) => {
          store[c] = snap.docs.map((d) => ({ ...d.data(), id: d.id }));
          loaded.add(c);
          scheduleRender();
        },
        (err) => { dbError = err; loaded.add(c); scheduleRender(); }
      );
    });
  });
  use('user').then(async (u) => {
    caps.user = u;
    if (u) {
      me = await u.me().catch(() => null);
      const w = await u.can('data.write').catch(() => null);
      if (w === false) readOnly = true;
    }
    scheduleRender();
  });
  use('sample').then((s) => { caps.sample = s; scheduleRender(); });
  use('downloads').then((d) => { caps.downloads = d; scheduleRender(); });
  use('mcp').then((m) => { caps.mcp = m; });
})();
