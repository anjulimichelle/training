/* Ticket Assessment Hub: trainer app (vanilla JS, no build step). */
(function () {
  'use strict';
  const C = window.HubCore;

  // ---------- state ----------
  const S = {
    data: null,
    tab: localGet('tab') || 'builder',
    editQuizId: null,
    draft: null,
    dirty: false,
    validation: null,
    respFilter: { quizId: '', cohort: '', status: '' },
    selectedSubmissionId: null,
    evalFilter: 'open',
    dash: { cohort: '', traineeId: '', quizId: '', trainingDay: '', status: '', from: '', to: '' },
    busy: {},
    reviewDraft: null,
  };

  function localGet(k) {
    try { return localStorage.getItem('tah.' + k); } catch (_) { return null; }
  }
  function localSet(k, v) {
    try { localStorage.setItem('tah.' + k, v); } catch (_) { /* storage unavailable */ }
  }

  // ---------- utils ----------
  const $ = (sel, root) => (root || document).querySelector(sel);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const safeUrl = (u) => (/^(https?:)?\/\//i.test(String(u || '')) || String(u || '').startsWith('/') ? esc(u) : '#');
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const fmtDate = (iso) => {
    if (!iso) return '';
    const d = new Date(iso.length === 10 ? iso + 'T00:00:00' : iso);
    if (isNaN(d)) return esc(iso);
    return d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
  };
  const fmtDateTime = (iso) => {
    if (!iso) return '';
    const d = new Date(iso);
    return isNaN(d) ? esc(iso) : d.toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
  };
  const pctTxt = (p) => (p == null ? '—' : `${p}%`);
  const absUrl = (u) => (u && u.startsWith('/') ? location.origin + u : u);

  function toast(msg, isError) {
    const t = $('#toast');
    t.textContent = msg;
    t.className = 'toast show' + (isError ? ' error' : '');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => (t.className = 'toast'), isError ? 6000 : 2800);
  }

  async function api(method, path, body) {
    const headers = { 'Content-Type': 'application/json' };
    const pass = sessionGet();
    if (pass) headers['x-trainer-passcode'] = pass;
    const res = await fetch(path, { method, headers, body: body ? JSON.stringify(body) : undefined });
    const data = await res.json().catch(() => ({}));
    if (res.status === 401) {
      const p = prompt('Trainer passcode');
      if (p) {
        sessionSet(p);
        return api(method, path, body);
      }
    }
    if (!res.ok) {
      const err = new Error(data.error || `Request failed (${res.status})`);
      err.details = data.errors;
      err.status = res.status;
      throw err;
    }
    return data;
  }
  function sessionGet() {
    try { return sessionStorage.getItem('tah.pass'); } catch (_) { return null; }
  }
  function sessionSet(v) {
    try { sessionStorage.setItem('tah.pass', v); } catch (_) { /* ignore */ }
  }

  async function refresh() {
    S.data = await api('GET', '/api/state');
    renderIntegrations();
  }

  function setPath(obj, path, value) {
    const parts = path.split('.');
    let o = obj;
    for (let i = 0; i < parts.length - 1; i++) o = o[isNaN(parts[i]) ? parts[i] : Number(parts[i])];
    o[parts[parts.length - 1]] = value;
  }

  const trainee = (id) => S.data.trainees.find((t) => t.id === id);
  const quiz = (id) => S.data.quizzes.find((q) => q.id === id);
  const cohorts = () => [...new Set([...S.data.trainees.map((t) => t.cohort), ...S.data.quizzes.map((q) => q.cohort)].filter(Boolean))].sort();
  const currentEval = (subId) => S.data.evaluations.find((e) => e.submissionId === subId && !e.supersededAt);

  function statusBadge(status) {
    const map = { 'Awaiting Submission': '', Submitted: 'badge-blue', 'Pending Evaluation': 'badge-amber', 'Needs Trainer Review': 'badge-amber', Evaluated: 'badge-green', Draft: '', Ready: 'badge-blue', Sent: 'badge-teal', Closed: '' };
    return `<span class="badge ${map[status] || ''}">${esc(status)}</span>`;
  }

  function renderIntegrations() {
    const c = S.data.config;
    const ev = c.evaluator === 'claude' ? `<span class="badge badge-teal" title="Claude grades answers. The API key stays on the server.">Evaluator: Claude</span>` : `<span class="badge badge-amber" title="No Claude API key on the server. A word-overlap heuristic (not AI) is used for staging tests.">Evaluator: offline heuristic</span>`;
    const g = c.googleConfigured ? `<span class="badge ${c.googleWritesApproved ? 'badge-teal' : 'badge-amber'}">Google: ${c.googleWritesApproved ? 'connected' : 'read-only'}</span>` : `<span class="badge" title="Apps Script connector not configured. Staging forms are used.">Google: not connected</span>`;
    $('#integrations').innerHTML = ev + g;
  }

  // ---------- modal ----------
  function openModal(title, body, foot, opts) {
    $('#modal-root').innerHTML = `<div class="modal-backdrop" data-action="modal-bg"><div class="modal" role="dialog" aria-modal="true" aria-label="${esc(title)}" style="${opts && opts.wide ? 'width:min(1100px,100%)' : ''}">
      <div class="modal-head"><h2>${esc(title)}</h2><button class="btn btn-ghost" data-action="close-modal" aria-label="Close">✕</button></div>
      <div class="modal-body">${body}</div>${foot ? `<div class="modal-foot">${foot}</div>` : ''}</div></div>`;
  }
  function closeModal() {
    $('#modal-root').innerHTML = '';
  }

  // ---------- render root ----------
  function render() {
    document.querySelectorAll('.tabs button').forEach((b) => b.setAttribute('aria-selected', b.dataset.tab === S.tab ? 'true' : 'false'));
    const app = $('#app');
    if (!S.data) return;
    if (S.tab === 'builder') app.innerHTML = renderBuilder();
    else if (S.tab === 'responses') app.innerHTML = renderResponses();
    else if (S.tab === 'evaluation') app.innerHTML = renderEvaluation();
    else app.innerHTML = renderDashboard();
  }

  // =====================================================================
  // TAB 1: QUIZ BUILDER
  // =====================================================================
  function startDraft(q) {
    S.editQuizId = q ? q.id : null;
    S.draft = q
      ? clone(q)
      : { id: null, title: '', cohort: cohorts()[0] || '', trainingDay: '', assessmentType: 'Ticket handling', assignedTrainees: [], deadline: '', passingThreshold: '', status: 'Draft', tickets: [C.newTicket()], form: null };
    if (S.draft.passingThreshold == null) S.draft.passingThreshold = '';
    S.dirty = false;
    S.validation = C.validateQuiz(S.draft);
  }

  function renderBuilder() {
    const quizzes = S.data.quizzes.slice().sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
    if (!S.draft && quizzes.length) startDraft(quizzes[0]);
    const list = quizzes.length
      ? quizzes.map((q) => `<button class="list-item ${S.editQuizId === q.id ? 'active' : ''}" data-action="pick-quiz" data-id="${q.id}">
          <div class="title">${esc(q.title || 'Untitled quiz')}</div>
          <div class="small muted">${esc(q.cohort)} · ${q.tickets.length} ticket${q.tickets.length === 1 ? '' : 's'}</div>
          <div style="margin-top:4px">${statusBadge(q.status)} ${q.form ? `<span class="badge badge-blue">${q.form.mode === 'google' ? 'Google Form' : 'Staging form'}</span>` : ''}</div></button>`).join('')
      : `<div class="empty small">No quizzes yet.<br>Create one, or load the fictional demo from Setup.</div>`;
    return `<div class="layout-2">
      <aside>
        <div class="section">
          <div class="section-head"><h3>Quizzes</h3><button class="btn btn-primary btn-sm" data-action="new-quiz">New quiz</button></div>
          <div class="list">${list}</div>
        </div>
        ${renderRosterPanel()}
      </aside>
      <div>${S.draft ? renderEditor() : '<div class="section empty">Select or create a quiz.</div>'}
      ${renderRegistry()}</div>
    </div>`;
  }

  function renderRosterPanel() {
    const rows = S.data.trainees.map((t) => `<tr><td>${esc(t.name)}<div class="small muted">${esc(t.email)}</div></td><td class="small">${esc(t.cohort)}</td><td><button class="btn btn-ghost btn-sm" data-action="del-trainee" data-id="${t.id}" title="Remove">✕</button></td></tr>`).join('');
    return `<details class="block white section" style="padding:0">
      <summary>Trainee roster <span class="badge">${S.data.trainees.length}</span></summary>
      <div class="block-body">
        <p class="small muted" style="margin-top:0">Staging uses fictional trainees only.</p>
        <div class="table-wrap"><table><tbody>${rows || '<tr><td class="muted small">No trainees yet.</td></tr>'}</tbody></table></div>
        <div class="grid" style="margin-top:10px">
          <input type="text" id="tr-name" placeholder="Name">
          <input type="email" id="tr-email" placeholder="Email">
          <input type="text" id="tr-cohort" placeholder="Cohort" list="cohort-list">
          <datalist id="cohort-list">${cohorts().map((c) => `<option value="${esc(c)}">`).join('')}</datalist>
          <button class="btn btn-sm" data-action="add-trainee">Add trainee</button>
        </div>
      </div></details>`;
  }

  function renderAssign(d) {
    const cohortTrainees = S.data.trainees.filter((t) => !d.cohort || t.cohort === d.cohort);
    const others = S.data.trainees.filter((t) => d.cohort && t.cohort !== d.cohort);
    const check = (t) => `<label><input type="checkbox" data-action="toggle-trainee" data-id="${t.id}" ${d.assignedTrainees.includes(t.id) ? 'checked' : ''}> ${esc(t.name)}${t.cohort !== d.cohort ? ` <span class="muted small">(${esc(t.cohort)})</span>` : ''}</label>`;
    return `        <div class="btn-row" style="justify-content:space-between"><span class="field" style="font-size:12px;color:var(--ink-2);font-weight:500">Assigned trainees (${d.assignedTrainees.length})</span>
          <span class="btn-row"><button class="btn btn-sm" data-action="assign-cohort">Assign whole cohort</button><button class="btn btn-ghost btn-sm" data-action="clear-assign">Clear</button></span></div>
        <div class="checks" style="margin-top:6px">${cohortTrainees.map(check).join('') || '<span class="muted small">No trainees in this cohort. Add them in the roster.</span>'}</div>
        ${others.length ? `<details class="block"><summary class="small">Trainees from other cohorts (${others.length})</summary><div class="block-body"><div class="checks">${others.map(check).join('')}</div></div></details>` : ''}`;
  }
  function refreshAssign() {
    const box = document.getElementById('assign-box');
    if (box) box.innerHTML = renderAssign(S.draft);
  }

  function renderEditor() {
    const d = S.draft;
    const v = S.validation || { errors: [], warnings: [] };
    const formPanel = d.form ? renderFormPanel(d) : '';
    return `<div class="section">
      <div class="section-head">
        <div><h2>${esc(d.title || 'New quiz')}</h2><p class="small muted">${d.id ? `Quiz ID <span class="mono">${esc(d.id)}</span>` : 'Not saved yet'} ${S.dirty ? ' · <span class="flag">Unsaved changes</span>' : ''}</p></div>
        <div class="btn-row">
          ${d.id ? `<button class="btn btn-ghost btn-sm" data-action="delete-quiz">Delete</button>` : ''}
          <button class="btn" data-action="save-draft">Save draft</button>
          <button class="btn btn-primary" data-action="preview">Preview &amp; generate form</button>
        </div>
      </div>
      ${formPanel}
      <h4>A. Quiz details</h4>
      <div class="grid grid-3" style="margin-top:10px">
        <label class="field span-2">Quiz title<input type="text" data-bind="title" value="${esc(d.title)}" placeholder="e.g. Day 3 Ticket Assessment"></label>
        <label class="field">Quiz status
          <select data-action="set-status">${C.QUIZ_STATUSES.map((s) => `<option ${d.status === s ? 'selected' : ''} ${!d.form && (s === 'Ready' || s === 'Sent') ? 'disabled' : ''}>${s}</option>`).join('')}</select></label>
        <label class="field">Cohort<input type="text" data-bind="cohort" data-rerender="1" value="${esc(d.cohort)}" list="cohort-list2"><datalist id="cohort-list2">${cohorts().map((c) => `<option value="${esc(c)}">`).join('')}</datalist></label>
        <label class="field">Training day<input type="text" data-bind="trainingDay" value="${esc(d.trainingDay)}" placeholder="e.g. Day 3"></label>
        <label class="field">Assessment type
          <select data-bind="assessmentType">${['Ticket handling', 'Knowledge check', 'Final assessment', 'Practice'].map((x) => `<option ${d.assessmentType === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
        <label class="field">Submission deadline<input type="date" data-bind="deadline" value="${esc(d.deadline)}"></label>
        <label class="field">Passing threshold (%) <span class="muted">optional</span><input type="number" min="1" max="100" data-bind="passingThreshold" value="${esc(d.passingThreshold)}" placeholder="Leave blank for no pass/fail"></label>
        <label class="field">Number of tickets<input type="text" value="${d.tickets.length}" disabled></label>
      </div>
      <div id="assign-box" style="margin-top:12px">${renderAssign(d)}</div>
    </div>

    <div class="section">
      <div class="section-head"><div><h4>B. Tickets</h4><p class="small muted">Each ticket gets its own section in the form, its own rubric and its own reference material.</p></div>
        <button class="btn btn-primary btn-sm" data-action="add-ticket">Add ticket</button></div>
      ${d.tickets.map((t, i) => renderTicket(t, i)).join('')}
      <details class="block"><summary>C. Standardized trainee questions (same for every ticket)</summary><div class="block-body">
        <ol>${C.QUESTIONS.map((q) => `<li>${esc(q.label)} <span class="muted small">(long answer)</span></li>`).join('')}</ol>
        <p class="small muted">The form also asks for the trainee's name and email, and records the cohort, quiz ID and submission time.</p></div></details>
    </div>
    ${renderValidation(v)}`;
  }

  function renderValidation(v) {
    if (!v) return '';
    return (v.errors.length ? `<div class="alert alert-error"><strong>Before you can generate the form:</strong><ul>${v.errors.map((e) => `<li>${esc(e)}</li>`).join('')}</ul></div>` : '') + (v.warnings.length ? `<div class="alert alert-warn"><strong>Check these:</strong><ul>${v.warnings.map((e) => `<li>${esc(e)}</li>`).join('')}</ul></div>` : '');
  }

  function renderFormPanel(d) {
    const f = d.form;
    const url = absUrl(f.formUrl);
    return `<div class="alert ${f.outOfDate ? 'alert-warn' : 'alert-ok'}">
      <div class="btn-row" style="justify-content:space-between">
        <div><strong>${f.mode === 'google' ? 'Google Form' : 'Staging test form'} generated</strong> ${fmtDateTime(f.createdAt)}<br>
        <a href="${safeUrl(f.formUrl)}" target="_blank" rel="noopener" class="mono">${esc(url)}</a>
        ${f.responseSheetUrl ? `<br><a href="${safeUrl(f.responseSheetUrl)}" target="_blank" rel="noopener">Response sheet</a>` : ''}
        ${f.outOfDate ? '<br>The quiz changed after the form was generated. Regenerate it so trainees see the current tickets.' : ''}</div>
        <div class="btn-row"><button class="btn btn-sm" data-action="copy-link" data-url="${esc(url)}">Copy link</button>
        <button class="btn btn-sm" data-action="send-dialog">Send…</button></div>
      </div></div>`;
  }

  function renderTicket(t, i) {
    const d = S.draft;
    const p = `tickets.${i}`;
    const r = t.rubric;
    const sum = C.rubricMax(r);
    const sumCls = Number(r.totalPoints) > 0 && sum === Number(r.totalPoints) ? 'sum-ok' : 'sum-bad';
    const ta = (path, val, ph, rows) => `<textarea data-bind="${p}.${path}" placeholder="${esc(ph || '')}" rows="${rows || 3}">${esc(val)}</textarea>`;
    const attach = (t.attachments || []).map((a, ai) => `<div class="btn-row small"><a href="${safeUrl(a.url)}" target="_blank" rel="noopener">${esc(a.name || a.url)}</a><button class="btn btn-ghost btn-sm" data-action="del-attach" data-t="${i}" data-a="${ai}">✕</button></div>`).join('');
    const crit = r.criteria
      .map(
        (c, ci) => `<tr>
        <td style="width:26%"><input type="text" data-bind="${p}.rubric.criteria.${ci}.name" value="${esc(c.name)}" placeholder="Criterion"></td>
        <td style="width:18%"><select data-bind="${p}.rubric.criteria.${ci}.category">${C.CRITERION_CATEGORIES.map((k) => `<option value="${k.key}" ${c.category === k.key ? 'selected' : ''}>${esc(k.label)}</option>`).join('')}</select></td>
        <td style="width:9%"><input type="number" min="0" step="0.5" data-bind="${p}.rubric.criteria.${ci}.maxPoints" data-type="number" data-sum="${i}" value="${esc(c.maxPoints)}"></td>
        <td><textarea rows="1" data-bind="${p}.rubric.criteria.${ci}.guidance" placeholder="What earns full points">${esc(c.guidance)}</textarea>
          ${(c.deductions || []).map((dd, di) => `<div class="btn-row" style="margin-top:4px;flex-wrap:nowrap"><input type="text" data-bind="${p}.rubric.criteria.${ci}.deductions.${di}.description" value="${esc(dd.description)}" placeholder="Deduction reason"><input type="number" min="0" step="0.5" style="width:80px" data-bind="${p}.rubric.criteria.${ci}.deductions.${di}.points" data-type="number" value="${esc(dd.points)}" title="Points to deduct"><button class="btn btn-ghost btn-sm" data-action="del-deduction" data-t="${i}" data-c="${ci}" data-d="${di}">✕</button></div>`).join('')}
          <button class="linkish small" data-action="add-deduction" data-t="${i}" data-c="${ci}">+ Add point deduction</button></td>
        <td style="width:40px"><button class="btn btn-ghost btn-sm" data-action="del-criterion" data-t="${i}" data-c="${ci}" title="Remove criterion">✕</button></td></tr>`
      )
      .join('');
    return `<div class="ticket-card" id="ticket-${i}">
      <div class="ticket-head"><span class="ticket-title">Ticket ${i + 1}</span>
        <span class="small muted">${esc(t.url || 'No URL yet')}</span>
        <span class="btn-row"><button class="btn btn-ghost btn-sm" data-action="move-ticket" data-t="${i}" data-dir="-1" ${i === 0 ? 'disabled' : ''} title="Move up">↑</button>
        <button class="btn btn-ghost btn-sm" data-action="move-ticket" data-t="${i}" data-dir="1" ${i === d.tickets.length - 1 ? 'disabled' : ''} title="Move down">↓</button>
        <button class="btn btn-ghost btn-sm" data-action="dup-ticket" data-t="${i}" title="Duplicate rubric into a new ticket">Duplicate</button>
        <button class="btn btn-danger btn-sm" data-action="del-ticket" data-t="${i}">Remove</button></span></div>

      <details class="block" open><summary>Ticket reference <span class="small muted">(trainees see this)</span></summary><div class="block-body grid grid-2">
        <label class="field span-2">Ticket URL<input type="url" data-bind="${p}.url" value="${esc(t.url)}" placeholder="https://…"></label>
        <label class="field">Ticket description (optional)${ta('description', t.description, 'Short summary shown under the ticket link', 2)}</label>
        <label class="field">Additional context or instructions${ta('context', t.context, 'Anything the trainee should know', 2)}</label>
        <div class="span-2"><div class="small" style="color:var(--ink-2);font-weight:500">Screenshots and attachments (optional)</div>${attach}
          <div class="btn-row" style="margin-top:6px"><input type="url" id="att-url-${i}" placeholder="Link to a screenshot or file" style="flex:1"><button class="btn btn-sm" data-action="add-attach-link" data-t="${i}">Add link</button>
          <label class="btn btn-sm">Upload<input type="file" accept="image/*,application/pdf" data-action="upload" data-t="${i}" hidden></label></div></div>
      </div></details>

      <details class="block"><summary>Scoring rubric <span class="badge ${sumCls === 'sum-ok' ? 'badge-green' : 'badge-amber'}" id="sum-${i}">${sum} of ${Number(r.totalPoints) || 0} points allocated</span></summary><div class="block-body">
        <div class="private-note">Trainer only. Never shown to trainees or written into the form.</div>
        <div class="grid grid-2">
          <label class="field">Total available points<input type="number" min="0" step="0.5" data-bind="${p}.rubric.totalPoints" data-type="number" data-sum="${i}" value="${esc(r.totalPoints)}"></label><div></div>
          <label class="field">Expected customer intent${ta('rubric.expectedIntent', r.expectedIntent, 'What C actually wants')}</label>
          <label class="field">Required CRM actions <span class="muted">(one per line)</span>${ta('rubric.requiredCrmActions', r.requiredCrmActions, 'Add internal note\nIssue refund…')}</label>
          <label class="field">Required customer resolution${ta('rubric.requiredResolution', r.requiredResolution)}</label>
          <label class="field">Critical errors${ta('rubric.criticalErrors', r.criticalErrors, 'Mistakes that must be flagged')}</label>
          <label class="field">Expected communication to C${ta('rubric.expectedC', r.expectedC)}</label>
          <label class="field">Expected communication to CP${ta('rubric.expectedCP', r.expectedCP, 'Write N/A if no CP message is expected')}</label>
          <label class="field">Acceptable alternative responses${ta('rubric.acceptableAlternatives', r.acceptableAlternatives)}</label>
          <label class="field">Automatic fail conditions <span class="muted">(optional)</span>${ta('rubric.autoFailConditions', r.autoFailConditions)}</label>
        </div>
        <div class="section-head" style="margin:16px 0 6px"><strong>Criteria and points</strong>
          <span class="btn-row"><button class="btn btn-sm" data-action="suggest-criteria" data-t="${i}" title="Adds the five standard criteria split across the total. Edit freely.">Insert suggested criteria</button><button class="btn btn-sm" data-action="add-criterion" data-t="${i}">Add criterion</button></span></div>
        <div class="table-wrap"><table class="crit-table"><thead><tr><th>Criterion</th><th>Assesses</th><th>Max points</th><th>Guidance and point deductions</th><th></th></tr></thead><tbody>${crit || '<tr><td colspan="5" class="muted small">No criteria yet.</td></tr>'}</tbody></table></div>
      </div></details>

      <details class="block"><summary>Reference materials <span class="small muted">(AI evaluation only)</span></summary><div class="block-body">
        <div class="private-note">Trainer only. Used by the evaluator, never shown to trainees.</div>
        <div class="grid grid-2">
          <label class="field">Policy context${ta('references.policyContext', t.references.policyContext)}</label>
          <label class="field">Knowledge Base references${ta('references.kbReferences', t.references.kbReferences)}</label>
          <label class="field">Expected resolution${ta('references.expectedResolution', t.references.expectedResolution)}</label>
          <label class="field">Additional evaluation instructions${ta('references.evaluatorInstructions', t.references.evaluatorInstructions)}</label>
          <label class="field">Sample response to C${ta('references.sampleC', t.references.sampleC)}</label>
          <label class="field">Sample message to CP${ta('references.sampleCP', t.references.sampleCP)}</label>
        </div></div></details>
    </div>`;
  }

  function renderRegistry() {
    const rows = S.data.quizzes
      .map((q) => `<tr><td>${esc(q.title)}</td><td>${q.form ? `<a href="${safeUrl(q.form.formUrl)}" target="_blank" rel="noopener">${q.form.mode === 'google' ? 'Google Form' : 'Staging form'}</a>` : '<span class="muted">Not generated</span>'}</td><td>${esc(q.cohort)}</td><td class="small">${q.assignedTrainees.map((id) => esc((trainee(id) || {}).name || id)).join(', ')}</td><td>${fmtDate(q.createdAt)}</td><td>${fmtDate(q.deadline)}</td><td>${statusBadge(q.status)}</td></tr>`)
      .join('');
    return `<details class="block white section" style="padding:0"><summary>Quiz registry <span class="badge">${S.data.quizzes.length}</span></summary><div class="block-body table-wrap">
      <table><thead><tr><th>Quiz title</th><th>Form</th><th>Cohort</th><th>Assigned trainees</th><th>Created</th><th>Deadline</th><th>Status</th></tr></thead><tbody>${rows || '<tr><td colspan="7" class="muted">No quizzes.</td></tr>'}</tbody></table></div></details>`;
  }

  function renderPreview() {
    const d = S.draft;
    const tv = C.traineeView(d);
    const v = C.validateQuiz(d);
    const c = S.data.config;
    const form = `<div class="form-preview">
      <div class="gf-card head"><div style="font-size:24px">${esc(tv.title)}</div>
        <div class="small muted" style="margin-top:6px">Cohort: ${esc(tv.cohort)} · ${esc(tv.trainingDay)} · ${esc(tv.assessmentType)}${tv.deadline ? ` · Due ${fmtDate(tv.deadline)}` : ''}<br>Quiz ID: ${esc(tv.id || '(assigned on save)')}</div></div>
      <div class="gf-card"><div class="gf-q">Full name *</div><div class="gf-input">Short answer text</div></div>
      <div class="gf-card"><div class="gf-q">Email *</div><div class="gf-input">Short answer text</div></div>
      <div class="gf-card"><div class="gf-q">Cohort *</div><div class="gf-input">${esc(tv.cohort)}</div></div>
      ${tv.tickets.map((t) => `<div class="gf-section">Section ${t.number + 1} of ${tv.tickets.length + 1}</div>
        <div class="gf-card head" style="border-top-left-radius:0"><div style="font-size:20px">Ticket ${t.number}</div>
          <div style="margin-top:6px">Ticket reference: <a href="${safeUrl(t.url)}" target="_blank" rel="noopener">${esc(t.url || '(missing URL)')}</a></div>
          ${t.description ? `<div class="small" style="margin-top:4px">${esc(t.description)}</div>` : ''}${t.context ? `<div class="small muted" style="margin-top:4px">${esc(t.context)}</div>` : ''}
          ${t.attachments.map((a) => `<div class="small">Attachment: <a href="${safeUrl(a.url)}" target="_blank" rel="noopener">${esc(a.name || a.url)}</a></div>`).join('')}</div>
        ${tv.questions.map((q) => `<div class="gf-card"><div class="gf-q">${esc(q.label)} *</div><div class="gf-input">Long answer text</div></div>`).join('')}`).join('')}
    </div>`;
    const rubric = d.tickets
      .map((t, i) => `<details class="block"><summary>Ticket ${i + 1} rubric (${C.rubricMax(t.rubric)} of ${t.rubric.totalPoints} points)</summary><div class="block-body">
        <table><thead><tr><th>Criterion</th><th>Assesses</th><th class="num">Points</th><th>Deductions</th></tr></thead><tbody>${t.rubric.criteria.map((c) => `<tr><td>${esc(c.name)}</td><td class="small">${esc((C.CRITERION_CATEGORIES.find((k) => k.key === c.category) || {}).label)}</td><td class="num">${esc(c.maxPoints)}</td><td class="small">${(c.deductions || []).map((x) => `${esc(x.description)} (−${esc(x.points)})`).join('<br>')}</td></tr>`).join('')}</tbody></table>
        <dl class="kv" style="margin-top:10px"><dt>Expected intent</dt><dd>${esc(t.rubric.expectedIntent) || '—'}</dd><dt>Required CRM actions</dt><dd style="white-space:pre-wrap">${esc(t.rubric.requiredCrmActions) || '—'}</dd><dt>Required resolution</dt><dd>${esc(t.rubric.requiredResolution) || '—'}</dd><dt>Critical errors</dt><dd>${esc(t.rubric.criticalErrors) || '—'}</dd><dt>Auto-fail</dt><dd>${esc(t.rubric.autoFailConditions) || '—'}</dd></dl></div></details>`)
      .join('');
    const googleOk = c.googleConfigured && c.googleWritesApproved;
    const googleWhy = !c.googleConfigured ? 'Google is not connected (see README, "Connecting Google").' : !c.googleWritesApproved ? 'External Google actions are not approved on the server (GOOGLE_EXTERNAL_ACTIONS_APPROVED).' : '';
    openModal(
      'Preview assessment',
      `${renderValidation(v)}
      <div class="grid grid-2" style="align-items:start">
        <div><h3 style="margin-bottom:8px">What trainees will see</h3>${form}</div>
        <div><h3 style="margin-bottom:8px">Trainer-only rubric</h3><div class="private-note">Not included in the form.</div>${rubric}
          <h3 style="margin:16px 0 8px">Ticket links</h3><ol>${d.tickets.map((t) => `<li><a href="${safeUrl(t.url)}" target="_blank" rel="noopener">${esc(t.url || '(missing)')}</a></li>`).join('')}</ol></div>
      </div>`,
      `<label class="small" style="margin-right:auto;display:flex;gap:6px;align-items:center"><input type="checkbox" id="approve-check" ${v.errors.length ? 'disabled' : ''}> I reviewed the links, questions and rubric and approve generating this form.</label>
       <button class="btn" data-action="close-modal">Back to editing</button>
       <button class="btn" data-action="generate" data-target="google" ${googleOk && !v.errors.length ? '' : 'disabled'} title="${esc(googleWhy)}">Generate Google Form</button>
       <button class="btn btn-primary" data-action="generate" data-target="staging" ${v.errors.length ? 'disabled' : ''}>Generate staging form</button>`,
      { wide: true }
    );
  }

  async function saveDraft(silent) {
    const out = await api('POST', '/api/quizzes', S.draft);
    await refresh();
    startDraft(quiz(out.quiz.id));
    S.validation = out.validation;
    if (!silent) toast('Draft saved.');
    render();
    return out;
  }

  // =====================================================================
  // TAB 2: QUIZ RESPONSES
  // =====================================================================
  function renderResponses() {
    const f = S.respFilter;
    const rows = S.data.rows.filter((r) => (!f.quizId || r.quizId === f.quizId) && (!f.cohort || r.cohort === f.cohort) && (!f.status || r.status === f.status));
    const googleQuizzes = S.data.quizzes.filter((q) => q.form && q.form.mode === 'google');
    const stagingQuizzes = S.data.quizzes.filter((q) => q.form && q.form.mode === 'staging');
    const body = rows
      .map((r) => `<tr class="${r.submissionId ? 'clickable' : ''}" ${r.submissionId ? `data-action="open-sub" data-id="${r.submissionId}"` : ''}>
        <td>${r.submissionId ? `<button class="linkish" data-action="open-sub" data-id="${r.submissionId}">${esc(r.traineeName)}</button>` : esc(r.traineeName)}${r.unassigned ? ' <span class="badge badge-amber" title="Submitted but not assigned to this quiz">Not assigned</span>' : ''}<div class="small muted">${esc(r.traineeEmail)}</div></td>
        <td>${esc(r.cohort)}</td><td>${esc(r.quizTitle)}</td><td>${r.submittedAt ? fmtDateTime(r.submittedAt) : '<span class="muted">—</span>'}</td>
        <td class="num">${r.submissionId ? `${r.ticketsAnswered} of ${r.ticketsTotal}` : '—'}</td><td>${statusBadge(r.status)}</td></tr>`)
      .join('');
    return `<div class="section">
      <div class="section-head"><div><h2>Quiz responses</h2><p class="small muted">Raw submissions are stored as received and never edited by evaluation.</p></div>
        <div class="btn-row">
          ${googleQuizzes.map((q) => `<button class="btn btn-sm" data-action="sync-google" data-id="${q.id}">Sync Google responses: ${esc(q.title)}</button>`).join('')}
          ${stagingQuizzes.length ? `<select id="sample-quiz">${stagingQuizzes.map((q) => `<option value="${q.id}">${esc(q.title)}</option>`).join('')}</select><button class="btn btn-sm" data-action="add-samples">Add fictional sample submissions</button>` : ''}
        </div></div>
      <div class="grid grid-3" style="margin-bottom:12px">
        <select data-filter="respFilter.quizId"><option value="">All quizzes</option>${S.data.quizzes.map((q) => `<option value="${q.id}" ${f.quizId === q.id ? 'selected' : ''}>${esc(q.title)}</option>`).join('')}</select>
        <select data-filter="respFilter.cohort"><option value="">All cohorts</option>${cohorts().map((c) => `<option ${f.cohort === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>
        <select data-filter="respFilter.status"><option value="">All statuses</option>${C.EVAL_STATUSES.map((s) => `<option ${f.status === s ? 'selected' : ''}>${s}</option>`).join('')}</select>
      </div>
      <div class="table-wrap"><table><thead><tr><th>Trainee</th><th>Cohort</th><th>Quiz</th><th>Submitted</th><th class="num">Tickets answered</th><th>Evaluation status</th></tr></thead>
      <tbody>${body || '<tr><td colspan="6" class="empty">No responses match these filters.</td></tr>'}</tbody></table></div></div>`;
  }

  function submissionAnswersHtml(sub) {
    const q = quiz(sub.quizId);
    return sub.answers
      .map((a) => `<details class="block" open><summary>Ticket ${a.ticketNumber} <a class="small" href="${safeUrl(a.ticketUrl)}" target="_blank" rel="noopener" onclick="event.stopPropagation()">${esc(a.ticketUrl)}</a></summary><div class="block-body">
        ${C.QUESTIONS.map((qq) => `<div class="qlabel">${esc(qq.label)}</div><div class="answer">${esc(a[qq.key]) || '<span class="muted">(blank)</span>'}</div>`).join('')}</div></details>`)
      .join('') + (q && q.tickets.length !== sub.answers.length ? `<div class="alert alert-warn">This quiz now has ${q.tickets.length} tickets; the submission was made against ${sub.answers.length}.</div>` : '');
  }

  function openSubmission(id) {
    const s = S.data.submissions.find((x) => x.id === id);
    if (!s) return;
    const q = quiz(s.quizId) || {};
    const ev = currentEval(s.id);
    openModal(
      `${s.traineeName}: ${q.title || ''}`,
      `<dl class="kv"><dt>Email</dt><dd>${esc(s.traineeEmail)}</dd><dt>Cohort</dt><dd>${esc(s.cohort)}</dd><dt>Submitted</dt><dd>${fmtDateTime(s.submittedAt)}</dd><dt>Source</dt><dd>${esc(s.source)}${s.externalResponseId ? ` (response ${esc(s.externalResponseId)})` : ''}</dd><dt>Submission ID</dt><dd class="mono">${esc(s.id)}</dd><dt>Integrity hash</dt><dd class="mono">${esc(s.rawHash.slice(0, 16))}…</dd></dl>
      ${s.lastError ? `<div class="alert alert-error">Last evaluation attempt failed: ${esc(s.lastError.message)}</div>` : ''}
      <div style="margin-top:12px">${submissionAnswersHtml(s)}</div>`,
      `<button class="btn" data-action="close-modal">Close</button><button class="btn btn-primary" data-action="go-eval" data-id="${s.id}">${ev ? 'Open evaluation' : 'Evaluate'}</button>`
    );
  }

  // =====================================================================
  // TAB 3: EVALUATION & SCORING
  // =====================================================================
  function renderEvaluation() {
    const subRows = S.data.rows.filter((r) => r.submissionId);
    const filt = S.evalFilter;
    const shown = subRows.filter((r) => filt === 'all' || (filt === 'open' ? r.status !== 'Evaluated' : r.status === 'Evaluated'));
    if (!S.selectedSubmissionId && shown.length) S.selectedSubmissionId = shown[0].submissionId;
    const pendingCount = subRows.filter((r) => r.status === 'Submitted' || r.status === 'Pending Evaluation').length;
    const list = shown
      .map((r) => `<button class="list-item ${S.selectedSubmissionId === r.submissionId ? 'active' : ''}" data-action="pick-sub" data-id="${r.submissionId}">
        <div class="title">${esc(r.traineeName)}</div><div class="small muted">${esc(r.quizTitle)}</div>
        <div style="margin-top:4px">${statusBadge(r.status)} ${r.percent != null ? `<span class="badge">${r.percent}%${r.scoreSource === 'ai' ? ' AI' : ''}</span>` : ''}</div></button>`)
      .join('');
    return `<div class="layout-2"><aside class="section">
      <div class="section-head"><h3>Submissions</h3></div>
      <select data-filter="evalFilter" style="margin-bottom:10px"><option value="open" ${filt === 'open' ? 'selected' : ''}>Not yet final</option><option value="done" ${filt === 'done' ? 'selected' : ''}>Evaluated (final)</option><option value="all" ${filt === 'all' ? 'selected' : ''}>All</option></select>
      <button class="btn btn-primary btn-sm" style="width:100%;justify-content:center;margin-bottom:10px" data-action="eval-all" ${pendingCount ? '' : 'disabled'}>${S.busy.evalAll ? '<span class="spinner"></span> Evaluating…' : `Evaluate all new submissions (${pendingCount})`}</button>
      <div class="list">${list || '<div class="empty small">Nothing here.</div>'}</div></aside>
      <div>${S.selectedSubmissionId ? renderEvalDetail(S.selectedSubmissionId) : '<div class="section empty">Select a submission.</div>'}</div></div>`;
  }

  function scoreBox(label, o, sub) {
    if (!o) return `<div class="score-box"><div class="small muted">${label}</div><div class="score-big muted">—</div><div class="small muted">${sub || ''}</div></div>`;
    return `<div class="score-box"><div class="small muted">${label}</div><div class="score-big">${o.earned} / ${o.possible}</div><div>${pctTxt(o.percent)}</div><div class="bar"><span style="width:${Math.max(0, Math.min(100, o.percent || 0))}%"></span></div><div class="small muted" style="margin-top:4px">${sub || ''}</div></div>`;
  }

  function resultBadge(result) {
    if (result === 'Pass') return '<span class="badge badge-green">Pass</span>';
    if (result === 'Fail') return '<span class="badge badge-red">Fail</span>';
    if (result === 'Needs retake') return '<span class="badge badge-amber">Needs retake</span>';
    return '<span class="badge" title="No passing threshold configured for this quiz">No pass/fail configured</span>';
  }

  function renderEvalDetail(subId) {
    const s = S.data.submissions.find((x) => x.id === subId);
    if (!s) return '<div class="section empty">Submission not found.</div>';
    const q = quiz(s.quizId);
    const ev = currentEval(s.id);
    const versions = S.data.evaluations.filter((e) => e.submissionId === s.id).sort((a, b) => a.version - b.version);
    const head = `<div class="section-head"><div><h2>${esc(s.traineeName)}</h2><p class="small muted">${esc(q.title)} · ${esc(s.cohort)} · submitted ${fmtDateTime(s.submittedAt)}</p></div>
      <div class="btn-row"><button class="btn btn-sm" data-action="open-sub" data-id="${s.id}">View raw submission</button>
      ${S.busy[s.id] ? '<span class="badge badge-amber"><span class="spinner"></span> Evaluating</span>' : ev ? `<button class="btn btn-sm" data-action="reevaluate" data-id="${s.id}">Request reevaluation</button>` : `<button class="btn btn-primary" data-action="evaluate" data-id="${s.id}">Run evaluation</button>`}</div></div>`;
    if (!ev) {
      return `<div class="section">${head}
        ${s.lastError ? `<div class="alert alert-error">Last attempt failed: ${esc(s.lastError.message)}<br>The raw answers are safe. Fix the cause and run the evaluation again.</div>` : ''}
        <div class="alert alert-info">Evaluator: ${S.data.config.evaluator === 'claude' ? `Claude (${esc(S.data.config.claudeModel)}) grades each ticket against your rubric and references.` : 'Offline heuristic (word overlap, not AI). Add ANTHROPIC_API_KEY on the server to grade with Claude.'}</div>
        ${submissionAnswersHtml(s)}</div>`;
    }
    const rd = S.reviewDraft && S.reviewDraft.evalId === ev.id ? S.reviewDraft : (S.reviewDraft = makeReviewDraft(ev));
    const ai = ev.ai;
    const trApproved = ev.trainer && ev.trainer.approvedAt;
    const draftOverall = C.summarize(rd.tickets.map((t) => C.recomputeTicket(t)), q);
    const changed = JSON.stringify(rd.tickets.map((t) => t.criteria.map((c) => [c.awarded, c.reason]))) !== JSON.stringify((ev.trainer || ai).tickets.map((t) => t.criteria.map((c) => [c.awarded, c.reason])));
    const overall = `<div class="section">${head}
      ${ai.evaluator === 'heuristic' ? '<div class="alert alert-warn">Scored by the offline heuristic (word overlap, not AI). Critical errors and auto-fail conditions were not checked. Review every score.</div>' : ''}
      <div class="score-pair">
        ${scoreBox('AI evaluation (original, read-only)', ai.overall, `${esc(ai.evaluator === 'claude' ? 'Claude ' + (ai.model || '') : ai.model)} · v${ev.version} · ${fmtDateTime(ev.createdAt)}`)}
        ${scoreBox('Trainer final', trApproved ? ev.trainer.overall : null, trApproved ? `Approved by ${esc(ev.trainer.approvedBy)} · ${fmtDateTime(ev.trainer.approvedAt)}` : 'Not approved yet')}
        <div id="draft-box">${scoreBox('Your edits (unsaved preview)', changed ? draftOverall : null, changed ? 'Save or approve to keep' : 'No changes')}</div>
      </div>
      <div class="grid grid-3" style="margin-top:14px">
        <div><h4>Result</h4><div style="margin-top:6px">${resultBadge(trApproved ? ev.trainer.overrideResult || ev.trainer.overall.result : ai.overall.result)} ${ai.overall.autoFail ? '<span class="badge badge-red">Auto-fail condition met</span>' : ''}</div>
          ${ai.overall.passingThreshold ? `<div class="small muted" style="margin-top:4px">Passing threshold ${ai.overall.passingThreshold}%</div>` : ''}</div>
        <div><h4>Main strengths</h4>${ai.overall.strengths.length ? `<ul class="clean small">${ai.overall.strengths.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : '<p class="small muted">None identified.</p>'}</div>
        <div><h4>Recurring weaknesses and coaching focus</h4>${ai.overall.recurringWeaknesses.length ? `<ul class="clean small">${ai.overall.recurringWeaknesses.map((x) => `<li>${esc(x.category)} <span class="muted">(${x.count} ticket${x.count === 1 ? '' : 's'})</span></li>`).join('')}</ul>` : '<p class="small muted">None identified.</p>'}
          ${ai.overall.coachingFocus.length ? `<p class="small"><strong>Focus:</strong> ${ai.overall.coachingFocus.map(esc).join(', ')}</p>` : ''}</div>
      </div></div>`;

    const ticketHtml = ai.tickets
      .map((t, ti) => {
        const rt = rd.tickets[ti];
        const ans = s.answers[ti] || {};
        return `<div class="section"><div class="section-head"><div><h3>Ticket ${t.ticketNumber} <a class="small" href="${safeUrl(t.ticketUrl)}" target="_blank" rel="noopener">${esc(t.ticketUrl)}</a></h3>
          <p class="small muted">AI ${t.earned} / ${t.possible} (${pctTxt(t.percent)}) · Trainer draft <strong id="tsum-${ti}">${rt.earned} / ${rt.possible}</strong></p></div>
          <div class="btn-row">${t.needsTrainerReview ? '<span class="badge badge-amber">Flagged for trainer review</span>' : ''}${t.autoFail ? '<span class="badge badge-red">Auto-fail</span>' : ''}</div></div>
          ${t.criticalErrors.length ? `<div class="alert alert-error"><strong>Critical errors</strong><ul>${t.criticalErrors.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''}
          ${t.autoFail ? `<div class="alert alert-error">Automatic fail condition: ${esc(t.autoFailReason)}</div>` : ''}
          ${t.reviewNotes ? `<div class="small muted" style="margin-bottom:8px">Evaluator note: ${esc(t.reviewNotes)}</div>` : ''}
          <div class="table-wrap"><table><thead><tr><th>Criterion</th><th class="num">Max</th><th class="num">AI</th><th>AI reason for deduction</th><th style="width:90px">Trainer</th><th>Trainer reason</th></tr></thead><tbody>
          ${t.criteria.map((c, ci) => `<tr><td>${esc(c.name)}${c.needsReview ? ' <span class="flag" title="Evaluator could not judge this with the information given">⚑ review</span>' : ''}</td><td class="num">${c.maxPoints}</td><td class="num">${c.awarded}</td><td class="small">${esc(c.reason) || '<span class="muted">No deduction</span>'}</td>
            <td><input type="number" min="0" max="${c.maxPoints}" step="0.5" data-review="${ti}.${ci}.awarded" value="${esc(rt.criteria[ci].awarded)}"></td>
            <td><textarea rows="1" data-review="${ti}.${ci}.reason" placeholder="Reason (optional)">${esc(rt.criteria[ci].reason)}</textarea></td></tr>`).join('')}
          </tbody></table></div>
          <div class="grid grid-3" style="margin-top:12px">
            <div><h4>What the trainee did well</h4>${t.strengths.length ? `<ul class="clean small">${t.strengths.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : '<p class="small muted">None identified.</p>'}</div>
            <div><h4>Areas for improvement</h4>${rt.improvements.length ? `<ul class="clean small">${rt.improvements.map((x, ii) => `<li><span class="badge">${esc(x.category)}</span> ${esc(x.detail)} <button class="linkish small" data-action="del-improvement" data-t="${ti}" data-i="${ii}">remove</button></li>`).join('')}</ul>` : '<p class="small muted">None.</p>'}
              <details class="block"><summary class="small">Add an area for improvement</summary><div class="block-body grid">
                <select id="imp-cat-${ti}">${C.IMPROVEMENT_CATEGORIES.map((x) => `<option>${esc(x)}</option>`).join('')}</select><input type="text" id="imp-det-${ti}" placeholder="Detail"><button class="btn btn-sm" data-action="add-improvement" data-t="${ti}">Add</button></div></details></div>
            <div><h4>Suggested coaching</h4>${t.coaching.length ? `<ul class="clean small">${t.coaching.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : '<p class="small muted">None.</p>'}</div>
          </div>
          <label class="field" style="margin-top:10px">Trainer notes for this ticket<textarea rows="2" data-review="${ti}.notes">${esc(rt.trainerNotes || '')}</textarea></label>
          <details class="block"><summary class="small">Trainee answers for Ticket ${t.ticketNumber}</summary><div class="block-body">${C.QUESTIONS.map((qq) => `<div class="qlabel">${esc(qq.label)}</div><div class="answer">${esc(ans[qq.key]) || '<span class="muted">(blank)</span>'}</div>`).join('')}</div></details>
        </div>`;
      })
      .join('');

    const review = `<div class="section"><h3>Trainer review</h3>
      <div class="grid grid-2" style="margin-top:10px">
        <label class="field">Trainer comments<textarea rows="3" data-review="comments">${esc(rd.comments)}</textarea></label>
        <div class="grid"><label class="field">Override overall result<select data-review="overrideResult">${['', 'Pass', 'Fail', 'Needs retake'].map((x) => `<option value="${x}" ${rd.overrideResult === x ? 'selected' : ''}>${x || 'No override (use calculated result)'}</option>`).join('')}</select></label>
          <label class="field">Reviewer name<input type="text" data-review="reviewer" value="${esc(rd.reviewer)}" placeholder="Your name"></label></div>
      </div>
      <div class="btn-row" style="margin-top:12px;justify-content:flex-end">
        <button class="btn" data-action="reset-review">Reset to ${ev.trainer ? 'saved trainer version' : 'AI scores'}</button>
        <button class="btn" data-action="save-review">Save edits</button>
        <button class="btn btn-primary" data-action="approve-review" title="Approves the scores in the Trainer column. Unedited rows keep the AI score.">Approve final evaluation</button>
      </div>
      <p class="small muted">The original AI evaluation is kept unchanged. Your edits are stored separately with timestamps.</p></div>`;

    const hist = versions.length > 1 || (ev.history && ev.history.length)
      ? `<details class="block white section" style="padding:0"><summary>Audit history</summary><div class="block-body small">
          ${versions.map((v) => `<div>v${v.version} created ${fmtDateTime(v.createdAt)} by ${esc(v.ai.evaluator)}: AI ${v.ai.overall.earned}/${v.ai.overall.possible}${v.supersededAt ? ` · superseded ${fmtDateTime(v.supersededAt)} (${esc(v.supersededReason)})` : ' · current'}${v.trainer && v.trainer.approvedAt ? ` · trainer final ${v.trainer.overall.earned}/${v.trainer.overall.possible}` : ''}</div>`).join('')}
          ${(ev.history || []).map((h) => `<div>${fmtDateTime(h.at)} · ${esc(h.by)} ${esc(h.action)} (score ${h.after.earned})</div>`).join('')}</div></details>`
      : '';
    return overall + ticketHtml + review + hist;
  }

  function makeReviewDraft(ev) {
    const src = ev.trainer || ev.ai;
    return { evalId: ev.id, tickets: clone(src.tickets).map((t) => ({ ...t, trainerNotes: t.trainerNotes || '' })), comments: ev.trainer ? ev.trainer.comments || '' : '', overrideResult: ev.trainer ? ev.trainer.overrideResult || '' : '', reviewer: localGet('reviewer') || '' };
  }

  async function submitReview(approve) {
    const rd = S.reviewDraft;
    if (approve && !rd.reviewer.trim()) return toast('Add your name as reviewer before approving.', true);
    localSet('reviewer', rd.reviewer);
    await api('POST', `/api/evaluations/${rd.evalId}/review`, {
      tickets: rd.tickets.map((t) => ({ criteria: t.criteria.map((c) => ({ awarded: c.awarded, reason: c.reason })), improvements: t.improvements, trainerNotes: t.trainerNotes })),
      comments: rd.comments,
      overrideResult: rd.overrideResult,
      reviewer: rd.reviewer,
      approve,
    });
    S.reviewDraft = null;
    await refresh();
    toast(approve ? 'Evaluation approved as final.' : 'Edits saved.');
    render();
  }

  async function runEvaluation(subId, reevaluate, reason) {
    S.busy[subId] = true;
    render();
    try {
      await api('POST', `/api/submissions/${subId}/evaluate`, { reevaluate: !!reevaluate, reason });
      S.reviewDraft = null;
      toast('Evaluation complete.');
    } catch (e) {
      toast(e.message, true);
    } finally {
      delete S.busy[subId];
      await refresh();
      render();
    }
  }

  // =====================================================================
  // TAB 4: PERFORMANCE DASHBOARD
  // =====================================================================
  function dashRows() {
    const f = S.dash;
    return S.data.rows.filter((r) => {
      if (f.cohort && r.cohort !== f.cohort) return false;
      if (f.traineeId && r.traineeId !== f.traineeId) return false;
      if (f.quizId && r.quizId !== f.quizId) return false;
      if (f.trainingDay && r.trainingDay !== f.trainingDay) return false;
      if (f.status && r.status !== f.status) return false;
      if (f.from && (!r.submittedAt || r.submittedAt.slice(0, 10) < f.from)) return false;
      if (f.to && (!r.submittedAt || r.submittedAt.slice(0, 10) > f.to)) return false;
      return true;
    });
  }

  function improvementSummary(rows) {
    const map = new Map();
    let n = 0;
    rows.forEach((r) => {
      if (!r.evaluationId) return;
      const ev = S.data.evaluations.find((e) => e.id === r.evaluationId);
      const eff = C.effectiveResult(ev);
      if (!eff) return;
      n++;
      eff.tickets.forEach((t) =>
        t.improvements.forEach((i) => {
          if (!map.has(i.category)) map.set(i.category, { count: 0, trainees: new Map() });
          const m = map.get(i.category);
          m.count++;
          m.trainees.set(r.traineeName, (m.trainees.get(r.traineeName) || 0) + 1);
        })
      );
    });
    return { n, items: [...map.entries()].sort((a, b) => b[1].count - a[1].count) };
  }

  function renderDashboard() {
    const f = S.dash;
    const rows = dashRows();
    const assigned = rows.filter((r) => !r.unassigned).length;
    const submitted = rows.filter((r) => r.submissionId).length;
    const pendingSub = rows.filter((r) => r.status === 'Awaiting Submission').length;
    const pendingEval = rows.filter((r) => ['Submitted', 'Pending Evaluation', 'Needs Trainer Review'].includes(r.status)).length;
    const scored = rows.filter((r) => r.percent != null);
    const aiOnly = scored.filter((r) => r.scoreSource === 'ai').length;
    const avg = scored.length ? C.round1(scored.reduce((a, r) => a + r.percent, 0) / scored.length) : null;
    const pf = scored.filter((r) => r.result === 'Pass' || r.result === 'Fail');
    const passRate = pf.length ? C.round1((pf.filter((r) => r.result === 'Pass').length / pf.length) * 100) : null;
    const days = [...new Set(S.data.quizzes.map((q) => q.trainingDay).filter(Boolean))];
    const metric = (label, value, sub) => `<div class="metric"><div class="label">${label}</div><div class="value">${value}</div>${sub ? `<div class="sub">${sub}</div>` : ''}</div>`;
    const imp = improvementSummary(rows);
    const quizzesInView = [...new Map(rows.map((r) => [r.quizId, r.quizTitle])).entries()];
    const traineesInView = [...new Map(rows.map((r) => [r.traineeId || r.traineeEmail, r.traineeName])).entries()];

    return `<div class="section"><div class="grid grid-4">
        <label class="field">Cohort<select data-filter="dash.cohort"><option value="">All</option>${cohorts().map((c) => `<option ${f.cohort === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>
        <label class="field">Trainee<select data-filter="dash.traineeId"><option value="">All</option>${S.data.trainees.map((t) => `<option value="${t.id}" ${f.traineeId === t.id ? 'selected' : ''}>${esc(t.name)}</option>`).join('')}</select></label>
        <label class="field">Quiz title<select data-filter="dash.quizId"><option value="">All</option>${S.data.quizzes.map((q) => `<option value="${q.id}" ${f.quizId === q.id ? 'selected' : ''}>${esc(q.title)}</option>`).join('')}</select></label>
        <label class="field">Training day<select data-filter="dash.trainingDay"><option value="">All</option>${days.map((d) => `<option ${f.trainingDay === d ? 'selected' : ''}>${esc(d)}</option>`).join('')}</select></label>
        <label class="field">Submitted from<input type="date" data-filter="dash.from" value="${esc(f.from)}"></label>
        <label class="field">Submitted to<input type="date" data-filter="dash.to" value="${esc(f.to)}"></label>
        <label class="field">Evaluation status<select data-filter="dash.status"><option value="">All</option>${C.EVAL_STATUSES.map((s) => `<option ${f.status === s ? 'selected' : ''}>${s}</option>`).join('')}</select></label>
        <div style="display:flex;align-items:flex-end"><button class="btn btn-ghost" data-action="clear-dash">Clear filters</button></div>
      </div></div>
      <div class="metrics">
        ${metric('Assigned assessments', assigned)}
        ${metric('Submissions', submitted)}
        ${metric('Pending submissions', pendingSub)}
        ${metric('Pending evaluations', pendingEval, 'Not yet approved as final')}
        ${metric('Average score', avg == null ? '—' : avg + '%', scored.length ? `${scored.length} scored${aiOnly ? `, ${aiOnly} AI-only (not yet approved)` : ''}` : 'No scored submissions')}
        ${metric('Passing rate', passRate == null ? '—' : passRate + '%', pf.length ? `${pf.length} with a pass/fail result` : 'No passing threshold configured')}
      </div>
      <div class="grid grid-2" style="align-items:start">
        <div class="section"><div class="section-head"><h3>Trainee performance</h3></div><div class="table-wrap"><table><thead><tr><th>Trainee</th><th>Quiz</th><th class="num">Score</th><th>Status</th></tr></thead><tbody>
          ${rows.map((r) => `<tr><td><button class="linkish" data-action="trainee-history" data-id="${esc(r.traineeId)}" data-name="${esc(r.traineeName)}">${esc(r.traineeName)}</button><div class="small muted">${esc(r.cohort)}</div></td><td>${esc(r.quizTitle)}</td><td class="num">${pctTxt(r.percent)}${r.scoreSource === 'ai' ? ' <span class="small muted" title="AI score, not yet approved">AI</span>' : ''}</td><td>${statusBadge(r.status)}</td></tr>`).join('') || '<tr><td colspan="4" class="empty">No assessments match these filters.</td></tr>'}
          </tbody></table></div></div>
        <div class="section"><div class="section-head"><div><h3>Areas for improvement</h3><p class="small muted">${imp.n ? `Based on ${imp.n} evaluated submission${imp.n === 1 ? '' : 's'} (trainer final where approved, otherwise AI).` : ''}</p></div></div>
          ${imp.n === 0 ? '<div class="empty small">No evaluated submissions in this view, so there is nothing to summarize yet.</div>' : imp.items.length ? `<table><thead><tr><th>Issue</th><th class="num">Occurrences</th><th>Trainees who need coaching</th></tr></thead><tbody>${imp.items.map(([cat, m]) => `<tr><td>${esc(cat)}</td><td class="num">${m.count}</td><td class="small">${[...m.trainees.entries()].sort((a, b) => b[1] - a[1]).map(([n, c]) => `${esc(n)} (${c})`).join(', ')}</td></tr>`).join('')}</tbody></table>` : '<div class="empty small">No areas for improvement recorded.</div>'}
        </div>
      </div>
      <div class="section"><div class="section-head"><div><h3>Compare across quizzes</h3><p class="small muted">Score per trainee and quiz. Blank means not assigned or not scored.</p></div></div>
        ${quizzesInView.length ? `<div class="table-wrap"><table class="matrix"><thead><tr><th>Trainee</th>${quizzesInView.map(([, t]) => `<th>${esc(t)}</th>`).join('')}</tr></thead><tbody>
        ${traineesInView.map(([tid, name]) => `<tr><td>${esc(name)}</td>${quizzesInView.map(([qid]) => { const r = rows.find((x) => (x.traineeId || x.traineeEmail) === tid && x.quizId === qid); return `<td class="cell">${r ? (r.percent != null ? r.percent + '%' : `<span class="muted small">${esc(r.status)}</span>`) : ''}</td>`; }).join('')}</tr>`).join('')}
        </tbody></table></div>` : '<div class="empty small">No data.</div>'}</div>
      ${renderExports()}`;
  }

  function renderExports() {
    const f = S.dash;
    const c = S.data.config;
    const q = (kind, fmt, extra) => `/api/export?kind=${kind}&format=${fmt}${f.cohort ? '&cohort=' + encodeURIComponent(f.cohort) : ''}${f.quizId ? '&quizId=' + encodeURIComponent(f.quizId) : ''}${extra || ''}`;
    const line = (label, kind) => `<tr><td>${label}</td><td><button class="btn btn-sm" data-action="download" data-href="${esc(q(kind, 'xlsx'))}">Excel</button> <button class="btn btn-sm" data-action="download" data-href="${esc(q(kind, 'csv'))}">CSV</button></td></tr>`;
    return `<div class="section"><div class="section-head"><div><h3>Export</h3><p class="small muted">Exports use the cohort and quiz filters above. Raw responses are exported as received.</p></div>
      <div class="btn-row"><button class="btn btn-primary" data-action="download" data-href="${esc(q('all', 'xlsx'))}">Full workbook (Excel, 6 tabs)</button>
      <button class="btn" data-action="sync-sheets" ${c.googleConfigured && c.googleWritesApproved ? '' : 'disabled'} title="${c.googleConfigured ? (c.googleWritesApproved ? '' : 'External Google writes are not approved on the server.') : 'Google is not connected.'}">Sync to Google Sheets</button></div></div>
      <table><tbody>${line('Raw trainee responses', 'raw')}${line('Evaluation results', 'evaluations')}${line('Score breakdown', 'breakdown')}${line('Areas for improvement', 'areas')}${line('Cohort performance summary', 'cohort')}${line('Quiz registry', 'registry')}
      <tr><td>Individual trainee report <select id="exp-trainee">${S.data.trainees.map((t) => `<option value="${t.id}" ${f.traineeId === t.id ? 'selected' : ''}>${esc(t.name)}</option>`).join('')}</select></td>
        <td><button class="btn btn-sm" data-action="download-trainee" data-fmt="xlsx">Excel</button> <button class="btn btn-sm" data-action="download-trainee" data-fmt="csv">CSV</button></td></tr></tbody></table></div>`;
  }

  function traineeHistory(id, name) {
    const rows = S.data.rows.filter((r) => (id ? r.traineeId === id : r.traineeName === name)).sort((a, b) => String(a.submittedAt || '').localeCompare(String(b.submittedAt || '')));
    openModal(
      `Assessment history: ${name}`,
      `<table><thead><tr><th>Quiz</th><th>Training day</th><th>Submitted</th><th class="num">Score</th><th>Result</th><th>Status</th><th></th></tr></thead><tbody>
      ${rows.map((r) => `<tr><td>${esc(r.quizTitle)}</td><td>${esc(r.trainingDay)}</td><td>${fmtDateTime(r.submittedAt)}</td><td class="num">${r.percent != null ? `${r.earned} / ${r.possible} (${r.percent}%)${r.scoreSource === 'ai' ? ' AI' : ''}` : '—'}</td><td>${r.result ? resultBadge(r.result) : ''}</td><td>${statusBadge(r.status)}</td><td>${r.submissionId ? `<button class="btn btn-sm" data-action="go-eval" data-id="${r.submissionId}">Open</button>` : ''}</td></tr>`).join('')}
      </tbody></table>`,
      `<button class="btn" data-action="close-modal">Close</button>`
    );
  }

  // =====================================================================
  // SETUP
  // =====================================================================
  function openSetup() {
    const c = S.data.config;
    openModal(
      'Setup and integrations',
      `<h3>Integrations</h3>
      <table><tbody>
        <tr><td>AI evaluator</td><td>${c.claudeConfigured ? `<span class="badge badge-teal">Claude connected</span> <span class="small muted">${esc(c.claudeModel)}</span>` : '<span class="badge badge-amber">Not connected</span> <span class="small">Set <span class="mono">ANTHROPIC_API_KEY</span> in the server environment. Until then the offline heuristic (not AI) scores staging tests.</span>'}</td></tr>
        <tr><td>Google Forms and Sheets</td><td>${c.googleConfigured ? '<span class="badge badge-teal">Apps Script connector configured</span>' : '<span class="badge">Not connected</span> <span class="small">Deploy <span class="mono">apps-script/Code.gs</span> and set <span class="mono">APPS_SCRIPT_URL</span> and <span class="mono">APPS_SCRIPT_TOKEN</span>. See README.</span>'}</td></tr>
        <tr><td>External Google actions</td><td>${c.googleWritesApproved ? '<span class="badge badge-red">Approved: forms and sheets can be created</span>' : '<span class="badge badge-green">Blocked (staging)</span> <span class="small">Nothing is created in Google until a trainer sets <span class="mono">GOOGLE_EXTERNAL_ACTIONS_APPROVED=yes</span>.</span>'}</td></tr>
        <tr><td>Trainer access</td><td>${c.passcodeRequired ? '<span class="badge badge-teal">Passcode required</span>' : '<span class="badge badge-amber">Open (local only)</span> <span class="small">Set <span class="mono">TRAINER_PASSCODE</span> before running anywhere but your own machine.</span>'}</td></tr>
      </tbody></table>
      ${c.googleConfigured ? '<button class="btn btn-sm" data-action="ping-google" style="margin-top:8px">Test Google connection</button>' : ''}
      <h3 style="margin-top:20px">Staging data</h3>
      <p class="small">Load fictional trainees and a two-ticket sample quiz. No real trainee or customer data is used.</p>
      <div class="btn-row"><button class="btn btn-primary" data-action="seed">Load fictional demo data</button><button class="btn btn-danger" data-action="reset">Reset all staging data…</button></div>`,
      `<button class="btn" data-action="close-modal">Close</button>`
    );
  }

  // =====================================================================
  // EVENTS
  // =====================================================================
  function confirmLeave() {
    return !S.dirty || confirm('You have unsaved changes to this quiz. Discard them?');
  }

  document.addEventListener('click', async (e) => {
    const tabBtn = e.target.closest('[data-tab]');
    if (tabBtn) {
      S.tab = tabBtn.dataset.tab;
      localSet('tab', S.tab);
      render();
      return;
    }
    const el = e.target.closest('[data-action]');
    if (!el) return;
    const a = el.dataset.action;
    if (a === 'modal-bg' && e.target !== el) return;
    const d = S.draft;
    const ti = el.dataset.t != null ? Number(el.dataset.t) : null;
    try {
      switch (a) {
        case 'modal-bg':
        case 'close-modal':
          closeModal();
          break;
        case 'open-setup':
          openSetup();
          break;
        case 'new-quiz':
          if (confirmLeave()) {
            startDraft(null);
            render();
          }
          break;
        case 'pick-quiz':
          if (confirmLeave()) {
            startDraft(quiz(el.dataset.id));
            render();
          }
          break;
        case 'save-draft':
          await saveDraft();
          break;
        case 'delete-quiz':
          if (confirm('Delete this quiz? This only works if it has no submissions.')) {
            await api('DELETE', `/api/quizzes/${d.id}`);
            S.draft = null;
            S.dirty = false;
            await refresh();
            render();
            toast('Quiz deleted.');
          }
          break;
        case 'preview':
          S.validation = C.validateQuiz(d);
          renderPreview();
          break;
        case 'generate': {
          if (!$('#approve-check') || !$('#approve-check').checked) return toast('Tick the approval box first.', true);
          const target = el.dataset.target;
          if (target === 'google' && !confirm('This creates a real Google Form (unpublished, not shared). Continue?')) return;
          el.disabled = true;
          await saveDraft(true);
          await api('POST', `/api/quizzes/${S.draft.id}/generate-form`, { approved: true, target, approvedBy: localGet('reviewer') || 'trainer' });
          closeModal();
          await refresh();
          startDraft(quiz(S.draft.id));
          render();
          toast(target === 'google' ? 'Google Form created. It is not shared with anyone.' : 'Staging form generated. Nothing was sent.');
          break;
        }
        case 'set-status':
          break;
        case 'copy-link':
          await navigator.clipboard.writeText(el.dataset.url).then(() => toast('Link copied.'), () => prompt('Copy this link', el.dataset.url));
          break;
        case 'send-dialog': {
          const names = d.assignedTrainees.map((id) => (trainee(id) || {}).name).filter(Boolean);
          const msg = `Hi team, please complete "${d.title}" by ${d.deadline || '(deadline)'}: ${absUrl(d.form.formUrl)}`;
          openModal('Send to trainees', `<div class="alert alert-warn">Sending is turned off in staging. Nothing is emailed or posted from this app. Copy the message and send it yourself once the prototype is approved.</div>
            <p class="small">Assigned trainees (${names.length}): ${names.map(esc).join(', ') || 'none'}</p><textarea rows="4" id="send-msg">${esc(msg)}</textarea>`,
            `<button class="btn" data-action="close-modal">Close</button><button class="btn btn-primary" data-action="copy-msg">Copy message</button>`);
          break;
        }
        case 'copy-msg':
          await navigator.clipboard.writeText($('#send-msg').value).then(() => toast('Message copied.'), () => toast('Copy failed. Select the text and copy it.', true));
          break;
        case 'toggle-trainee': {
          const id = el.dataset.id;
          d.assignedTrainees = el.checked ? [...new Set([...d.assignedTrainees, id])] : d.assignedTrainees.filter((x) => x !== id);
          markDirty(false);
          refreshAssign();
          break;
        }
        case 'assign-cohort':
          d.assignedTrainees = [...new Set([...d.assignedTrainees, ...S.data.trainees.filter((t) => t.cohort === d.cohort).map((t) => t.id)])];
          markDirty(false);
          refreshAssign();
          break;
        case 'clear-assign':
          d.assignedTrainees = [];
          markDirty(false);
          refreshAssign();
          break;
        case 'add-ticket':
          d.tickets.push(C.newTicket());
          markDirty(true);
          break;
        case 'del-ticket':
          if (d.tickets.length === 1) return toast('A quiz needs at least one ticket.', true);
          if (confirm(`Remove Ticket ${ti + 1} and its rubric?`)) {
            d.tickets.splice(ti, 1);
            markDirty(true);
          }
          break;
        case 'dup-ticket': {
          const copy = clone(d.tickets[ti]);
          copy.id = C.uid('tkt');
          copy.url = '';
          copy.rubric.criteria.forEach((c) => (c.id = C.uid('crit')));
          d.tickets.splice(ti + 1, 0, copy);
          markDirty(true);
          toast('Duplicated. Enter the new ticket URL.');
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
          markDirty(true, ti);
          break;
        case 'suggest-criteria': {
          const r = d.tickets[ti].rubric;
          if (r.criteria.length && !confirm('Replace the current criteria with the suggested five?')) return;
          if (!(Number(r.totalPoints) > 0)) r.totalPoints = 20;
          r.criteria = C.suggestedCriteria(r.totalPoints);
          markDirty(true, ti);
          break;
        }
        case 'del-criterion':
          d.tickets[ti].rubric.criteria.splice(Number(el.dataset.c), 1);
          markDirty(true, ti);
          break;
        case 'add-deduction':
          d.tickets[ti].rubric.criteria[Number(el.dataset.c)].deductions.push({ description: '', points: 1 });
          markDirty(true, ti);
          break;
        case 'del-deduction':
          d.tickets[ti].rubric.criteria[Number(el.dataset.c)].deductions.splice(Number(el.dataset.d), 1);
          markDirty(true, ti);
          break;
        case 'add-attach-link': {
          const inp = $(`#att-url-${ti}`);
          const url = inp.value.trim();
          if (!/^https?:\/\//i.test(url)) return toast('Enter a link that starts with https://', true);
          d.tickets[ti].attachments.push({ name: url.split('/').pop() || url, url });
          markDirty(true, ti);
          break;
        }
        case 'del-attach':
          d.tickets[ti].attachments.splice(Number(el.dataset.a), 1);
          markDirty(true, ti);
          break;
        case 'add-trainee': {
          const name = $('#tr-name').value.trim();
          const email = $('#tr-email').value.trim();
          const cohort = $('#tr-cohort').value.trim();
          await api('POST', '/api/trainees', { name, email, cohort });
          await refresh();
          render();
          toast('Trainee added.');
          break;
        }
        case 'del-trainee':
          if (confirm('Remove this trainee from the roster?')) {
            await api('DELETE', `/api/trainees/${el.dataset.id}`);
            await refresh();
            if (d) d.assignedTrainees = d.assignedTrainees.filter((x) => x !== el.dataset.id);
            render();
          }
          break;
        case 'open-sub':
          e.stopPropagation();
          openSubmission(el.dataset.id);
          break;
        case 'go-eval':
          closeModal();
          S.selectedSubmissionId = el.dataset.id;
          S.evalFilter = 'all';
          S.tab = 'evaluation';
          render();
          break;
        case 'sync-google': {
          el.disabled = true;
          const out = await api('POST', `/api/quizzes/${el.dataset.id}/sync-responses`);
          await refresh();
          render();
          toast(`Fetched ${out.fetched} response(s): ${out.added} new, ${out.duplicatesSkipped} already stored.`);
          break;
        }
        case 'add-samples': {
          const out = await api('POST', `/api/quizzes/${$('#sample-quiz').value}/sample-submissions`);
          await refresh();
          render();
          toast(`Added ${out.added} fictional submission(s).`);
          break;
        }
        case 'pick-sub':
          S.selectedSubmissionId = el.dataset.id;
          S.reviewDraft = null;
          render();
          break;
        case 'evaluate':
          await runEvaluation(el.dataset.id);
          break;
        case 'reevaluate': {
          const reason = prompt('Reevaluation creates a new version and keeps the old one for the audit trail. Reason for reevaluating?');
          if (reason !== null) await runEvaluation(el.dataset.id, true, reason || 'Trainer requested reevaluation');
          break;
        }
        case 'eval-all': {
          const ids = S.data.rows.filter((r) => r.status === 'Submitted' || r.status === 'Pending Evaluation').map((r) => r.submissionId);
          S.busy.evalAll = true;
          render();
          let ok = 0;
          let fail = 0;
          for (const id of ids) {
            try {
              await api('POST', `/api/submissions/${id}/evaluate`, {});
              ok++;
            } catch (_) {
              fail++;
            }
          }
          delete S.busy.evalAll;
          await refresh();
          render();
          toast(`${ok} evaluated${fail ? `, ${fail} failed (see each submission)` : ''}.`, fail > 0);
          break;
        }
        case 'del-improvement':
          S.reviewDraft.tickets[ti].improvements.splice(Number(el.dataset.i), 1);
          render();
          break;
        case 'add-improvement': {
          const det = $(`#imp-det-${ti}`).value.trim();
          if (!det) return toast('Write the detail first.', true);
          S.reviewDraft.tickets[ti].improvements.push({ category: $(`#imp-cat-${ti}`).value, detail: det });
          render();
          break;
        }
        case 'reset-review':
          S.reviewDraft = null;
          render();
          break;
        case 'save-review':
          await submitReview(false);
          break;
        case 'approve-review':
          await submitReview(true);
          break;
        case 'trainee-history':
          traineeHistory(el.dataset.id, el.dataset.name);
          break;
        case 'clear-dash':
          S.dash = { cohort: '', traineeId: '', quizId: '', trainingDay: '', status: '', from: '', to: '' };
          render();
          break;
        case 'download':
          await download(el.dataset.href);
          break;
        case 'download-trainee':
          await download(`/api/export?kind=trainee&format=${el.dataset.fmt}&traineeId=${encodeURIComponent($('#exp-trainee').value)}`);
          break;
        case 'sync-sheets':
          if (confirm('Write the six database tabs to the connected Google Sheet?')) {
            await api('POST', '/api/google/sync-sheets');
            toast('Google Sheets updated.');
          }
          break;
        case 'ping-google': {
          const out = await api('POST', '/api/google/ping');
          toast(`Google connection OK${out && out.spreadsheetUrl ? '' : ''}.`);
          break;
        }
        case 'seed':
          await api('POST', '/api/demo/seed');
          closeModal();
          await refresh();
          S.draft = null;
          S.tab = 'builder';
          render();
          toast('Fictional demo data loaded.');
          break;
        case 'reset': {
          const phrase = prompt('This deletes all staging quizzes, submissions and evaluations. Type RESET STAGING DATA to confirm.');
          if (phrase === null) return;
          await api('POST', '/api/demo/reset', { confirm: phrase });
          closeModal();
          S.draft = null;
          S.dirty = false;
          await refresh();
          render();
          toast('Staging data reset.');
          break;
        }
      }
    } catch (err) {
      toast(err.message + (err.details ? ' ' + err.details.join(' ') : ''), true);
      if (el.disabled !== undefined) el.disabled = false;
    }
  });

  async function download(href) {
    const headers = {};
    const pass = sessionGet();
    if (pass) headers['x-trainer-passcode'] = pass;
    const res = await fetch(href, { headers });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.error || 'Export failed.');
    }
    const blob = await res.blob();
    const name = (/filename="([^"]+)"/.exec(res.headers.get('Content-Disposition') || '') || [])[1] || 'export';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => (URL.revokeObjectURL(a.href), a.remove()), 1000);
    toast(`Downloaded ${name}`);
  }

  function markDirty(rerender, ticketIdx) {
    S.dirty = true;
    S.validation = C.validateQuiz(S.draft);
    if (rerender) {
      const open = [...document.querySelectorAll('details[open]')].map((x) => x.querySelector('summary') && x.querySelector('summary').textContent.trim().slice(0, 40) + '|' + (x.closest('.ticket-card') || {}).id);
      const y = window.scrollY;
      render();
      document.querySelectorAll('details').forEach((x) => {
        const key = x.querySelector('summary') && x.querySelector('summary').textContent.trim().slice(0, 40) + '|' + (x.closest('.ticket-card') || {}).id;
        if (open.includes(key)) x.open = true;
      });
      if (ticketIdx != null) {
        const card = document.getElementById(`ticket-${ticketIdx}`);
        if (card) card.querySelectorAll('details').forEach((x, i) => i === 1 && (x.open = true));
      }
      window.scrollTo(0, y);
    }
  }

  document.addEventListener('input', (e) => {
    const el = e.target;
    if (el.dataset.bind && S.draft) {
      let v = el.value;
      if (el.dataset.type === 'number') v = v === '' ? '' : Number(v);
      setPath(S.draft, el.dataset.bind, v);
      S.dirty = true;
      S.validation = C.validateQuiz(S.draft);
      if (el.dataset.sum != null) {
        const i = Number(el.dataset.sum);
        const r = S.draft.tickets[i].rubric;
        const sum = C.rubricMax(r);
        const b = document.getElementById(`sum-${i}`);
        if (b) {
          b.textContent = `${sum} of ${Number(r.totalPoints) || 0} points allocated`;
          b.className = 'badge ' + (Number(r.totalPoints) > 0 && sum === Number(r.totalPoints) ? 'badge-green' : 'badge-amber');
        }
      }
      return;
    }
    if (el.dataset.review && S.reviewDraft) {
      const parts = el.dataset.review.split('.');
      const rd = S.reviewDraft;
      if (parts.length === 1) rd[parts[0]] = el.value;
      else if (parts[1] === 'notes') rd.tickets[Number(parts[0])].trainerNotes = el.value;
      else {
        const t = rd.tickets[Number(parts[0])];
        const c = t.criteria[Number(parts[1])];
        if (parts[2] === 'awarded') {
          c.awarded = el.value === '' ? 0 : Number(el.value);
          C.recomputeTicket(t);
          const s = document.getElementById(`tsum-${parts[0]}`);
          if (s) s.textContent = `${t.earned} / ${t.possible}`;
          const box = document.getElementById('draft-box');
          const ev = S.data.evaluations.find((x) => x.id === rd.evalId);
          if (box && ev) box.innerHTML = scoreBox('Your edits (unsaved preview)', C.summarize(rd.tickets, quiz(ev.quizId)), 'Save or approve to keep');
        } else c.reason = el.value;
      }
    }
  });

  document.addEventListener('change', async (e) => {
    const el = e.target;
    if (el.dataset.filter) {
      setPath(S, el.dataset.filter, el.value);
      if (el.dataset.filter === 'evalFilter') S.selectedSubmissionId = null;
      render();
      return;
    }
    if (el.dataset.bind && el.dataset.rerender) {
      refreshAssign();
      return;
    }
    if (el.dataset.action === 'set-status') {
      const status = el.value;
      if (!S.draft.id) {
        S.draft.status = status;
        return;
      }
      if (status === 'Sent' && !confirm('Mark as Sent? This only changes the status label. Nothing is sent to trainees from this app.')) return render();
      try {
        await api('POST', `/api/quizzes/${S.draft.id}/status`, { status });
        await refresh();
        S.draft.status = status;
        render();
        toast(`Status set to ${status}.`);
      } catch (err) {
        toast(err.message, true);
        render();
      }
      return;
    }
    if (el.dataset.action === 'upload') {
      const file = el.files[0];
      const ti = Number(el.dataset.t);
      if (!file) return;
      if (file.size > 8 * 1024 * 1024) return toast('Files must be 8 MB or smaller.', true);
      const dataUrl = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result);
        r.onerror = rej;
        r.readAsDataURL(file);
      });
      try {
        const out = await api('POST', '/api/uploads', { name: file.name, dataUrl });
        S.draft.tickets[ti].attachments.push(out);
        markDirty(true, ti);
        toast('Attachment uploaded.');
      } catch (err) {
        toast(err.message, true);
      }
    }
  });

  window.addEventListener('beforeunload', (e) => {
    if (S.dirty) {
      e.preventDefault();
      e.returnValue = '';
    }
  });

  // ---------- boot ----------
  (async function boot() {
    try {
      await refresh();
      render();
    } catch (err) {
      $('#app').innerHTML = `<div class="section alert alert-error">Could not load the app: ${esc(err.message)}</div>`;
    }
  })();
})();
