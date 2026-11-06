'use strict';
/*
 * Staging end-to-end test. Runs the 15 staging checks from the spec with fictional data.
 * External services are replaced with local mocks:
 *   - a mock Claude Messages API (checks the request shape and score handling)
 *   - a mock Google Apps Script web app (checks the connector contract and approval gate)
 * Nothing leaves this machine.
 */
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const assert = require('assert/strict');
const ExcelJS = require('exceljs');

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'tah-test-'));
process.env.DATA_FILE = path.join(TMP, 'db.json');
delete process.env.ANTHROPIC_API_KEY;
delete process.env.APPS_SCRIPT_URL;
delete process.env.APPS_SCRIPT_TOKEN;
delete process.env.GOOGLE_EXTERNAL_ACTIONS_APPROVED;
delete process.env.TRAINER_PASSCODE;

const { server } = require('../server/server');
const demo = require('../server/demo');
const core = require('../server/core');

let BASE;
let passed = 0;
const results = [];
async function step(name, fn) {
  try {
    await fn();
    passed++;
    results.push(`PASS  ${name}`);
    console.log(`PASS  ${name}`);
  } catch (err) {
    results.push(`FAIL  ${name}: ${err.message}`);
    console.log(`FAIL  ${name}\n      ${err.stack.split("\n").slice(0, 12).join('\n      ')}`);
    process.exitCode = 1;
  }
}

async function req(method, p, body) {
  const res = await fetch(BASE + p, { method, headers: { 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
  const ct = res.headers.get('content-type') || '';
  const data = ct.includes('json') ? await res.json() : Buffer.from(await res.arrayBuffer());
  return { status: res.status, data, headers: res.headers };
}

// ---------- mocks ----------
const claudeCalls = [];
let claudeMode = 'ok';
const mockClaude = http.createServer((rq, rs) => {
  let b = '';
  rq.on('data', (c) => (b += c));
  rq.on('end', () => {
    const body = JSON.parse(b);
    claudeCalls.push({ url: rq.url, headers: rq.headers, body });
    if (claudeMode === 'fail') {
      rs.writeHead(500, { 'Content-Type': 'application/json' });
      return rs.end(JSON.stringify({ type: 'error', error: { type: 'api_error', message: 'mock outage' } }));
    }
    const prompt = body.messages[0].content;
    const crit = JSON.parse(/<criteria>\n([\s\S]*?)\n<\/criteria>/.exec(prompt)[1]);
    const weak = /full refund|Rescheduled to Monday/.test(prompt);
    const out = {
      // Drop the last criterion on purpose to check the server flags it; over-award the first to check clamping.
      criteria: crit.slice(0, -1).map((c, i) => ({ criterionId: c.criterionId, awarded: i === 0 ? c.maxPoints + 50 : weak ? 0 : c.maxPoints, reason: weak ? `Missed: ${c.name}` : '', needsReview: false })),
      strengths: weak ? [] : ['Correct refund amount'],
      improvements: weak ? [{ category: 'Incorrect customer resolution', detail: 'Offered a full refund instead of 25%' }] : [],
      coaching: weak ? ['Walk through the partial refund policy'] : [],
      criticalErrors: weak && /full refund/.test(prompt) ? ['Promised a full refund'] : [],
      autoFail: false,
      autoFailReason: '',
      needsTrainerReview: weak,
      reviewNotes: 'mock',
    };
    rs.writeHead(200, { 'Content-Type': 'application/json', 'request-id': 'req_mock' });
    rs.end(JSON.stringify({ id: 'msg_mock', type: 'message', role: 'assistant', model: body.model, content: [{ type: 'text', text: JSON.stringify(out) }], stop_reason: 'end_turn', stop_sequence: null, usage: { input_tokens: 100, output_tokens: 100 } }));
  });
});

const googleCalls = [];
let googleQuiz = null;
const mockGoogle = http.createServer((rq, rs) => {
  let b = '';
  rq.on('data', (c) => (b += c));
  rq.on('end', () => {
    const body = JSON.parse(b);
    googleCalls.push(body);
    const send = (o) => (rs.writeHead(200, { 'Content-Type': 'application/json' }), rs.end(JSON.stringify(o)));
    if (body.token !== 'test-token') return send({ ok: false, error: 'Unauthorized' });
    if (body.action === 'createForm') {
      googleQuiz = body.payload.quiz;
      return send({ ok: true, data: { formId: 'form_mock', publishedUrl: 'https://docs.google.com/forms/d/e/mock/viewform', editUrl: 'https://docs.google.com/forms/d/mock/edit', responseSheetUrl: 'https://docs.google.com/spreadsheets/d/mock', itemMap: {}, acceptingResponses: false } });
    }
    if (body.action === 'fetchResponses') {
      const t = googleQuiz.tickets;
      return send({ ok: true, data: [
        { responseId: 'resp_1', submittedAt: '2026-11-06T15:00:00.000Z', traineeEmail: 'delta.trainee@example.com', traineeName: 'Test Trainee Delta', cohort: googleQuiz.cohort, answers: t.map((x) => ({ ticketId: x.id, q1: 'g1', q2: 'g2', q3: 'g3', q4: 'N/A' })) },
        { responseId: 'resp_2', submittedAt: '2026-11-06T15:05:00.000Z', traineeEmail: 'echo.trainee@example.com', traineeName: 'Test Trainee Echo', cohort: googleQuiz.cohort, answers: t.map((x) => ({ ticketId: x.id, q1: 'g1', q2: 'g2', q3: 'g3', q4: 'N/A' })) },
      ] });
    }
    if (body.action === 'syncSheets') return send({ ok: true, data: { rowsWritten: Object.fromEntries(Object.entries(body.payload.tables).map(([k, v]) => [k, v.length])) } });
    return send({ ok: true, data: {} });
  });
});

const listen = (srv) => new Promise((r) => srv.listen(0, '127.0.0.1', () => r(srv.address().port)));

(async function main() {
  const port = await listen(server);
  BASE = `http://127.0.0.1:${port}`;
  const claudePort = await listen(mockClaude);
  const googlePort = await listen(mockGoogle);

  // Seed fictional trainees only (no quiz) so the quiz is built through the API like a trainer would.
  const now = new Date().toISOString();
  for (const t of demo.TRAINEES) await req('POST', '/api/trainees', t);
  let state = (await req('GET', '/api/state')).data;
  const template = demo.sampleQuiz(state.trainees);
  delete template.id;

  let quiz;
  await step('1. Create a sample quiz with two tickets', async () => {
    const r = await req('POST', '/api/quizzes', template);
    assert.equal(r.status, 200);
    quiz = r.data.quiz;
    assert.equal(quiz.tickets.length, 2);
    assert.equal(quiz.status, 'Draft');
    assert.deepEqual(r.data.validation.errors, []);
  });

  await step('2. Two different ticket URLs, and duplicate URLs are flagged', async () => {
    assert.notEqual(quiz.tickets[0].url, quiz.tickets[1].url);
    const dup = JSON.parse(JSON.stringify(quiz));
    dup.tickets[1].url = dup.tickets[0].url;
    assert.ok(core.validateQuiz(dup).warnings.some((w) => /same URL/.test(w)));
  });

  await step('3. Separate scoring guidelines per ticket, with point totals enforced', async () => {
    const [a, b] = quiz.tickets;
    assert.equal(a.rubric.totalPoints, 20);
    assert.equal(b.rubric.totalPoints, 10);
    assert.notDeepEqual(a.rubric.criteria.map((c) => c.name), b.rubric.criteria.map((c) => c.name));
    const bad = JSON.parse(JSON.stringify(quiz));
    bad.tickets[1].rubric.totalPoints = 12;
    assert.ok(core.validateQuiz(bad).errors.some((e) => /add up to 10/.test(e)));
  });

  await step('4. Different policy contexts and sample responses per ticket', async () => {
    assert.notEqual(quiz.tickets[0].references.policyContext, quiz.tickets[1].references.policyContext);
    assert.notEqual(quiz.tickets[0].references.sampleC, quiz.tickets[1].references.sampleC);
  });

  await step('5. Preview trainee form: generated only after approval, and it hides rubrics', async () => {
    assert.equal((await req('GET', `/api/public/quiz/${quiz.id}`)).status, 404, 'form must not exist before generation');
    assert.equal((await req('POST', `/api/quizzes/${quiz.id}/generate-form`, {})).status, 400, 'approval required');
    const g = await req('POST', `/api/quizzes/${quiz.id}/generate-form`, { approved: true, target: 'staging', approvedBy: 'Test Trainer' });
    assert.equal(g.status, 200);
    assert.equal(g.data.status, 'Ready');
    assert.equal(g.data.form.mode, 'staging');
    const pub = await req('GET', `/api/public/quiz/${quiz.id}`);
    assert.equal(pub.status, 200);
    const text = JSON.stringify(pub.data);
    assert.equal(pub.data.quiz.tickets.length, 2);
    assert.equal(pub.data.quiz.tickets[0].url, quiz.tickets[0].url);
    assert.equal(pub.data.quiz.tickets[1].url, quiz.tickets[1].url);
    assert.deepEqual(pub.data.quiz.questions.map((q) => q.label), core.QUESTIONS.map((q) => q.label));
    for (const t of quiz.tickets) {
      for (const secret of [t.rubric.expectedIntent, t.rubric.criticalErrors, t.rubric.requiredCrmActions.split('\n')[0], t.references.policyContext, t.references.sampleC, t.rubric.autoFailConditions].filter(Boolean)) {
        assert.ok(!text.includes(secret), `trainee view leaked: ${secret.slice(0, 40)}`);
      }
    }
    assert.ok(!/rubric|references|criteria/.test(text), 'trainee view contains rubric fields');
    const html = await req('GET', `/form/${quiz.id}`);
    assert.equal(html.status, 200);
  });

  let gQuiz;
  await step('6. Google Form generation: blocked until connected and approved; payload has no rubric (mock Apps Script)', async () => {
    const r1 = await req('POST', '/api/quizzes', { ...template, title: 'STAGING Google path copy' });
    gQuiz = r1.data.quiz;
    const notConnected = await req('POST', `/api/quizzes/${gQuiz.id}/generate-form`, { approved: true, target: 'google' });
    assert.equal(notConnected.status, 503);
    process.env.APPS_SCRIPT_URL = `http://127.0.0.1:${googlePort}/exec`;
    process.env.APPS_SCRIPT_TOKEN = 'test-token';
    const notApproved = await req('POST', `/api/quizzes/${gQuiz.id}/generate-form`, { approved: true, target: 'google' });
    assert.equal(notApproved.status, 403);
    assert.equal(googleCalls.length, 0, 'no Google call may happen before approval');
    process.env.GOOGLE_EXTERNAL_ACTIONS_APPROVED = 'yes';
    const ok = await req('POST', `/api/quizzes/${gQuiz.id}/generate-form`, { approved: true, target: 'google' });
    assert.equal(ok.status, 200, JSON.stringify(ok.data));
    assert.equal(ok.data.form.mode, 'google');
    const sent = JSON.stringify(googleCalls[0].payload);
    assert.ok(!/rubric|references|criticalErrors|sampleC/.test(sent), 'Google payload leaked trainer-only data');
    assert.equal(googleCalls[0].payload.quiz.tickets[1].url, template.tickets[1].url);
    delete process.env.GOOGLE_EXTERNAL_ACTIONS_APPROVED;
  });

  const answers = demo.sampleAnswers(quiz);
  let strongId, weakId;
  await step('7. Submit sample trainee answers through the staging form endpoint', async () => {
    const s = await req('POST', `/api/public/quiz/${quiz.id}/submit`, { traineeName: 'Test Trainee Alpha', traineeEmail: 'alpha.trainee@example.com', cohort: quiz.cohort, answers: answers.strong });
    const w = await req('POST', `/api/public/quiz/${quiz.id}/submit`, { traineeName: 'Test Trainee Bravo', traineeEmail: 'bravo.trainee@example.com', cohort: quiz.cohort, answers: answers.weak });
    assert.equal(s.status, 200);
    assert.equal(w.status, 200);
    strongId = s.data.submissionId;
    weakId = w.data.submissionId;
    assert.equal((await req('POST', `/api/public/quiz/${quiz.id}/submit`, { traineeName: '', traineeEmail: '' })).status, 400);
  });

  await step('8. Retrieve responses (staging table, and Google sync with de-duplication)', async () => {
    state = (await req('GET', '/api/state')).data;
    const rows = state.rows.filter((r) => r.quizId === quiz.id);
    assert.equal(rows.find((r) => r.traineeName === 'Test Trainee Alpha').status, 'Submitted');
    assert.equal(rows.find((r) => r.traineeName === 'Test Trainee Bravo').ticketsAnswered, 2);
    assert.equal(rows.find((r) => r.traineeName === 'Test Trainee Charlie').status, 'Awaiting Submission');
    const sync1 = await req('POST', `/api/quizzes/${gQuiz.id}/sync-responses`);
    assert.equal(sync1.data.added, 2);
    const sync2 = await req('POST', `/api/quizzes/${gQuiz.id}/sync-responses`);
    assert.equal(sync2.data.added, 0);
    assert.equal(sync2.data.duplicatesSkipped, 2);
  });

  let evStrong, evWeak;
  await step('9. Run scoring (offline heuristic, then Claude through a mock API)', async () => {
    const a = await req('POST', `/api/submissions/${strongId}/evaluate`, {});
    const b = await req('POST', `/api/submissions/${weakId}/evaluate`, {});
    assert.equal(a.status, 200, JSON.stringify(a.data));
    assert.equal(b.status, 200);
    assert.equal(a.data.ai.evaluator, 'heuristic');
    evStrong = a.data;
    evWeak = b.data;
    assert.ok(evStrong.ai.overall.percent > evWeak.ai.overall.percent, 'strong answer should outscore weak answer');
    assert.equal((await req('POST', `/api/submissions/${weakId}/evaluate`, {})).status, 409, 'duplicate evaluation must be refused');

    // Claude path against the local mock.
    process.env.ANTHROPIC_API_KEY = 'sk-ant-test-dummy';
    process.env.ANTHROPIC_BASE_URL = `http://127.0.0.1:${claudePort}`;
    claudeMode = 'fail';
    const failed = await req('POST', `/api/submissions/${weakId}/evaluate`, { reevaluate: true, reason: 'test outage' });
    assert.equal(failed.status, 502);
    state = (await req('GET', '/api/state')).data;
    const subAfterFail = state.submissions.find((s) => s.id === weakId);
    assert.equal(subAfterFail.evaluationState, 'failed');
    assert.equal(subAfterFail.answers[0].q2, 'Issue full refund', 'raw answers must survive a failed evaluation');
    assert.ok(state.evaluations.find((e) => e.id === evWeak.id && !e.supersededAt), 'previous evaluation stays current after a failure');
    claudeMode = 'ok';
    claudeCalls.length = 0;
    const c = await req('POST', `/api/submissions/${weakId}/evaluate`, { reevaluate: true, reason: 'Switch to Claude' });
    assert.equal(c.status, 200, JSON.stringify(c.data));
    evWeak = c.data;
    assert.equal(evWeak.ai.evaluator, 'claude');
    assert.equal(evWeak.version, 2);
    assert.equal(claudeCalls.length, 2, 'one Claude call per ticket');
    const call = claudeCalls[0];
    assert.equal(call.body.model, 'claude-opus-5-5');
    assert.equal(call.body.output_config.format.type, 'json_schema');
    assert.ok(call.body.system.includes('Do not invent policy requirements'));
    assert.ok(call.body.messages[0].content.includes('FICTIONAL STAGING POLICY'), 'references must reach the evaluator');
    assert.ok(call.headers['x-api-key'] === 'sk-ant-test-dummy' || call.headers.authorization === 'Bearer sk-ant-test-dummy', 'server-side key sent by the SDK');
    state = (await req('GET', '/api/state')).data;
    assert.ok(state.evaluations.find((e) => e.id === evStrong.id), 'old version kept');
    assert.equal(state.evaluations.filter((e) => e.submissionId === weakId && !e.supersededAt).length, 1);
  });

  await step('10. Deductions match the rubric (clamped, totals recomputed, reasons given, gaps flagged)', async () => {
    for (const ev of [evStrong, evWeak]) {
      ev.ai.tickets.forEach((t, i) => {
        const rubric = quiz.tickets[i].rubric;
        assert.equal(t.possible, rubric.totalPoints);
        assert.equal(t.criteria.length, rubric.criteria.length);
        let sum = 0;
        t.criteria.forEach((c, ci) => {
          assert.equal(c.maxPoints, rubric.criteria[ci].maxPoints);
          assert.ok(c.awarded >= 0 && c.awarded <= c.maxPoints, 'awarded within 0..max');
          if (c.awarded < c.maxPoints) assert.ok(c.reason.trim() || c.needsReview, `deduction without reason: ${c.name}`);
          sum += c.awarded;
        });
        assert.equal(t.earned, core.round1(sum));
      });
    }
    const t1 = evWeak.ai.tickets[0];
    assert.equal(t1.criteria[0].awarded, t1.criteria[0].maxPoints, 'over-award clamped to max');
    const last = t1.criteria[t1.criteria.length - 1];
    assert.equal(last.needsReview, true, 'criterion missing from model output is flagged');
    assert.equal(last.awarded, 0);
    assert.ok(t1.needsTrainerReview);
    assert.deepEqual(t1.criticalErrors, ['Promised a full refund']);
  });

  await step('11. Areas for improvement are relevant to what the trainee missed', async () => {
    const cats = evWeak.ai.tickets.flatMap((t) => t.improvements.map((i) => i.category));
    assert.ok(cats.includes('Incorrect customer resolution'));
    assert.ok(cats.includes('Critical error'));
    const heurCats = evStrong.ai.tickets.flatMap((t) => t.improvements.map((i) => i.category));
    heurCats.forEach((c) => assert.ok(core.IMPROVEMENT_CATEGORIES.includes(c)));
    assert.ok(evWeak.ai.overall.coachingFocus.length > 0);
  });

  let reviewed;
  await step('12. Trainer edits and approves an evaluation', async () => {
    const tickets = evWeak.ai.tickets.map((t) => ({ criteria: t.criteria.map((c) => ({ awarded: c.awarded, reason: c.reason })), improvements: t.improvements }));
    tickets[1].criteria[1].awarded = 2;
    tickets[1].criteria[1].reason = 'Trainer: partial credit for checking availability';
    const ed = await req('POST', `/api/evaluations/${evWeak.id}/review`, { tickets, comments: 'Discussed in 1:1', reviewer: 'Test Trainer' });
    assert.equal(ed.status, 200);
    assert.equal(ed.data.trainer.approvedAt, undefined);
    const ap = await req('POST', `/api/evaluations/${evWeak.id}/review`, { approve: true, overrideResult: 'Needs retake', reviewer: 'Test Trainer' });
    assert.equal(ap.status, 200);
    reviewed = ap.data;
    assert.ok(reviewed.trainer.approvedAt);
    assert.equal(reviewed.trainer.approvedBy, 'Test Trainer');
    assert.equal(reviewed.history.length, 2);
    state = (await req('GET', '/api/state')).data;
    const row = state.rows.find((r) => r.submissionId === weakId);
    assert.equal(row.status, 'Evaluated');
    assert.equal(row.scoreSource, 'trainer');
    assert.equal(row.result, 'Needs retake');
  });

  await step('13. Trainer final score is stored separately from the original AI score', async () => {
    assert.equal(reviewed.ai.overall.earned, evWeak.ai.overall.earned, 'AI score unchanged');
    assert.equal(reviewed.trainer.overall.earned, core.round1(evWeak.ai.overall.earned + 2));
    assert.equal(reviewed.ai.tickets[1].criteria[1].awarded, evWeak.ai.tickets[1].criteria[1].awarded);
    assert.equal(reviewed.trainer.tickets[1].criteria[1].reason, 'Trainer: partial credit for checking availability');
    const sub = state.submissions.find((s) => s.id === weakId);
    const crypto = require('crypto');
    assert.equal(crypto.createHash('sha256').update(JSON.stringify(sub.answers)).digest('hex'), sub.rawHash, 'raw submission unchanged');
  });

  const outDir = path.join(TMP, 'exports');
  fs.mkdirSync(outDir);
  let rawCsv, evalCsv, workbookPath;
  await step('14. Export raw responses and evaluation results (CSV and Excel)', async () => {
    const r1 = await req('GET', `/api/export?kind=raw&format=csv&quizId=${quiz.id}`);
    const r2 = await req('GET', `/api/export?kind=evaluations&format=csv`);
    const r3 = await req('GET', `/api/export?kind=all&format=xlsx`);
    const r4 = await req('GET', `/api/export?kind=trainee&format=xlsx&traineeId=${state.trainees.find((t) => t.name === 'Test Trainee Bravo').id}`);
    const r5 = await req('GET', `/api/export?kind=all&format=csv`);
    [r1, r2, r3, r4].forEach((r) => assert.equal(r.status, 200));
    assert.equal(r5.status, 400, 'multi-table CSV is refused with a clear message');
    rawCsv = r1.data.toString('utf8');
    evalCsv = r2.data.toString('utf8');
    workbookPath = path.join(outDir, 'all.xlsx');
    fs.writeFileSync(path.join(outDir, 'raw.csv'), rawCsv);
    fs.writeFileSync(path.join(outDir, 'evaluations.csv'), evalCsv);
    fs.writeFileSync(workbookPath, r3.data);
    fs.writeFileSync(path.join(outDir, 'trainee.xlsx'), r4.data);
    assert.match(r3.headers.get('content-disposition'), /full-workbook-\d{4}-\d{2}-\d{2}\.xlsx/);
  });

  await step('15. Exported files contain the correct data', async () => {
    assert.ok(rawCsv.includes('Test Trainee Alpha') && rawCsv.includes('Issue full refund'));
    assert.ok(rawCsv.includes(quiz.tickets[0].url) && rawCsv.includes(quiz.tickets[1].url));
    core.QUESTIONS.forEach((q) => assert.ok(rawCsv.includes(q.label.includes(',') ? `"${q.label}"` : q.label)));
    assert.equal(rawCsv.trim().split('\r\n').length, 1 + 2 * 2, 'header + 2 submissions x 2 tickets');
    assert.ok(evalCsv.includes('Needs retake') && evalCsv.includes('No (superseded)'));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(workbookPath);
    assert.deepEqual(wb.worksheets.map((w) => w.name), ['Quiz Registry', 'Raw Responses', 'Evaluation Results', 'Score Breakdown', 'Areas for Improvement', 'Cohort Performance']);
    const er = wb.getWorksheet('Evaluation Results');
    const hdr = er.getRow(1).values;
    const iAi = hdr.indexOf('AI points earned');
    const iTr = hdr.indexOf('Trainer points earned');
    const iCur = hdr.indexOf('Current');
    const iName = hdr.indexOf('Trainee name');
    let found = false;
    er.eachRow((row, n) => {
      if (n > 1 && row.values[iName] === 'Test Trainee Bravo' && row.values[iCur] === 'Yes') {
        found = true;
        assert.equal(row.values[iAi], reviewed.ai.overall.earned);
        assert.equal(row.values[iTr], reviewed.trainer.overall.earned);
      }
    });
    assert.ok(found, 'current Bravo evaluation row present');
    const sb = wb.getWorksheet('Score Breakdown');
    assert.equal(sb.rowCount - 1, 3 * (quiz.tickets[0].rubric.criteria.length + quiz.tickets[1].rubric.criteria.length), 'one breakdown row per criterion per evaluation version');
    const tw = new ExcelJS.Workbook();
    await tw.xlsx.readFile(path.join(outDir, 'trainee.xlsx'));
    assert.equal(tw.worksheets[0].name, 'Summary');
    let alphaLeak = false;
    tw.worksheets.forEach((w) => w.eachRow((r) => JSON.stringify(r.values).includes('Alpha') && (alphaLeak = true)));
    assert.ok(!alphaLeak, 'individual report contains only that trainee');
  });

  await step('Extra: Sheets sync is blocked until approved; CSV escapes formula injection', async () => {
    assert.equal((await req('POST', '/api/google/sync-sheets')).status, 403);
    process.env.GOOGLE_EXTERNAL_ACTIONS_APPROVED = 'yes';
    const ok = await req('POST', '/api/google/sync-sheets');
    assert.equal(ok.status, 200);
    assert.equal(Object.keys(ok.data.rowsWritten).length, 6);
    delete process.env.GOOGLE_EXTERNAL_ACTIONS_APPROVED;
    assert.ok(core.toCSV([{ a: '=HYPERLINK("x")' }]).includes(`"'=HYPERLINK(""x"")"`));
  });

  console.log(`\n${passed}/${results.length} checks passed. Exports written to ${outDir}`);
  server.close();
  mockClaude.close();
  mockGoogle.close();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
