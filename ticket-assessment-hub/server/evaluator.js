'use strict';
// Evaluates a submission ticket by ticket.
// With ANTHROPIC_API_KEY set: Claude grades against the trainer's rubric (key stays server-side).
// Without it: a clearly labelled offline heuristic, so the staging workflow can still be tested.
const core = require('./core');

let Anthropic = null;
try {
  Anthropic = require('@anthropic-ai/sdk');
  Anthropic = Anthropic.default || Anthropic;
} catch (_) {
  Anthropic = null;
}

function claudeConfigured() {
  return !!(process.env.ANTHROPIC_API_KEY && Anthropic);
}

function evaluatorName() {
  if (process.env.EVALUATOR === 'heuristic') return 'heuristic';
  return claudeConfigured() ? 'claude' : 'heuristic';
}

const MODEL = () => process.env.CLAUDE_MODEL || 'claude-opus-5-5';

let client = null;
function getClient() {
  if (!client) client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY, maxRetries: 3, timeout: 180000 });
  return client;
}

async function claudeEvaluateTicket(ticket, ticketNumber, answers) {
  const params = {
    model: MODEL(),
    max_tokens: 16000,
    system: core.EVALUATOR_SYSTEM_PROMPT,
    output_config: {
      effort: process.env.CLAUDE_EFFORT || 'high',
      format: { type: 'json_schema', schema: core.TICKET_EVAL_SCHEMA },
    },
    messages: [{ role: 'user', content: core.buildTicketPrompt(ticket, ticketNumber, answers) }],
  };
  let res;
  if (process.env.CLAUDE_FALLBACKS !== 'off') {
    // Server-side fallback: if a safety classifier declines, the API reruns on its recommended model.
    try {
      res = await getClient().beta.messages.create({ ...params, betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' });
    } catch (err) {
      if (!(err && err.status === 400)) throw err;
      res = await getClient().messages.create(params);
    }
  } else {
    res = await getClient().messages.create(params);
  }
  if (res.stop_reason === 'refusal') throw new Error(`Claude declined to grade Ticket ${ticketNumber}. A trainer needs to score it by hand.`);
  if (res.stop_reason === 'max_tokens') throw new Error(`Claude's answer for Ticket ${ticketNumber} was cut off. Try reevaluating.`);
  const text = res.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (_) {
    throw new Error(`Claude returned an unreadable result for Ticket ${ticketNumber}. Try reevaluating.`);
  }
  return { raw: parsed, model: res.model, usage: res.usage };
}

async function evaluateSubmission(quiz, submission) {
  const mode = evaluatorName();
  const byTicket = new Map(submission.answers.map((a) => [a.ticketId, a]));
  const usage = [];
  let model = mode === 'claude' ? MODEL() : 'heuristic-v1 (not AI)';
  const tickets = await Promise.all(
    quiz.tickets.map(async (t, i) => {
      const ans = byTicket.get(t.id) || { q1: '', q2: '', q3: '', q4: '' };
      if (mode === 'claude') {
        const out = await claudeEvaluateTicket(t, i + 1, ans);
        model = out.model || model;
        usage.push(out.usage);
        return core.normalizeTicketEval(t, out.raw, i + 1);
      }
      return core.normalizeTicketEval(t, core.heuristicEvaluateTicket(t, ans), i + 1);
    })
  );
  return {
    evaluator: mode,
    model,
    usage,
    tickets,
    overall: core.summarize(tickets, quiz),
    rubricSnapshot: quiz.tickets.map((t) => ({ ticketId: t.id, rubric: t.rubric, references: t.references })),
  };
}

module.exports = { evaluateSubmission, evaluatorName, claudeConfigured, MODEL };
