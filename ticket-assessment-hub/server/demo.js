'use strict';
// Fictional staging data. No real trainees, customers, cleaners or tickets.
const core = require('./core');

const TRAINEES = [
  { name: 'Test Trainee Alpha', email: 'alpha.trainee@example.com', cohort: 'Staging Cohort A' },
  { name: 'Test Trainee Bravo', email: 'bravo.trainee@example.com', cohort: 'Staging Cohort A' },
  { name: 'Test Trainee Charlie', email: 'charlie.trainee@example.com', cohort: 'Staging Cohort A' },
  { name: 'Test Trainee Delta', email: 'delta.trainee@example.com', cohort: 'Staging Cohort B' },
  { name: 'Test Trainee Echo', email: 'echo.trainee@example.com', cohort: 'Staging Cohort B' },
];

function crit(name, category, maxPoints, guidance, deductions) {
  return { id: core.uid('crit'), name, category, maxPoints, guidance: guidance || '', deductions: deductions || [] };
}

function sampleQuiz(trainees) {
  const cohortA = trainees.filter((t) => t.cohort === 'Staging Cohort A').map((t) => t.id);
  const t1 = core.newTicket();
  Object.assign(t1, {
    url: 'https://example.com/fictional-tickets/T-1001',
    description: 'FICTIONAL. C reports the CP arrived 90 minutes late and left the kitchen uncleaned. C asks for a partial refund.',
    context: 'Read the full ticket thread before answering.',
  });
  t1.rubric = {
    totalPoints: 20,
    expectedIntent: 'C wants a partial refund because the CP was 90 minutes late and did not clean the kitchen.',
    requiredCrmActions: 'Add an internal note summarizing the complaint\nIssue a 25% partial refund to C\nLog a lateness incident on the CP profile',
    requiredResolution: 'Approve a 25% partial refund and confirm it with C. Inform CP of the lateness incident.',
    expectedC: 'Apologize, confirm the 25% partial refund and when it will appear, thank C.',
    expectedCP: 'Let CP know about the lateness report and the kitchen complaint, and remind CP to message C when running late.',
    criticalErrors: 'Promising a full refund. Sharing C personal details with CP.',
    acceptableAlternatives: 'Offering a free re-clean of the kitchen instead of the refund is acceptable if C accepts it.',
    autoFailConditions: 'Shares C phone number or address with the CP.',
    criteria: [
      crit('Identifies the customer concern', 'intent', 3, 'Must mention both lateness and the uncleaned kitchen.'),
      crit('Required CRM actions', 'crm', 6, 'Two points per required action.', [{ description: 'Missing internal note', points: 2 }]),
      crit('Correct resolution', 'resolution', 5, '25% partial refund, or the accepted re-clean alternative.', [{ description: 'Offers full refund', points: 5 }]),
      crit('Response to C', 'c_comm', 3),
      crit('Message to CP', 'cp_comm', 3),
    ],
  };
  t1.references = {
    policyContext: 'FICTIONAL STAGING POLICY: Lateness over 60 minutes plus a missed area qualifies for a 25% partial refund.',
    kbReferences: 'Fictional KB: "Late arrival complaints" (staging).',
    expectedResolution: '25% partial refund, lateness incident on CP profile.',
    sampleC: "I'm sorry your cleaner arrived late and missed the kitchen. I've issued a 25% refund, which will show on your card in 5 to 7 business days. Thank you for letting us know.",
    sampleCP: 'Hi, C reported you arrived about 90 minutes late and the kitchen was not cleaned. Please message C directly if you are running late.',
    evaluatorInstructions: 'Do not deduct for not stating the refund timeline to C unless the response is otherwise vague.',
  };

  const t2 = core.newTicket();
  Object.assign(t2, {
    url: 'https://example.com/fictional-tickets/T-2002',
    description: 'FICTIONAL. C asks to move tomorrow\'s cleaning to next Monday. The CP has no Monday availability.',
  });
  t2.rubric = {
    totalPoints: 10,
    expectedIntent: 'C wants to reschedule tomorrow\'s cleaning to next Monday.',
    requiredCrmActions: 'Check CP availability for Monday\nReschedule the job to a Monday slot with a new CP or offer an alternative day with the same CP',
    requiredResolution: 'Explain that the current CP is unavailable Monday and offer either another CP on Monday or another day with the same CP.',
    expectedC: 'Offer both options clearly and ask C to choose.',
    expectedCP: 'N/A',
    criticalErrors: 'Cancelling the job without C agreement.',
    acceptableAlternatives: '',
    autoFailConditions: '',
    criteria: [
      crit('Identifies the request', 'intent', 2),
      crit('CRM actions', 'crm', 3),
      crit('Resolution offered', 'resolution', 3),
      crit('Response to C', 'c_comm', 2),
    ],
  };
  t2.references = {
    policyContext: 'FICTIONAL STAGING POLICY: Reschedules more than 24 hours ahead carry no fee.',
    kbReferences: '',
    expectedResolution: 'Two options offered: another CP on Monday, or same CP on a different day.',
    sampleC: 'Your current cleaner is not available Monday. I can book another cleaner on Monday, or keep your cleaner on Tuesday. Which would you prefer?',
    sampleCP: 'N/A',
    evaluatorInstructions: 'A message to CP is not needed for this ticket. If the trainee writes N/A for q4, that is correct.',
  };

  const now = new Date();
  const deadline = new Date(now.getTime() + 3 * 864e5).toISOString().slice(0, 10);
  return {
    id: core.uid('quiz'),
    title: 'STAGING Day 3 Ticket Assessment (fictional)',
    cohort: 'Staging Cohort A',
    trainingDay: 'Day 3',
    assessmentType: 'Ticket handling',
    assignedTrainees: cohortA,
    deadline,
    status: 'Draft',
    passingThreshold: 80,
    tickets: [t1, t2],
    form: null,
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
  };
}

// Sample answers: one strong, one weaker. Used by the tests and the "load sample submissions" button.
function sampleAnswers(quiz) {
  const [t1, t2] = quiz.tickets;
  return {
    strong: [
      { ticketId: t1.id, q1: 'C wants a partial refund because the CP was 90 minutes late and the kitchen was not cleaned.', q2: 'Add internal note summarizing the complaint\nIssue 25% partial refund to C\nLog lateness incident on CP profile', q3: "I'm so sorry your cleaner was late and missed the kitchen. I've issued a 25% partial refund; you'll see it in 5-7 business days. Thanks for telling us.", q4: 'Hi! C reported you arrived about 90 minutes late and the kitchen was not cleaned. Please message C if you are running late next time.' },
      { ticketId: t2.id, q1: "C wants to reschedule tomorrow's cleaning to next Monday.", q2: 'Checked CP availability for Monday (none)\nOffered C another CP on Monday or same CP on another day', q3: 'Your cleaner is not available Monday. I can book a different cleaner on Monday or keep your cleaner on Tuesday. Which do you prefer?', q4: 'N/A' },
    ],
    weak: [
      { ticketId: t1.id, q1: 'Customer is unhappy.', q2: 'Issue full refund', q3: 'Sorry about that, I gave you a full refund.', q4: 'N/A' },
      { ticketId: t2.id, q1: 'C wants to move the cleaning to Monday.', q2: 'Rescheduled to Monday', q3: 'Done, moved to Monday.', q4: 'N/A' },
    ],
  };
}

module.exports = { TRAINEES, sampleQuiz, sampleAnswers };
