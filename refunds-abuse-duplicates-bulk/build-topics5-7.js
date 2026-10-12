// ---------- Deck 9, topics 5–7 (merged into the live Refunds deck before its closing slide) ----------
const GREEN = '2E7D32', GREENSOFT = 'E4F2E1';

function doRows(s, x, w, y, items, h = 0.6, gap = 0.08) {
  items.forEach(([lead, rest], i) => {
    const yy = y + i * (h + gap);
    box(s, x, yy, w, h, C.white, C.border);
    badge(s, x + 0.14, yy + (h - 0.3) / 2, i + 1);
    T(s, [{ text: lead, options: { bold: true, fontFace: HEAD } }, ...(rest ? [{ text: '  ' + rest, options: { color: C.soft } }] : [])],
      { x: x + 0.56, y: yy, w: w - 0.7, h, fontSize: 9.5, valign: 'middle' });
  });
}
// Flow box with a bold title and a soft second line.
async function fbox(s, x, y, w, h, title, sub, fill, line, col) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 1 }, rectRadius: 0.08 });
  T(s, [{ text: title, options: { bold: true, fontFace: HEAD, color: col, breakLine: !!sub } }, ...(sub ? [{ text: sub, options: { fontSize: 8.5, color: col === C.white ? C.white : C.ink } }] : [])],
    { x: x + 0.08, y, w: w - 0.16, h, fontSize: 9.5, align: 'center', valign: 'middle', paraSpaceAfter: 2 });
}
const lbl = (s, x, y, t) => T(s, t, { x, y, w: 0.8, h: 0.2, fontSize: 8.5, bold: true, color: C.soft });

(async () => {
  topicNo = 4;

  // ================= 5. ETF CHARGE DECLINED =================
  const ED = 'ETF Charge Declined';
  await topic('ETF Charge Declined', 'C agreed to cancel and pay the ETF, but the charge was declined.',
    ['Why the membership is still active', 'Protecting C from a surprise MF', 'Retrying with a Timed Reminder', 'What to do if it fails again'],
    'ill_pay.png',
    'Customer Agreed to Pay the ETF, But the Charge Was Declined. Category: Retention and Membership Cancellation. A customer agrees to cancel their FC membership and pay the associated ETF, but when we attempt to charge it, the payment is declined, so the membership remains active. The cancellation isn\'t actually complete yet: the membership technically remains active until the ETF successfully processes.');

  {
    const s = content(ED, 'What is this?', 'The cancellation isn\'t complete until the ETF goes through.',
      'A customer agrees to cancel their FC membership and pay the associated ETF, but when we attempt to charge it, the payment is declined, so the membership remains active. The cancellation isn\'t actually complete yet: the membership technically remains active until the ETF successfully processes. Why this matters: if we don\'t proactively manage this, the customer can end up blindsided by a monthly fee they thought was already cancelled. They agreed to leave, not to keep paying. The goal is to protect the customer from that surprise while still giving the payment a real chance to go through.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiCreditCard', title: 'Still active', body: 'The ETF was declined, so FC stays active until it goes through.' },
      { ico: 'FiAlertTriangle', title: 'Avoid the surprise', body: 'C agreed to leave, not to keep paying. Don\'t let an MF blindside them.' },
      { ico: 'FiRefreshCw', title: 'Give it a real chance', body: 'Retry the ETF before the next MF date.' },
    ]);
    await tip(s, 3.45, 'Related:', 'ForeverClean Membership, ETF Waivers, MF & ETF Refund Requests.', 'FiBookOpen');
  }

  {
    const s = content(ED, 'The process', 'Tell C, set a TR, retry. Then confirm, or follow up.',
      '1) Tell the customer their FC membership is still active since the ETF was declined, and that they\'ll be charged another MF on the next billing date if the ETF charge doesn\'t succeed before then. 2) Set a Timed Reminder for at least 3 days before that next MF date, so there\'s time to retry the ETF charge before the customer risks another membership fee. 3) Once the TR triggers, attempt the cancellation and ETF charge again. If it succeeds: confirm the cancellation with the customer and close out the ticket. Don\'t leave it assumed: the customer agreed to leave because they expected this to be handled, so a quick confirmation closes the loop. If it still fails: consider calling the customer (if within business hours) to explain that the cancellation hasn\'t completed yet due to the declined payment. You can consider issuing a free week or two as a goodwill gesture, but be explicit with the customer about when their next MF will be charged if the membership stays active in the meantime.');
    const w = 2.85, gap = 0.275;
    const steps = [['Tell C', 'FC is still active. Next MF charges if the ETF doesn\'t go through first'], ['Set a TR', 'At least 3 days before the next MF date'], ['TR triggers', 'Retry the cancellation and the ETF charge']];
    for (let i = 0; i < 3; i++) {
      const x = 0.45 + i * (w + gap);
      await fbox(s, x, 1.42, w, 0.8, steps[i][0], steps[i][1], i === 2 ? C.goldSoft : C.white, i === 2 ? C.gold : C.border, i === 2 ? C.gold : C.ink);
      if (i < 2) flowLine(s, [[x + w, 1.82], [x + w + gap, 1.82]], true);
    }
    const cx = 0.45 + 2 * (w + gap) + w / 2;
    flowLine(s, [[cx, 2.22], [cx, 2.4]]); flowLine(s, [[2.675, 2.4], [cx, 2.4]]);
    flowLine(s, [[2.675, 2.4], [2.675, 2.58]], true); flowLine(s, [[7.325, 2.4], [7.325, 2.58]], true);
    lbl(s, 2.75, 2.38, 'Succeeds'); lbl(s, 7.4, 2.38, 'Still fails');
    headPanel(s, 0.45, 2.58, 4.45, 1.85, 'Confirm and close', ['Confirm the cancellation with C', 'Close out the ticket', 'Don\'t leave it assumed: a quick confirmation closes the loop'], GREEN, GREENSOFT);
    headPanel(s, 5.1, 2.58, 4.45, 1.85, 'Follow up', ['Consider calling C (business hours) to explain', 'Consider a free week or two as goodwill', 'Be explicit about when the next MF will be charged'], RED, PINK);
  }

  await knowledgeCheck(ED, [
    ['C agreed to cancel, but the ETF was declined. Is FC cancelled?', 'No. It stays active until the ETF successfully processes.'],
    ['When should the Timed Reminder fire?', 'At least 3 days before the next MF date.'],
    ['The retry succeeds. What\'s left to do?', 'Confirm the cancellation with C and close the ticket.'],
    ['The retry fails again. Options?', 'Call C (business hours); consider a free week or two, and say when the next MF charges.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(ED, [
    ['Is the cancellation complete?', 'Was the ETF actually charged?'],
    ['What actions are you going to take?', 'Tell C, TR timing, retry, follow-up.'],
    ['5-min writing challenge', 'Explain the declined ETF to C.'],
  ], 'TRAINER: share a live ticket where the ETF charge was declined after the customer agreed to cancel. Give trainees a few minutes to review it on their own, then work through it together. 1) Is the cancellation complete? 2) What actions are you going to take? Tell C FC is still active and the next MF will charge if the ETF doesn\'t succeed first; set a TR at least 3 days before the next MF; retry; confirm and close, or call and consider a free week or two. 3) 5-minute writing challenge: each trainee writes the message to C. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 6. FAILED CHARGE / OVERDUE MF =================
  const FC = 'Failed Charge / Overdue MF';
  await topic('Failed Charge or Overdue MF', 'What to check first, before you ask C to update their card.',
    ['When it comes up', 'Rule 1: don\'t chase payment while an issue is open', 'Rule 2: the upcoming-appointment exception', 'What counts as "resolved"'],
    'ill12.png',
    'Failed Job Charge or Overdue Membership Fee: What to Check First. Category: Retention and Membership Cancellation. This covers what to do the moment you notice a failed job charge or an overdue membership fee on an account, whether that\'s why the customer contacted you, or something you happened to spot while working on a different issue.');

  {
    const s = content(FC, 'What is this?', 'You spot money owed. Should you bring it up right now?',
      'This covers what to do the moment you notice a failed job charge or an overdue membership fee on an account, whether that\'s why the customer contacted you, or something you happened to spot while working on a different issue entirely. Why this matters: the instinct when you spot money owed is to fix it: ask the customer to update their card, get the charge current. But doing that at the wrong moment can make the interaction feel like we\'re chasing payment instead of taking care of the customer, especially if there\'s already an unresolved complaint on the account. Knowing when to raise it, and the one case where you should raise it proactively, matters more than the mechanics of fixing the charge itself. For how many overdue MFs to bring current, when to issue free months, and cancellation-intent branching, see Understanding Offer Logic > Free Months (Job 2). This article is the "should I even bring this up right now" gate that comes before that.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiEye', title: 'When it comes up', body: 'Why C contacted us, or spotted while working on something else.' },
      { ico: 'FiClock', title: 'Timing matters', body: 'At the wrong moment it feels like chasing payment, not taking care of C.' },
      { ico: 'FiFilter', title: 'A gate first', body: 'Then Understanding Offer Logic › Free Months (Job 2) for the counts.' },
    ]);
  }

  {
    const s = content(FC, 'What to check first', 'Protect the upcoming job. Otherwise, resolve the issue before the card.',
      'Rule 1: don\'t chase payment while an issue is still open. If the customer has an unresolved issue and a failed job charge or overdue MF, don\'t ask them to update their card so we can collect what\'s owed until the issue itself is resolved. Example: a customer raises a PCQ complaint, and you notice a failed job charge while investigating. Handle the PCQ first. Only bring up the card/payment once the PCQ is actually resolved. "Resolved" means one of two things: either we did what the customer asked, or the customer explicitly accepted our resolution. If neither has happened, treat it as still open: don\'t bring up the card, and don\'t send the link to update it. Rule 2: an upcoming appointment with an invalid card is the one exception. If the customer has an upcoming appointment and an invalid card on file, proactively advise them to update their card, otherwise they risk that appointment getting cancelled. This applies regardless of whether Rule 1\'s unresolved-issue condition is also in play: protecting the upcoming job takes priority.');
    await fbox(s, 3.2, 1.42, 3.6, 0.45, 'Failed job charge or overdue MF spotted', '', C.ink, C.ink, C.white);
    flowLine(s, [[5.0, 1.87], [5.0, 2.02]], true);
    await fbox(s, 3.0, 2.02, 4.0, 0.5, 'Upcoming appointment + invalid card?', 'Rule 2', C.goldSoft, C.gold, C.gold);
    flowLine(s, [[7.0, 2.27], [7.35, 2.27]], true); lbl(s, 7.03, 2.05, 'Yes');
    await fbox(s, 7.35, 1.95, 2.2, 0.65, 'Advise C to update the card now', 'Protect the job', GREENSOFT, GREEN, GREEN);
    flowLine(s, [[5.0, 2.52], [5.0, 2.7]], true); lbl(s, 5.08, 2.5, 'No');
    await fbox(s, 3.0, 2.7, 4.0, 0.5, 'Unresolved issue on the account?', 'Rule 1', C.goldSoft, C.gold, C.gold);
    flowLine(s, [[3.0, 2.95], [2.65, 2.95]], true); lbl(s, 2.68, 2.73, 'Yes');
    await fbox(s, 0.45, 2.6, 2.2, 0.7, 'Resolve the issue first', 'No card talk, no update link yet', PINK, RED, RED);
    flowLine(s, [[5.0, 3.2], [5.0, 3.38]], true); lbl(s, 5.08, 3.18, 'No');
    await fbox(s, 3.0, 3.38, 4.0, 0.55, 'OK to bring up the card / payment', 'Then: Offer Logic › Free Months (Job 2)', C.tealSoft, C.teal, C.teal);
    box(s, 0.45, 4.1, 9.1, 0.8, C.white, C.border);
    T(s, [{ text: '"Resolved" means:  ', options: { bold: true, fontFace: HEAD, color: C.teal } }, { text: 'we did what C asked, or C explicitly accepted our resolution. Neither? It\'s still open.', options: { breakLine: true } },
          { text: 'Example:  ', options: { bold: true, fontFace: HEAD, color: C.teal } }, { text: 'PCQ complaint + failed job charge: handle the PCQ first, the card after.', options: { color: C.soft } }],
      { x: 0.65, y: 4.1, w: 8.7, h: 0.8, fontSize: 9.5, valign: 'middle', paraSpaceAfter: 4 });
  }

  await knowledgeCheck(FC, [
    ['PCQ complaint still open + failed job charge. Ask C to update their card?', 'No. Resolve the PCQ first, then bring up the card.'],
    ['What counts as "resolved"?', 'We did what C asked, or C explicitly accepted our resolution.'],
    ['Open complaint, but C has a cleaning tomorrow and an invalid card. Raise it?', 'Yes. Rule 2: protecting the upcoming job takes priority.'],
    ['Where do you check how many overdue MFs to bring current?', 'Understanding Offer Logic › Free Months (Job 2).'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(FC, [
    ['Is there an upcoming job + invalid card?', 'Rule 2 comes first.'],
    ['Is an issue still unresolved?', 'Did we do what C asked, or did C accept?'],
    ['What actions are you going to take?', 'And when do you raise the payment?'],
  ], 'TRAINER: share a live ticket with a failed job charge or overdue MF (ideally alongside another issue). Give trainees a few minutes to review it on their own, then work through it together. 1) Is there an upcoming appointment with an invalid card? If so, advise C to update it now. 2) Is there an unresolved issue? Resolved means we did what C asked, or C explicitly accepted our resolution. 3) What actions are you going to take, and when do you raise the payment? Then open the floor for questions.');

  // ================= 7. GRACEFUL CLOSURE =================
  const GC = 'Graceful Closure';
  await topic('Graceful Closure', 'When the answer is no, and C keeps pushing.',
    ['The two entry-gate questions', 'The 4 steps', 'The No Re-Opening Rule', 'When to stop: legal threats and ETF waivers'],
    'ill9.png',
    'Graceful Closure: When the Answer Is No. Category: Retention and Membership Cancellation. This page inherits the voice standards defined in Care Philosophy & Communication Standards. Graceful Closure is what you do once a customer has heard your final answer, declined it, and keeps pushing anyway. It\'s the last step in a resolution, not an easy out for an uncomfortable conversation, and not something you reach for the moment a customer pushes back once.');

  {
    const s = content(GC, 'What is this?', 'The last step, once C has heard the final answer and keeps pushing.',
      'Without a defined last step, conversations go one of two bad ways: agents keep negotiating indefinitely (which erodes the offers we\'ve made policy-wide and turns every ticket into a precedent), or they cut the conversation short before understanding what the customer wants, mistaking silence or a declined offer for "nothing left to try." Graceful Closure gives agents a clean, respectful way to end the conversation, but only after the entry gates are cleared, because once you\'re in it, you can\'t reopen or add concessions (the No Re-Opening Rule). That\'s what makes it a one-way door. Policy: if the customer declines our resolution and continues to insist, we move to Graceful Closure: clearly and calmly restate the final position, acknowledge the customer\'s perspective and disappointment, avoid repetitive back-and-forth, set expectations respectfully, and close the conversation professionally and confidently.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiFlag', title: 'The last step', body: 'Not an easy out, and not after the first pushback.' },
      { ico: 'FiLock', title: 'A one-way door', body: 'Once you\'re in it: no re-opening, no new concessions.' },
      { ico: 'FiMessageCircle', title: 'Calm and clear', body: 'Restate, acknowledge, set expectations, close professionally.' },
    ]);
    await warn(s, 3.45, 'Two bad ways:', 'negotiating forever, or closing before you know what C actually wants.');
  }

  {
    const s = content(GC, 'Before you enter: two questions', 'Both must be "yes" before Step 1.',
      'Ask yourself two questions first. 1) Have I actually offered everything I could here based on the root cause? If not, go back and make the right offer first: check Understanding Offer Logic if you\'re not sure what that is. 2) Do I actually know why this customer wants to cancel or won\'t take the offer? If not, you\'re not done digging: go back and ask, don\'t close. The one that trips people up the most: if a customer doesn\'t answer your question, that doesn\'t mean you\'ve run out of options; it means you don\'t have enough information yet. Silence isn\'t a green light to move to Graceful Closure and close the ticket. It\'s a sign to ask again, in a different way.');
    await fbox(s, 2.0, 1.42, 4.6, 0.55, 'Offered everything based on the root cause?', 'Gate 1', C.goldSoft, C.gold, C.gold);
    flowLine(s, [[6.6, 1.695], [6.95, 1.695]], true); lbl(s, 6.63, 1.47, 'No');
    await fbox(s, 6.95, 1.42, 2.6, 0.55, 'Make the right offer first', 'See Understanding Offer Logic', PINK, RED, RED);
    flowLine(s, [[4.3, 1.97], [4.3, 2.17]], true); lbl(s, 4.38, 1.95, 'Yes');
    await fbox(s, 2.0, 2.17, 4.6, 0.55, 'Know why C wants to cancel / won\'t accept?', 'Gate 2', C.goldSoft, C.gold, C.gold);
    flowLine(s, [[6.6, 2.445], [6.95, 2.445]], true); lbl(s, 6.63, 2.22, 'No');
    await fbox(s, 6.95, 2.17, 2.6, 0.55, 'Keep digging', 'Ask again, differently', PINK, RED, RED);
    flowLine(s, [[4.3, 2.72], [4.3, 2.92]], true); lbl(s, 4.38, 2.7, 'Yes');
    await fbox(s, 2.0, 2.92, 4.6, 0.5, 'Graceful Closure: Step 1', '', C.teal, C.teal, C.white);
    await warn(s, 3.7, 'Silence isn\'t a green light.', 'No answer = not enough information yet. Ask again, in a different way.', 0.5);
  }

  {
    const s = content(GC, 'The 4 steps', 'Clarify, Educate, Offer, Protect.',
      'Step 1, Clarify: state the final decision clearly. Be direct, neutral and unambiguous: "I\'m not able to waive the cancellation fee." Step 2, Educate: explain the policy and the next steps. "This fee exists to ensure the cleaner is paid fairly for the heavily discounted first cleaning in case the membership is canceled early. Your membership will remain active and the monthly fees will continue unless it\'s cancelled." Step 3, Offer: present approved options (pre-approved recovery, retention, or alternative solutions). "While I can\'t fully waive the fee, since you have already paid for [X] months of membership, I can apply those payments as a credit against the Early Termination Fee. This reduces your cancellation fee from $105 to [reduced amount], so you aren\'t paying the full fee on top of what you\'ve already spent." Step 4, Protect: close professionally and keep boundaries. Signal the conversation is complete while acknowledging the customer\'s feelings; stand firm if the customer repeats the same concern without new information. "I understand this isn\'t the outcome you were hoping for. If there\'s nothing else I can assist with, I\'ll consider this conversation complete on my end." or "I understand this isn\'t the outcome you were hoping for. We\'ve reviewed all available options, and this is the final resolution we can provide." Customers may still feel disappointed; that\'s okay when the decision is correct and clearly communicated.');
    const w = (9.1 - 0.45) / 4;
    const st = [
      ['Clarify', 'State the final decision. Direct, neutral, unambiguous.', 'I\'m not able to waive the cancellation fee.'],
      ['Educate', 'Explain the policy and what happens next.', 'Your membership will remain active and the monthly fees will continue unless it\'s cancelled.'],
      ['Offer', 'Present pre-approved options.', 'I can apply what you\'ve already paid as a credit against the ETF.'],
      ['Protect', 'Close professionally. Hold the boundary.', 'We\'ve reviewed all available options, and this is the final resolution we can provide.'],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.45 + i * (w + 0.15);
      box(s, x, 1.45, w, 3.05, C.white, C.border);
      badge(s, x + 0.15, 1.6, i + 1);
      T(s, st[i][0], { x: x + 0.55, y: 1.6, w: w - 0.65, h: 0.3, fontFace: HEAD, bold: true, fontSize: 12, valign: 'middle' });
      T(s, st[i][1], { x: x + 0.15, y: 2.02, w: w - 0.3, h: 0.7, fontSize: 9, color: C.soft });
      box(s, x + 0.12, 2.8, w - 0.24, 1.55, C.tealSoft);
      T(s, '“' + st[i][2] + '”', { x: x + 0.22, y: 2.85, w: w - 0.44, h: 1.45, fontSize: 9, italic: true, valign: 'middle' });
    }
    await tip(s, 4.65, 'Still disappointed?', 'that\'s okay when the decision is correct and clearly communicated.', 'FiHeart');
  }

  {
    const s = content(GC, 'The No Re-Opening Rule', 'Once all final options are on the table.',
      'Once all final options have been presented: Do not introduce new concessions: adding benefits/exceptions in response to repeated requests creates an expectation that agents will always offer more, undermining policy and fairness. C: "Can you also refund last month\'s charge?" Agent: "We\'ve already reviewed all available options, and we cannot provide additional refunds beyond what was offered." Do not reframe repeatedly: re-explaining the same decision in different words can confuse the customer, look inconsistent, and prolong the conversation. C: "But why not just waive the fee?" Agent: "Our company does not allow waiving this fee to make sure cleaners are paid fairly. This is the final resolution." Do not re-enter the negotiation loop: continued bargaining after closure frustrates both sides and leads to inconsistent outcomes. C: "I still want a refund. Can you make an exception?" Agent: "We\'ve looked at all the options we can offer, and this is the final resolution. I can\'t make any more changes." Two Attempts Rule: make two clear statements of closure, including the offered options and final resolution. If the customer repeats the same concern without new information after the second attempt, the ticket may be marked complete.');
    table(s, ['Don\'t', 'Why', 'Example reply'], [
      ['Introduce new concessions', 'Sets the expectation agents will always offer more', '"We\'ve already reviewed all available options, and we cannot provide additional refunds beyond what was offered."'],
      ['Reframe repeatedly', 'Confusing, inconsistent, drags it out', '"Our company does not allow waiving this fee to make sure cleaners are paid fairly. This is the final resolution."'],
      ['Re-enter the negotiation loop', 'Frustrating, inconsistent outcomes', '"We\'ve looked at all the options we can offer, and this is the final resolution. I can\'t make any more changes."'],
    ], { y: 1.45, colW: [2.2, 2.6, 4.3], fontSize: 9, rowH: 0.7 });
    await tip(s, 4.45, 'Two Attempts Rule:', 'two clear closure statements. Same concern, nothing new after the 2nd? Mark it complete.', 'FiCheckCircle');
  }

  {
    const s = content(GC, 'Exceptions: when Graceful Closure stops', 'Closed in error, a legal threat, or an ETF waiver reason.',
      'The No Re-Opening Rule applies to a correctly reached closure. It doesn\'t protect a closure invoked in error. If a previous agent moved to Graceful Closure without exhausting options, or before the customer\'s underlying concern was understood, correcting it isn\'t re-opening: it\'s fixing a wrong resolution. Prior agent closed but a valid retention/exit offer was never presented: go back and offer it now (not a new concession, the correct one that was missed). Prior agent closed after the customer didn\'t answer a probing question: re-open the investigation, probe again, then follow normal offer logic before considering closure again. Prior agent closed correctly and the customer is simply pushing again: stand firm. Material legal threat (Attorney General, cease-and-desist, class-action or litigation threats, small claims threats; see ETF Waivers) at any point, even after closure: stop Graceful Closure immediately. If C wants to cancel and be refunded: do it, no retention attempt, no further steps. Charges older than 120 days need Legal & Compliance review: raise it to your TL. Charges within 120 days can be refunded directly. ETF waiver reasons override Graceful Closure too: if C states a qualifying ETF waiver reason at any point, stop and cancel under it.');
    eyebrow(s, 'A previous closure', 0.45, 1.42, 4);
    doRows(s, 0.45, 5.0, 1.65, [
      ['Valid offer never presented?', 'Offer it now. It\'s the missed one, not a new concession.'],
      ['Closed after C didn\'t answer?', 'Re-open: probe again, then normal offer logic.'],
      ['Closed correctly, C pushing again?', 'Stand firm. Don\'t reopen.'],
    ], 0.78, 0.1);
    headPanel(s, 5.6, 1.45, 3.95, 1.55, 'Material legal threat', ['Stop Graceful Closure, even after it\'s reached', 'Cancel and refund. No retention', 'Over 120 days? Legal & Compliance via TL'], RED, PINK, 9);
    headPanel(s, 5.6, 3.15, 3.95, 1.35, 'ETF waiver reason', ['Stops Graceful Closure at any point', 'Cancel under the waiver. No second "no" needed'], C.gold, C.goldSoft, 9);
  }

  await knowledgeCheck(GC, [
    ['C declined your first offer. Move to Graceful Closure?', 'Not yet. Check both gates: offered everything? Know why?'],
    ['C didn\'t answer your probing question. Close?', 'No. Silence isn\'t a green light. Ask again, differently.'],
    ['After closure: "Can you also refund last month?" Reply?', 'No new concessions: all options were reviewed; this is the final resolution.'],
    ['After closure, C threatens a class action. What now?', 'Stop Graceful Closure. Cancel and refund (over 120 days: TL for Legal).'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(GC, [
    ['Are both gates cleared?', 'Everything offered? Root cause known?'],
    ['Which step are you on?', 'Clarify, Educate, Offer, Protect.'],
    ['5-min writing challenge', 'Write your closure message.'],
  ], 'TRAINER: share a live ticket where the customer keeps pushing after a final answer. Give trainees a few minutes to review it on their own, then work through it together. 1) Are both entry gates cleared (everything offered based on the root cause; the reason for declining is known)? 2) Which step are you on: Clarify, Educate, Offer, Protect? Check the No Re-Opening Rule, the Two Attempts Rule, and whether a legal threat or ETF waiver reason overrides closure. 3) 5-minute writing challenge: each trainee writes the closure message. Read a few aloud and compare. Then open the floor for questions.');

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
