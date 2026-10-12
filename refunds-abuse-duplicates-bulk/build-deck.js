// ---------- Deck 9: MF & ETF Refunds, Abusive Behavior, Duplicate Accounts, Bulk Charges ----------
const GREEN = '2E7D32', GREENSOFT = 'E4F2E1';

// Numbered "do" rows: white card, badge, bold lead + detail.
function doRows(s, x, w, y, items, h = 0.6, gap = 0.08) {
  items.forEach(([lead, rest], i) => {
    const yy = y + i * (h + gap);
    box(s, x, yy, w, h, C.white, C.border);
    badge(s, x + 0.14, yy + (h - 0.3) / 2, i + 1);
    T(s, [{ text: lead, options: { bold: true, fontFace: HEAD } }, ...(rest ? [{ text: '  ' + rest, options: { color: C.soft } }] : [])],
      { x: x + 0.56, y: yy, w: w - 0.7, h, fontSize: 9.5, valign: 'middle' });
  });
}

// Quoted example lines (customer messages) in a soft box.
function quotes(s, x, y, w, h, label, lines, col, soft) {
  box(s, x, y, w, h, soft, col);
  T(s, [{ text: label, options: { bold: true, fontFace: HEAD, fontSize: 10.5, color: col, breakLine: true } },
        ...lines.map((t, i) => ({ text: '“' + t + '”', options: { italic: true, fontSize: 9.5, color: C.ink, breakLine: i < lines.length - 1 } }))],
    { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: h - 0.24, paraSpaceAfter: 6 });
}

(async () => {
  // ---------- Cover ----------
  {
    const s = pres.addSlide({ masterName: 'COVER' });
    eyebrow(s, 'Homeaglow  ·  New Hire Care Training', 0.62, 0.6);
    T(s, 'Refunds, Abuse,\nDuplicates &\nBulk Charges', { x: 0.62, y: 0.95, w: 4.9, h: 1.75, fontFace: HEAD, bold: true, fontSize: 30 });
    T(s, 'MF and ETF refund requests, abusive customers, duplicate accounts and bulk MF charges.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers four Knowledge Library articles from the Retention and Membership Cancellation category: MF & ETF Refund Requests; Customer Abusive Behavior Towards a Cleaner; Duplicate Accounts; and Bulk Charges. The thread through all four: work out the real root cause first (what happened, and whose error it was), then apply the policy that fits it, no more and no less.');
  }

  // ---------- Agenda ----------
  const agendaLinks = [], topicSlides = [];
  let agendaSlide;
  const addTopic = topic;
  topic = async (...args) => { const t = await addTopic(...args); topicSlides.push(t._slideNum); return t; };
  {
    const s = agendaSlide = content('Overview', 'Agenda', 'Four topics. Click a topic to jump to it.',
      'Topic 1 covers MF and ETF refund requests: the root cause drives the decision, eligible doesn\'t automatically mean a cash refund, and billing errors are fixed without waiting to be asked. Topic 2 covers customers confirmed abusive toward a cleaner: no retention, only a fair exit. Topic 3 covers duplicate accounts: check Thumbtack, cancel the duplicate FC, then offer a refund or free months. Topic 4 covers bulk MF charges: which of the three branches you\'re in, and what to do in each. In slideshow mode, click any topic to jump straight to it.');
    const rows = [
      ['MF & ETF Refund Requests', 'The root cause drives the decision.'],
      ['Customer Abusive Behavior', 'Protect the CP. No retention, a fair exit.'],
      ['Duplicate Accounts', 'Stop the double charge, then make them whole.'],
      ['Bulk Charges', 'Several missed MFs charged on one day.'],
    ];
    const cw = (9.1 - 0.15) / 2, h = 1.2, gy = 0.15;
    for (let i = 0; i < rows.length; i++) {
      const [title, line] = rows[i], x = 0.45 + (i % 2) * (cw + 0.15), y = 1.45 + Math.floor(i / 2) * (h + gy);
      const link = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(link);
      s.addText([
        { text: title, options: { color: C.ink, bold: true, fontFace: HEAD, fontSize: 13, breakLine: true, hyperlink: link } },
        { text: line, options: { color: C.soft, fontSize: 10, hyperlink: link } },
      ], { shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w: cw, h, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.border, width: 0.75 },
        fontFace: BODY, valign: 'middle', margin: [56, 30, 0, 0], hyperlink: link });
      const num = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(num);
      s.addText(String(i + 1).padStart(2, '0'), { x: x + 0.1, y, w: 0.55, h, fontFace: HEAD, bold: true, fontSize: 16, color: C.teal, align: 'center', valign: 'middle', hyperlink: num });
      const arrow = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(arrow);
      s.addText('›', { x: x + cw - 0.42, y, w: 0.35, h, fontFace: HEAD, bold: true, fontSize: 20, color: C.teal, align: 'center', valign: 'middle', hyperlink: arrow });
    }
  }

  // ================= 1. MF & ETF REFUND REQUESTS =================
  const RF = 'MF & ETF Refunds';
  await topic('MF & ETF Refund Requests', 'Refund requests for Monthly Fees and the Early Termination Fee.',
    ['Why the root cause drives the decision', 'Eligible doesn\'t mean cash: free months first', 'MF refunds by root cause, and ETF refunds', 'Billing errors, high-risk and hardship cases'],
    'ill_pay.png',
    'MF & ETF Refund Requests. Category: Retention and Membership Cancellation. Covers refund requests for Monthly Fees (MF) and/or the Early Termination Fee (ETF), whether the customer is mid-cancellation, already cancelled, or just contesting a charge. This is separate from job-charge refunds (Premium Fee, Lockout, LMC: see the AG-Refund pages).');

  {
    const s = content(RF, 'What is this?', 'Refunds for MFs and the ETF. Not job charges (Premium Fee, Lockout, LMC).',
      'Why this matters: not every MF or ETF refund request has the same root cause, and the root cause is what should drive the decision, not just "how long ago was the charge." A customer who was never told about the membership deserves a different answer than one who knew and simply forgot to cancel in time, even if both are asking for the same thing. Getting this right protects against two failure modes: reflexively refusing refunds that are genuinely our error, and reflexively granting refunds that reward not paying attention. No refund request, no refund, but a request can be implied as well as explicit: the customer doesn\'t need to use the word "refund"; contesting a charge with the expectation it be reversed is a request. The exception is our own error: where we charged in error (a duplicate charge, a charge the customer should never have had), correct it without waiting to be asked.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiSearch', title: 'Root cause decides', body: 'Never told about the membership vs. knew and forgot: same ask, different answer.' },
      { ico: 'FiMessageCircle', title: 'No request, no refund', body: 'A request can be implied. Contesting a charge to get it reversed counts.' },
      { ico: 'FiAlertCircle', title: 'Our error? Fix it', body: 'Duplicate or wrong charges: correct them without waiting to be asked.' },
    ]);
    await warn(s, 3.4, 'Two failure modes:', 'refusing refunds that are our error, and granting refunds that reward not paying attention.', 0.5);
  }

  {
    const s = content(RF, 'Eligible doesn\'t mean cash', 'Eligible = a legitimate claim to the value of those unused MFs.',
      'READ THIS FIRST: "Eligible" means the customer has a legitimate claim to the value of those unused MFs. It doesn\'t mean the default resolution is a refund. If FC is still active and the customer is retainable, lead with free months equal to the eligible amount (capped at 3). Refund in cash if: FC has already been cancelled, the customer declines the free months, or the customer isn\'t retainable and the route is to cancel FC. This doesn\'t apply to the agent-error row (cancellation requested after MCT but FC kept renewing), or to Billing Errors / High-Risk Legal Situations: those are refunded directly regardless of active status.');
    await flowBox(s, 3.25, 1.42, 3.5, 0.55, 'Eligible for an MF refund', 'decision');
    flowSplit(s, 5.0, 1.97, 2.17, 2.675, 7.325, 2.4, 'Yes', 'No');
    T(s, 'FC active and C retainable?', { x: 5.1, y: 1.98, w: 2.5, h: 0.2, fontSize: 8.5, bold: true, color: C.soft });
    headPanel(s, 0.45, 2.4, 4.45, 1.5, 'Lead with free months', ['Equal to the eligible amount', 'Capped at 3 free months'], C.teal, C.tealSoft);
    headPanel(s, 5.1, 2.4, 4.45, 1.5, 'Refund in cash', ['FC already cancelled', 'C declines the free months', 'Not retainable: the route is cancelling FC'], C.gold, C.goldSoft);
    await warn(s, 4.1, 'Refund directly, any FC status:', 'agent error after MCT, billing errors, high-risk / legal situations.', 0.5);
  }

  {
    const s = content(RF, 'The MCT test', 'An unused MF isn\'t a refund reason on its own. The fee buys membership, not attendance.',
      'An MF going unused is not, on its own, a reason to refund it. The test is paid months minus the customer\'s MCT. A customer can go months without booking and still owe every one of those MFs, because the fee buys membership, not attendance. Inside the MCT, an unused month is owed: refund it only under a named exemption (a material legal threat, extreme hardship, compassion, or a history of disputes). Beyond the MCT, an unused MF is refundable when all of these hold: it was charged in the last 120 days; it is still disputable; and the reason is that the customer was unaware of the membership or believed they had already cancelled, not that they simply forgot to cancel in time. Example: C has 7 paid months against a 6-month MCT and three MFs went unused. Only one of those sits beyond the MCT. The other two are owed: they went unused inside the commitment term, which is not a refund ground.');
    headPanel(s, 0.45, 1.45, 4.45, 1.75, 'Inside the MCT: owed', ['Refund only under a named exemption:', 'Material legal threat, extreme hardship, compassion, or a dispute history'], RED, PINK);
    headPanel(s, 5.1, 1.45, 4.45, 1.75, 'Beyond the MCT: refundable if ALL', ['Charged in the last 120 days', 'Still disputable', 'C was unaware, or thought they\'d cancelled (not "forgot")'], C.teal, C.tealSoft);
    eyebrow(s, 'Example', 0.45, 3.38, 3, C.soft);
    const w = (9.1 - 0.3) / 4;
    stat(s, 0.45, 3.6, w, 1.25, '7', 'paid months');
    stat(s, 0.45 + (w + 0.1), 3.6, w, 1.25, '6', 'month MCT');
    stat(s, 0.45 + 2 * (w + 0.1), 3.6, w, 1.25, '3', 'unused MFs');
    stat(s, 0.45 + 3 * (w + 0.1), 3.6, w, 1.25, '1', 'refundable: the only one beyond MCT', C.gold);
  }

  {
    const s = content(RF, 'MF refunds by root cause', 'Unused MFs beyond MCT, within 120 days, unless noted.',
      'MF refunds sorted by root cause. Unaware of the membership, charged multiple MFs: eligible for unused MFs beyond MCT that are still disputable (within 120 days). Thought they\'d cancelled (cancelled RC/upcoming jobs, not FC): same. C requested cancellation to Support after completing MCT, but FC continued renewing monthly without their explicit agreement: MFs charged after the cancellation request are eligible regardless of the 120-day window, since this is an agent error. If C has completed MCT and requested cancellation, don\'t charge another MF without their explicit agreement to continue; the agent should have created a Timed Reminder to cancel FC before the next renewal if C didn\'t respond to the retention attempt. Aware of the membership, just forgot to cancel in time: generally no. The membership benefit (discounted access) was available every month they were charged; forgetting to cancel isn\'t a service failure or communication gap on our side, and refunding would mean refunding based on non-use. Duplicate account, duplicate MF charged: the duplicate charges within 120 days are eligible if the customer declined the free months. C already paid the ETF, now wants an MF refund too, no escalation language: no ("I canceled and paid my ETF. Can I have a refund on the MF I just paid?"). With escalating/dispute language or legal threats: yes, the MF may be refundable to avoid further costs ("Either refund the MF or the ETF or I will be contacting my bank and disputing both of these charges"). C has a previous bank dispute: refund on request as long as it\'s still disputable (within 120 days). Exception: a successful dispute on a charge that should have been non-disputable opens everything charged after it; if the customer successfully disputed a charge already past 120 days, every later charge is also refundable on request.');
    table(s, ['Root cause', 'Eligible?'], [
      ['Unaware of the membership', 'Yes: unused MFs beyond MCT, within 120 days'],
      ['Thought they\'d cancelled (cancelled RC/jobs, not FC)', 'Yes: unused MFs beyond MCT, within 120 days'],
      ['Asked to cancel after MCT, FC kept renewing', 'Yes: every MF after the request, any age (agent error)'],
      ['Knew about it, forgot to cancel', 'Generally no. The benefit was available every month'],
      ['Duplicate account, duplicate MF', 'Yes, within 120 days, if C declines free months'],
      ['Paid the ETF, now wants the MF back', 'No. With dispute or legal language: yes'],
      ['Has a previous bank dispute', 'Yes, while still disputable (120 days)'],
    ], { y: 1.45, colW: [4.1, 5.0], fontSize: 9, rowH: 0.4 });
    await tip(s, 4.75, 'Exception:', 'a successful dispute past 120 days opens every charge after it to refund.', 'FiUnlock');
  }

  {
    const s = content(RF, 'ETF refunds', 'Full or partial, depending on the situation.',
      'ETF refunds. Full refund if any of: the customer has a history of filing a dispute, asked for a refund before/after filing one, and no longer wants the service; dispute risk is above the percentage-based threshold; extenuating circumstances apply, especially if they never used the service or were unhappy with it (check the Extenuating Circumstances article: Financial Hardship, Military, Illness, Death, ADA, Disaster). Partial refund if: the ETF was miscalculated (system bug, overcharged hours); or the customer wants the voucher price/MF/card charges deducted for hours a DHJ voucher didn\'t cover, and is using escalating language.');
    headPanel(s, 0.45, 1.45, 4.45, 2.6, 'Full refund, if any', ['Dispute history + asked for a refund + no longer wants the service', 'Dispute risk above the percentage-based threshold', 'Extenuating circumstances, especially if they never used it or were unhappy'], C.teal, C.tealSoft);
    headPanel(s, 5.1, 1.45, 4.45, 2.6, 'Partial refund, if', ['ETF miscalculated (system bug, overcharged hours)', 'Wants DHJ voucher price / MF / card charges deducted for uncovered hours, with escalating language'], C.gold, C.goldSoft);
    await tip(s, 4.25, 'Extenuating circumstances:', 'Financial Hardship, Military, Illness, Death, ADA, Disaster.', 'FiBookOpen');
  }

  {
    const s = content(RF, 'Billing errors & high-risk cases', 'Our mistake, or a real legal or dispute risk.',
      'Billing errors (MF or ETF): refund immediately, no 120-day window, since this is our mistake, not a judgment call: a charge applied incorrectly or an MF charged twice by error; a promised discount never applied; we said we\'d cancel after MCT and didn\'t (refund anything charged after MCT regardless of the 120-day window: agent error). High-risk / material legal situations, refund if any apply: material legal threat (Attorney General, class action, etc.): refund charges within 120 days; beyond that needs Legal and Compliance Team approval, so escalate to your TL. Dispute risk above the percentage threshold: the 120-day limit still applies. If C has already filed a dispute, the 120-day limit still applies, even if they request a refund before or after filing, unless the bank allows a dispute for a charge that is normally non-disputable. If none of the above apply and dispute risk is low: don\'t refund. If the membership is still active, offer the free-months-in-lieu alternative instead (unused MF value converted to free months, capped at 3: see General Retention Playbook).');
    headPanel(s, 0.45, 1.45, 4.45, 2.55, 'Billing errors: refund now', ['Charged incorrectly, or an MF charged twice', 'A promised discount never applied', 'We said we\'d cancel after MCT and didn\'t', 'No 120-day window'], C.teal, C.tealSoft);
    headPanel(s, 5.1, 1.45, 4.45, 2.55, 'High-risk / legal', ['Material legal threat: refund within 120 days; beyond needs Legal approval (via TL)', 'Dispute risk above threshold: 120 days', 'Already filed a dispute: 120 days still applies'], RED, PINK, 9);
    await tip(s, 4.2, 'None apply, low dispute risk?', 'no refund. FC active? Offer free months in lieu (max 3).', 'FiGift');
  }

  {
    const s = content(RF, 'Compassion & hardship', 'Within or beyond MCT. A human, personal approach.',
      'Compassion / hardship cases (extenuating circumstances) apply whether the customer is within or beyond MCT. Take a human, personalized approach rather than running the standard scenario logic. If the customer doesn\'t specify how many MFs they want refunded, start with 1 as the baseline; don\'t offer more than asked. Refunds beyond the normal 120-day/count limits are allowed when the circumstances justify it; in severe cases, multiple MFs or a refund beyond the normal window is fine, but it needs support approval. Related macros: [Fragment] No refund — monthly fee; [Fragment] No refund — ETF; [Fragment] Lower ETF offer – MF deducted from ETF; [Fragment] Lower ETF offer – Voucher price deducted from ETF; [Fragment] Free month offer.');
    await cards(s, 1.45, 1.85, [
      { ico: 'FiHeart', title: 'Human first', body: 'Personal approach, not the standard scenario logic.' },
      { ico: 'FiHash', title: 'Start with 1 MF', body: 'If C doesn\'t say how many, start with 1. Don\'t offer more than asked.' },
      { ico: 'FiCheckCircle', title: 'Beyond the limits', body: 'Severe cases: multiple MFs or beyond 120 days, with support approval.' },
    ]);
    box(s, 0.45, 3.5, 9.1, 0.95, C.goldSoft, C.gold);
    s.addImage({ data: await icon('FiMessageSquare', C.gold), x: 0.63, y: 3.65, w: 0.18, h: 0.18 });
    T(s, [{ text: 'Macros', options: { bold: true, fontFace: HEAD, color: C.gold, breakLine: true } },
          { text: '[Fragment] No refund — monthly fee   ·   [Fragment] No refund — ETF   ·   [Fragment] Free month offer', options: { breakLine: true } },
          { text: '[Fragment] Lower ETF offer – MF deducted from ETF   ·   [Fragment] Lower ETF offer – Voucher price deducted from ETF' }],
      { x: 0.9, y: 3.56, w: 8.5, h: 0.85, fontSize: 9.5, paraSpaceAfter: 4 });
  }

  await knowledgeCheck(RF, [
    ['7 paid months, 6-month MCT, 3 unused MFs (all within 120 days, C was unaware). How many are refundable?', 'One: only the MF beyond the MCT. The other two are owed.'],
    ['C is eligible, FC is active and C is retainable. Refund?', 'Lead with free months equal to the eligible amount (max 3).'],
    ['"I knew about the membership, I just forgot to cancel." Refund?', 'Generally no. The benefit was available every month.'],
    ['C asked to cancel after MCT; FC kept renewing for 5 months. Refund?', 'Yes, every MF after the request, any age: agent error.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(RF, [
    ['What\'s the root cause?', 'Unaware, thought cancelled, forgot, our error?'],
    ['What actions are you going to take?', 'MCT test, 120 days, free months or refund?'],
    ['5-min writing challenge', 'Write your reply to the customer.'],
  ], 'TRAINER: share a live ticket with an MF or ETF refund request. Give trainees a few minutes to review it on their own, then work through it together. 1) What\'s the root cause: unaware of the membership, thought they\'d cancelled, forgot to cancel, a billing or agent error, a legal threat or dispute risk, or hardship? 2) What actions are you going to take? Apply the MCT test (paid months minus MCT) and the 120-day/disputable check; if eligible and FC is active and C is retainable, lead with free months (max 3); otherwise refund. 3) 5-minute writing challenge: each trainee writes the reply. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 2. CUSTOMER ABUSIVE BEHAVIOR =================
  const AB = 'Abusive Behavior';
  await topic('Customer Abusive Behavior', 'Towards a cleaner. Once abuse is confirmed, the goal shifts from retention to protecting the CP.',
    ['What counts as abusive (and what doesn\'t)', 'Why retention is off the table', 'The process', 'What\'s allowed vs. not allowed'],
    'ill30.png',
    'Customer Abusive Behavior Towards a Cleaner. Category: Retention and Membership Cancellation. Covers how to handle a cancellation or ETF conversation once a customer has been confirmed abusive toward their cleaner: personal, threatening, or intentionally harmful hostility, not ordinary frustration. It governs what offers are and aren\'t allowed once abuse is confirmed. It doesn\'t cover the CP-facing investigation or consequences: see Trust and Safety Overview for that side.');

  {
    const s = content(AB, 'What is this?', 'Abuse is confirmed. Retention stops; the CP comes first.',
      'Almost everything else in the Knowledge Library is about finding the right offer to keep the relationship. This article is the deliberate exception. Once a customer has been abusive toward a cleaner, the goal shifts from retention to protecting the CP and holding a clear boundary. Extending the normal retention toolkit here would reward the exact behavior we need to discourage, and would ask this CP (or the next one) to absorb more of it. The customer still gets a fair, policy-consistent exit: an ETF or MCT reduction is still on the table if it\'s warranted on its own terms, but nothing beyond that. The CP-side investigation and consequences are covered in Trust and Safety Overview.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiShield', title: 'Protect the CP', body: 'Retention would reward the behavior, and ask a CP to absorb more of it.' },
      { ico: 'FiSlash', title: 'Hold the boundary', body: 'No incentives, no persuasion to continue.' },
      { ico: 'FiLogOut', title: 'A fair exit', body: 'ETF or MCT reduction if warranted on its own terms. Nothing beyond that.' },
    ]);
    await tip(s, 3.45, 'CP-side investigation:', 'see Trust and Safety Overview. This topic is the customer side.', 'FiBookOpen');
  }

  {
    const s = content(AB, 'Intent, target, tone', 'Not just profanity. Is it aimed at the cleaner, personally?',
      'The distinction that matters most: intent, target, and tone, not just profanity. Someone venting about a bad outcome ("this is f***ing frustrating, my place still isn\'t clean") hasn\'t done anything actionable. Someone directing insults or threats at the cleaner personally has. Swearing alone is never enough to escalate on its own. Abusive (actionable): "You\'re f***ing useless. If you don\'t fix this, you\'ll regret it."; "I\'m going to come back there and hurt you."; repeatedly calling the cleaner names or directing profanity at them despite being asked to stop. Not abusive (don\'t escalate solely on this): "This is f***ing frustrating. My place still isn\'t clean."; "I\'m really pissed about how this turned out," without insults or threats aimed at the cleaner.');
    quotes(s, 0.45, 1.45, 4.45, 2.4, 'Abusive: actionable', ['You\'re f***ing useless. If you don\'t fix this, you\'ll regret it.', 'I\'m going to come back there and hurt you.', 'Repeated name-calling or profanity aimed at the cleaner, after being asked to stop'], RED, PINK);
    quotes(s, 5.1, 1.45, 4.45, 2.4, 'Not abusive: don\'t escalate on this alone', ['This is f***ing frustrating. My place still isn\'t clean.', 'I\'m really pissed about how this turned out.'], GREEN, GREENSOFT);
    await warn(s, 4.05, 'Swearing alone is never enough.', 'Look for insults or threats aimed at the cleaner personally.');
  }

  {
    const s = content(AB, 'The process', 'Explain the terms, keep it neutral, offer a clean exit.',
      '1) Do not attempt retention: no incentives, no persuasion to continue. 2) Explain membership and cancellation terms: clearly walk through FC details and ETF terms, factual and neutral, not an opening for negotiation. 3) Handling ETF objections: if the customer refuses to pay the ETF, an ETF or MCT reduction may still be considered using standard reduction guidelines and discretion. 4) Allowed: cancelling with applicable terms; a reduced ETF or MCT, where warranted. Not allowed: MF reductions; vouchers; credits or free months. Key principle: the priority is protecting the cleaner and enforcing the boundary, not retention. Resolutions should focus on a clean exit, consistent with policy.');
    doRows(s, 0.45, 4.45, 1.45, [
      ['No retention.', 'No incentives, no persuasion.'],
      ['Explain the terms.', 'FC and ETF: factual, neutral, not a negotiation.'],
      ['Refuses the ETF?', 'An ETF or MCT reduction may be considered (standard guidelines).'],
    ], 0.75, 0.1);
    headPanel(s, 5.1, 1.45, 2.15, 2.45, 'Allowed', ['Cancel with applicable terms', 'Reduced ETF or MCT, where warranted'], C.teal, C.tealSoft, 9);
    headPanel(s, 7.4, 1.45, 2.15, 2.45, 'Not allowed', ['MF reductions', 'Vouchers', 'Credits or free months'], RED, PINK, 9);
    await tip(s, 4.1, 'Key principle:', 'protect the cleaner and enforce the boundary. Aim for a clean exit, consistent with policy.', 'FiShield');
  }

  await knowledgeCheck(AB, [
    ['"This is f***ing frustrating, my place still isn\'t clean." Abusive?', 'No. Venting about the outcome, not aimed at the cleaner.'],
    ['"You\'re useless. You\'ll regret it." Abusive?', 'Yes. A personal insult and a threat aimed at the cleaner.'],
    ['Abuse confirmed. Can you offer a free month to keep them?', 'No. No retention, no credits, free months, vouchers or MF reductions.'],
    ['Abuse confirmed and C refuses to pay the ETF. Options?', 'An ETF or MCT reduction, using standard reduction guidelines.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(AB, [
    ['Is it abusive?', 'Intent, target, tone. Not just profanity.'],
    ['What actions are you going to take?', 'Terms, ETF, what\'s allowed and not.'],
    ['5-min writing challenge', 'Write a neutral, factual reply.'],
  ], 'TRAINER: share a live ticket where a customer has been abusive toward a cleaner (or one that\'s just frustrated). Give trainees a few minutes to review it on their own, then work through it together. 1) Is it abusive? Apply the intent, target and tone test. 2) What actions are you going to take? No retention; explain FC and ETF terms neutrally; an ETF or MCT reduction only if warranted; no MF reductions, vouchers, credits or free months. 3) 5-minute writing challenge: each trainee writes a neutral, factual reply. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 3. DUPLICATE ACCOUNTS =================
  const DA = 'Duplicate Accounts';
  await topic('Duplicate Accounts', 'Two accounts, two memberships, two MF charges every cycle.',
    ['Which account is the duplicate', 'Check Thumbtack first', 'Cancel the duplicate FC', 'Refund or free months'],
    'ill24.png',
    'Duplicate Accounts. Category: Retention and Membership Cancellation. A duplicate account is when a customer has inadvertently created two separate Homeaglow accounts, resulting in two active memberships and two Membership Fee (MF) charges each cycle. The customer may or may not realize they\'re paying twice: sometimes they contact us confused about an unrecognized charge, sometimes they\'ve never noticed at all.');

  {
    const s = content(DA, 'What is this?', 'The duplicate is the account the customer isn\'t using.',
      'The duplicate is the account the customer isn\'t using. That is the test: not which is older, and not which has more paid months. If the customer is actively using both, ask which they want to keep rather than deciding for them. If the duplicate doesn\'t surface on the obvious search, ask for identifying information (phone number, email address, and the card used) and search on those. "I can\'t find a second account" is not a conclusion until you have asked. This article assumes both accounts have a triggered FC. If one account has an active FC and the other holds an unused Legacy DHJ voucher with no FC ever triggered, the refund-or-free-months choice does not apply: refund the voucher and invalidate it. The customer paid a voucher price, not a membership fee, so there\'s no membership on that account for a free month to extend. Refundable only within 1 year of purchase. Why this matters: it\'s an account-management issue with a real financial impact, not a service failure or a customer mistake. Don\'t apologize as if something went wrong with their cleaning, and don\'t treat the customer as though they should have caught it. The job: stop the duplicate charges going forward, then make the customer whole in whichever form they prefer.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiUserX', title: 'Not being used', body: 'Not the older one, not the one with more paid months. Using both? Ask which to keep.' },
      { ico: 'FiSearch', title: 'Can\'t find it?', body: 'Ask for their phone, email and the card used, then search on those.' },
      { ico: 'FiTag', title: 'Unused DHJ voucher?', body: 'No FC on it: refund and invalidate the voucher (within 1 year).' },
    ]);
    await tip(s, 3.45, 'Tone:', 'not a service failure, not their mistake. Stop the double charge, then make them whole.', 'FiMessageCircle');
    T(s, [{ text: 'Good to know:  ', options: { bold: true, fontFace: HEAD } }, { text: 'one account per email, and one email can\'t hold both a C and a CP account. A C account can\'t be deactivated while an FC, RC or cleaning is still active.', options: { color: C.soft } }],
      { x: 0.45, y: 4.05, w: 9.1, h: 0.5, fontSize: 9.5 });
  }

  {
    const s = content(DA, 'Steps 0–1: Thumbtack, then cancel', 'Check the flag first. Then stop the next charge.',
      'IMPORTANT REMINDER: before you touch anything on the ticket, check whether the duplicate account is associated with Thumbtack. Thumbtack customers have their own handling process, so if it hits, this isn\'t your ticket to resolve. How to check: Legacy CRM > View > Flag > look for project_thumbtack. If the flag is present: reassign the ticket to the Thumbtack queue, and stop. If not: continue to Step 1. Step 1, confirm the duplicate and stop it from recurring: cancel the FC membership on the duplicate account immediately, regardless of which branch applies. This is the first action in every case, since it stops a third charge while you sort out the rest. Cancel any upcoming appointments or recurring cleaning plan on the duplicate account too, unless the customer has given you a reason to think they want to keep them. Leave an internal note on both accounts that includes the ticket link or C ID of the duplicate/main account, so the next agent can follow through.');
    await flowBox(s, 0.45, 1.45, 2.6, 0.62, 'Legacy CRM › View › Flag', 'end');
    flowLine(s, [[3.05, 1.76], [3.4, 1.76]], true);
    await flowBox(s, 3.4, 1.45, 2.6, 0.62, 'project_thumbtack?', 'decision');
    flowLine(s, [[6.0, 1.76], [6.35, 1.76]], true);
    T(s, 'Yes', { x: 6.0, y: 1.5, w: 0.35, h: 0.2, fontSize: 8.5, bold: true, color: C.soft, align: 'center' });
    await flowBox(s, 6.35, 1.45, 3.2, 0.62, 'Reassign to the Thumbtack queue. Stop.', 'penalty');
    T(s, 'No: continue', { x: 4.75, y: 2.12, w: 1.5, h: 0.2, fontSize: 8.5, bold: true, color: C.soft });
    flowLine(s, [[4.7, 2.07], [4.7, 2.4]], true);
    eyebrow(s, 'Step 1: on the duplicate account', 0.45, 2.45, 4, C.teal);
    doRows(s, 0.45, 9.1, 2.7, [
      ['Cancel the FC right away.', 'Every case. It stops a third charge while you sort out the rest.'],
      ['Cancel upcoming appointments and the RC,', 'unless C gave you a reason to keep them.'],
      ['Internal note on both accounts:', 'the ticket link or C ID of the other account.'],
    ], 0.58, 0.08);
  }

  {
    const s = content(DA, 'Steps 2–3: refund or free months', 'Let C choose. The bonus month comes only with free months.',
      'Step 2, determine the situation. Customer is currently paying two MFs, is unaware and not contesting, or is actively contesting the charges: same resolution in all three cases, present the refund/free-months choice. If they choose free months, a bonus month applies on top; if they choose a refund, no bonus applies. Customer insists on a refund and rejects free months: refund the duplicate within 120 days; beyond 120 days only if the duplicate was caused by a system/agent error. Step 3, present the choice: refund the duplicate charges, or apply them as free months on the main account, adjusting the MCT accordingly. If the customer contests the charges and chooses free months: add 1 additional month on top, as a goodwill gesture. If their total paid months reach 6, explain they\'ve effectively fulfilled their MCT. If they contest and choose a refund: no bonus month; the refund covers the duplicate charges only. This asymmetry is intentional: a free month costs Homeaglow far less than the refund, so the goodwill gesture is only offered on the cheaper option. Policy: refunds limited to charges within the past 120 days unless caused by a system or agent error; a maximum of 3 free months when merging duplicate charges into free months. Count the customer\'s history across both accounts wherever tiering matters (e.g. lockout customer category): 22 paid months and 8 jobs split across two accounts = Very High-Value.');
    headPanel(s, 0.45, 1.45, 4.45, 2.35, 'Free months on the main account', ['Duplicate charges become free months (max 3)', 'Adjust the MCT', 'Contesting? +1 bonus month', 'Paid months reach 6? MCT effectively fulfilled'], C.teal, C.tealSoft);
    headPanel(s, 5.1, 1.45, 4.45, 2.35, 'Refund', ['Duplicate charges within 120 days', 'Beyond 120 days: only for a system or agent error', 'No bonus month'], C.gold, C.goldSoft);
    await tip(s, 3.98, 'Why the asymmetry?', 'a free month costs far less than a refund, so the bonus comes with free months only.', 'FiGift');
    await tip(s, 4.5, 'Tiering:', 'count history across both accounts (22 paid months + 8 jobs split = Very High-Value).', 'FiLayers');
  }

  await knowledgeCheck(DA, [
    ['C has two accounts: one older, one they actually use. Which is the duplicate?', 'The one they aren\'t using, not the older one.'],
    ['What do you check before anything else?', 'The project_thumbtack flag (Legacy CRM › View › Flag). Present? Reassign and stop.'],
    ['What\'s your first action on the duplicate account?', 'Cancel its FC (plus upcoming jobs and RC), and note both accounts.'],
    ['C is contesting the charges and picks free months. Anything extra?', 'Yes, 1 bonus month on top. No bonus with a refund.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(DA, [
    ['Which account is the duplicate?', 'Thumbtack flag? Both triggered FCs?'],
    ['What actions are you going to take?', 'Cancel FC, notes, refund or free months.'],
    ['5-min writing challenge', 'Write your reply offering the choice.'],
  ], 'TRAINER: share a live ticket with a duplicate account. Give trainees a few minutes to review it on their own, then work through it together. 1) Which account is the duplicate (the one not being used)? Is there a project_thumbtack flag? Do both accounts have a triggered FC, or is one an unused DHJ voucher? 2) What actions are you going to take? Cancel the duplicate FC and any upcoming jobs/RC, leave internal notes on both accounts, then offer a refund (within 120 days) or free months (max 3, +1 bonus if contesting). 3) 5-minute writing challenge: each trainee writes the reply offering the choice. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 4. BULK CHARGES =================
  const BC = 'Bulk Charges';
  await topic('Bulk Charges', 'Several overdue MFs charged in one go, on the same day.',
    ['What a bulk charge is (and isn\'t)', 'The three branches', 'Branch A: head it off with free months', 'Branch C: save mode or exit mode'],
    'ill27.png',
    'Bulk MF Charges. Category: Retention and Membership Cancellation. Bulk charges happen when a customer\'s ForeverClean (FC) monthly fee has been overdue for multiple billing cycles, and a payment attempt finally goes through. Instead of billing one overdue month at a time, the system charges all the past-due months in a single transaction.');

  {
    const s = content(BC, 'What is this?', 'Each charge is a different missed month, not a duplicate.',
      'Bulk charges happen when a customer\'s FC monthly fee has been overdue for multiple billing cycles, and a payment attempt finally goes through. The system charges all the past-due months in a single transaction, so the customer sees several MF charges posted on the same day: each one for a different month they missed, not a duplicate of the same charge. This applies to both DHJ and FCF memberships: the overdue-then-bulk-charge mechanic and the true-up/retention logic are the same. Why it\'s happening more now: as of September 12, 2024, the system stopped auto-issuing free months to true up overdue accounts. Before that, if a customer got successfully charged after falling behind, the system quietly forgave the gap with free months. Now it bills the whole backlog at once, so the proactive true-up (Branch A) falls to agents. Why this matters: a wall of same-day charges reads as a billing error to the customer, even when every dollar was owed. Your job isn\'t to apologize for a correct charge: it\'s to figure out whether the customer got real value for those months (were they using the service, or did they not even know they were being billed) and let that drive how much you fix. Prevention: if you can see a bulk charge coming, it\'s cheaper for everyone to head it off with free months than to untangle a large charge and an upset customer afterwards.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiLayers', title: 'Several months, one day', body: 'Each charge is a different missed month. DHJ and FCF alike.' },
      { ico: 'FiCalendar', title: 'Since Sep 12, 2024', body: 'No more automatic free months. The true-up is now on agents.' },
      { ico: 'FiSearch', title: 'Did they get value?', body: 'Using the service, or didn\'t know they were billed? That drives the fix.' },
    ]);
    eyebrow(s, 'Three reasons it reaches you', 0.45, 3.4, 5);
    const w = (9.1 - 0.3) / 3;
    [['A', 'FC overdue, not why C wrote in'], ['B', 'FC cancelled, wants a refund'], ['C', 'FC active, cancelling over it']].forEach(([l, t], i) => {
      const x = 0.45 + i * (w + 0.15);
      box(s, x, 3.65, w, 0.75, C.white, C.border);
      s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 3.85, w: 0.36, h: 0.36, fill: { color: C.teal }, line: { color: C.teal } });
      T(s, l, { x: x + 0.15, y: 3.85, w: 0.36, h: 0.36, fontFace: HEAD, bold: true, fontSize: 11, color: C.white, align: 'center', valign: 'middle' });
      T(s, t, { x: x + 0.62, y: 3.65, w: w - 0.75, h: 0.75, fontFace: HEAD, bold: true, fontSize: 10, valign: 'middle' });
    });
  }

  {
    const s = content(BC, 'Branch A: head it off', 'You spotted an overdue FC on an unrelated ticket.',
      'Branch A: FC is overdue, but that\'s not why the customer\'s writing in. Pure housekeeping: you spotted an overdue FC while working an unrelated ticket, and you want to head off a bulk charge. Overdue by 1 MF, next charge due in 7 days or less: issue 1 free month so the customer doesn\'t get hit with two MF charges close together. Overdue by 1 MF, next charge more than 7 days out: no action needed. Overdue by 2+ MFs: issue free months/weeks to bring the paid-through date current. "Current" means current as of the billing anchor, not as of today. MF billing runs on a fixed day of the month, so true-up months move the paid-through date to the last missed anchor; they don\'t stretch to cover the gap between that anchor and today. Worked example: today is Aug 24. MF bills on the 10th. The customer is overdue by 2 months (the July 10 and Aug 10 charges failed). Paid-through is currently June 10. Add 2 free months: paid-through moves to Aug 10, next MF due Sept 10.');
    table(s, ['Situation', 'Action'], [
      ['Overdue by 1 MF, next charge in 7 days or less', 'Issue 1 free month'],
      ['Overdue by 1 MF, next charge more than 7 days out', 'No action needed'],
      ['Overdue by 2+ MFs', 'Free months/weeks to bring paid-through current'],
    ], { y: 1.45, colW: [4.8, 4.3], fontSize: 9.5, rowH: 0.4 });
    eyebrow(s, 'Worked example: today is Aug 24, MF bills on the 10th', 0.45, 3.2, 8, C.soft);
    const w = 2.1, gap = 0.233;
    const flow = [['Jul 10 + Aug 10 failed', 'end'], ['Paid-through: Jun 10', 'decision'], ['Add 2 free months', 'action'], ['Paid-through Aug 10. Next MF Sep 10', 'action']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 3.45, w, 0.62, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 3.76], [x + w + gap, 3.76]], true);
    }
    await tip(s, 4.3, '"Current" =', 'the last missed billing anchor, not today.', 'FiCalendar');
  }

  {
    const s = content(BC, 'Branches B & C', 'Already cancelled, or cancelling because of the bulk charge.',
      'Branch B: ForeverClean is already cancelled, and the customer wants a refund for the bulk charge. This is just the standard refund policy, nothing bulk-charge-specific: unused MF charges beyond MCT, made within the past 120 days, are refundable. See MF & ETF Refund Requests for the full nuance (system-error exceptions, high-dispute-risk overrides, etc.). Branch C: FC is still active, and the bulk charge is why the customer wants to cancel. First, check whether the account is still overdue even after the bulk charge landed, then check MCT status, because that determines save mode or exit mode. Beyond MCT: the ETF doesn\'t apply; the customer is free to leave. The only question is whether it\'s worth trying to keep them and what they\'re owed either way. The free-month offer isn\'t a discount: it gives back value for months they already paid for and can\'t use anymore. The $20 credit is a separate, bulk-charge-specific incentive to get them back in for a cleaning. Within MCT: the ETF applies, and the retention attempt happens first. If it doesn\'t land, the exit path branches on the membership model: FCF customers who haven\'t had a first cleaning have no ETF at all (FCF\'s ETF only kicks in after the first completed cleaning). For everyone else, apply whichever option costs the customer less: crediting the unused bulk MFs against the ETF, or reducing the MCT.');
    headPanel(s, 0.45, 1.45, 2.9, 3.0, 'B: FC already cancelled', ['Standard refund policy', 'Unused MFs beyond MCT, within 120 days', 'See MF & ETF Refund Requests'], C.ink, C.white, 9);
    eyebrow(s, 'C: FC active. Still overdue? Then check the MCT', 3.5, 1.42, 6, C.teal);
    headPanel(s, 3.5, 1.68, 2.95, 2.77, 'Beyond MCT: no ETF', ['Free to leave', 'Free months give back value for paid, unusable months', '+ $20 credit to get them back in'], C.teal, C.tealSoft, 9);
    headPanel(s, 6.6, 1.68, 2.95, 2.77, 'Within MCT: ETF applies', ['Retention first', 'FCF, no first cleaning: no ETF', 'Otherwise the cheaper of: credit bulk MFs to the ETF, or reduce the MCT'], C.gold, C.goldSoft, 9);
  }

  {
    const s = content(BC, 'Within MCT: ETF credit or MCT reduction?', 'Near the end of the term, run the comparison.',
      'In practice these two levers usually aren\'t interchangeable: MCT reduction is capped at 2 months, and only usable when that reduction is cheaper than a reduced ETF. So the comparison only really matters when the customer is close to the end of their term, where it\'s usually more cost-effective to keep the membership at a reduced MCT than to pay a reduced ETF. Otherwise, the ETF reduction is the only lever that applies. Worked example: the customer has 1 month left on MCT, owes a $150 ETF, and has $100 in unused bulk MFs to credit. ETF credit path: $100 credit, $50 ETF still owed. MCT reduction path: reduce the last month, MCT complete, ETF waived entirely, $0 owed. MCT reduction wins here. Run this comparison whenever the customer is close to the end of their term; don\'t default to the ETF credit just because it\'s the more familiar move.');
    const w = (9.1 - 0.3) / 3;
    stat(s, 0.45, 1.45, w, 1.1, '1', 'month left on MCT');
    stat(s, 0.45 + w + 0.15, 1.45, w, 1.1, '$150', 'ETF owed');
    stat(s, 0.45 + 2 * (w + 0.15), 1.45, w, 1.1, '$100', 'unused bulk MFs to credit');
    headPanel(s, 0.45, 2.75, 4.45, 1.45, 'ETF credit path', ['$150 ETF − $100 credit', '$50 still owed'], C.ink, C.white);
    headPanel(s, 5.1, 2.75, 4.45, 1.45, 'MCT reduction path: wins', ['Reduce the last month: MCT complete', 'ETF waived: $0 owed'], C.teal, C.tealSoft);
    await tip(s, 4.4, 'MCT reduction:', 'capped at 2 months, and only when it\'s cheaper than a reduced ETF.', 'FiScissors');
  }

  await knowledgeCheck(BC, [
    ['C sees four MF charges on one day. Duplicates?', 'No. Each is a different missed month, billed at once.'],
    ['Unrelated ticket: FC overdue by 1 MF, next charge in 5 days. Action?', 'Issue 1 free month.'],
    ['Overdue July 10 and Aug 10, today is Aug 24. Free months? New next MF?', '2 free months. Paid-through Aug 10, next MF Sept 10.'],
    ['1 month left on MCT, $150 ETF, $100 unused bulk MFs. Best exit?', 'Reduce the MCT: ETF waived, $0 owed (vs $50).'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(BC, [
    ['Which branch is it?', 'A, B or C? Still overdue? Within MCT?'],
    ['What actions are you going to take?', 'Free months, refund, $20 credit, ETF vs MCT.'],
    ['5-min writing challenge', 'Explain the charges and your offer.'],
  ], 'TRAINER: share a live ticket with a bulk MF charge (or an overdue FC). Give trainees a few minutes to review it on their own, then work through it together. 1) Which branch is it: A (overdue, not why C wrote in), B (FC already cancelled, wants a refund) or C (FC active, wants to cancel because of the bulk charge)? Is the account still overdue? Within or beyond MCT? 2) What actions are you going to take? Branch A: true up to the last missed anchor. Branch B: standard refund policy. Branch C: beyond MCT, free months for the paid unusable months + $20 credit; within MCT, retention first, then the cheaper of an ETF credit or an MCT reduction (FCF with no first cleaning: no ETF). 3) 5-minute writing challenge: each trainee explains the charges and the offer. Read a few aloud and compare. Then open the floor for questions.');

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Thank you!', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Close the module. Recap: MF & ETF Refund Requests (root cause drives the decision; eligible means free months first if FC is active and C is retainable; the MCT test and the 120-day/disputable rule; billing errors and agent errors are refunded directly; high-risk and hardship cases). Customer Abusive Behavior (intent, target, tone; no retention; neutral terms; only an ETF or MCT reduction where warranted). Duplicate Accounts (the duplicate is the one not being used; Thumbtack flag first; cancel the duplicate FC; refund within 120 days or free months, max 3, +1 bonus if contesting). Bulk Charges (each charge is a different missed month; Branch A true-up to the last anchor; Branch B standard refund policy; Branch C save or exit mode by MCT, compare ETF credit vs MCT reduction).');
  }

  if (topicSlides.length * 3 !== agendaLinks.length) throw new Error(`agenda has ${agendaLinks.length / 3} rows but the deck has ${topicSlides.length} topics`);
  agendaLinks.forEach((l, i) => { agendaSlide._rels.find(r => r.rId === l._rId).Target = String(topicSlides[Math.floor(i / 3)]); });
  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
