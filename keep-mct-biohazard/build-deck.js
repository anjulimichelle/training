// ---------- Deck 8: Keep MCT & Biohazard Reports ----------
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

// Customer message bubble.
function bubble(s, x, y, w, h, text, label = 'CUSTOMER') {
  box(s, x, y, w, h, C.bg, C.border);
  T(s, [{ text: label, options: { bold: true, fontFace: HEAD, fontSize: 7.5, color: C.soft, breakLine: true } },
        { text: '“' + text + '”', options: { fontSize: 9.5, color: C.ink } }], { x: x + 0.16, y, w: w - 0.32, h, valign: 'middle', paraSpaceAfter: 2 });
}

(async () => {
  // ---------- Cover ----------
  {
    const s = pres.addSlide({ masterName: 'COVER' });
    eyebrow(s, 'Homeaglow  ·  New Hire Care Training', 0.62, 0.6);
    T(s, 'Keep MCT &\nBiohazard\nReports', { x: 0.62, y: 0.95, w: 4.9, h: 1.75, fontFace: HEAD, bold: true, fontSize: 30 });
    T(s, 'Customers who plan to finish their term, and homes reported as unsafe for cleaners.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers two Knowledge Library articles from the Retention and Membership Cancellation category: Customer Agreed to Keep MCT (a customer who plans to complete their Minimum Commitment Term and then cancel, to avoid the ETF), and CP Reported C\'s Home as Biohazardous or Unsanitary (the customer-side actions after a CP\'s report has been validated). The thread through both: act within the terms the customer agreed to, set the right Timed Reminders, and give the customer a fair chance before closing things out.');
  }

  // ---------- Agenda ----------
  const agendaLinks = [], topicSlides = [];
  let agendaSlide;
  const addTopic = topic;
  topic = async (...args) => { const t = await addTopic(...args); topicSlides.push(t._slideNum); return t; };
  {
    const s = agendaSlide = content('Overview', 'Agenda', 'Two topics. Click a topic to jump to it.',
      'Topic 1 covers customers who agree to complete their MCT and cancel afterwards: what to say (and not say) about the ETF, why not to waive it, and what to do when the Timed Reminder triggers. Topic 2 covers the customer side of a validated CP biohazard or hoarder report: the 1st Response warning, the 2nd Response evaluation, and refunds by membership status. In slideshow mode, click any topic to jump straight to it.');
    const rows = [
      ['Customer Agreed to Keep MCT', 'They\'ll finish the term, then cancel. No ETF talk, no ETF waiver.'],
      ['CP Reported C\'s Home as Biohazardous or Unsanitary', 'Restrict the account, give them time, then evaluate.'],
    ];
    const cw = 9.1, h = 1.0, gy = 0.15;
    for (let i = 0; i < rows.length; i++) {
      const [title, line] = rows[i], x = 0.45, y = 1.45 + i * (h + gy);
      const link = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(link);
      s.addText([
        { text: title, options: { color: C.ink, bold: true, fontFace: HEAD, fontSize: 14, breakLine: true, hyperlink: link } },
        { text: line, options: { color: C.soft, fontSize: 10.5, hyperlink: link } },
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

  // ================= 1. CUSTOMER AGREED TO KEEP MCT =================
  const MCT = 'Keep MCT';
  await topic('Customer Agreed to Keep MCT', 'They plan to finish the term, then cancel, to avoid the ETF.',
    ['What this pattern looks like', 'The two traps to avoid', 'Said it vs. showed it (no bookings)', 'What to do when the TR triggers'],
    'ill27.png',
    'Customer Agreed to Keep MCT (Customer Agreed to Fulfill MCT). Category: Retention and Membership Cancellation. This covers a customer who intends to complete their Minimum Commitment Term and then cancel, usually saying so explicitly, sometimes just by not booking any cleanings during that period, in order to avoid paying the Early Termination Fee.');

  {
    const s = content(MCT, 'What is this?', 'C plans to complete their MCT, then cancel, so they don\'t pay the ETF.',
      'What is this: a customer who intends to complete their Minimum Commitment Term and then cancel, usually explicitly stating this, sometimes just by not booking any cleanings during that period, in order to avoid paying the Early Termination Fee. Why this matters: the instinct is tricky in both directions. Trap 1: some agents over-warn the customer about the ETF when they\'ve already made peace with paying off the term, which just sounds naggy. Trap 2: others think "they\'re not using the service, we\'ve basically already been paid enough" and waive the ETF as the path of least resistance. That\'s the wrong call. '
      + 'When the customer hasn\'t booked any jobs on purpose, it requires real judgment, because it\'s easy to talk yourself into the wrong answer. Example: the customer has 1 invoiced job and 4 paid months, the ETF is $105, and they ask when they\'ll hit their MCT because they plan to cancel then. The reasoning trap: "the unused monthly fees basically already cover the ETF" or "we\'ve already earned enough from them". Neither is a valid reason to waive the ETF. The customer isn\'t disputing anything, and nothing unfair has happened to them: they\'ve simply accepted the financial terms and are planning around them. A waiver exists to fix something unfair or relieve a real burden. There\'s neither here. Don\'t waive the ETF. Why this matters for the business: MCT and ETF exist to keep terms fair and consistent for everyone. Waiving it here, just because a customer chose not to book and the math looks favorable, undermines the policy and sets an unrealistic precedent for every other customer who did book and is paying the ETF as agreed.');
    const sig = [
      ['FiMessageCircle', 'They say it', 'Tells us they\'ll cancel once the MCT is done.'],
      ['FiCalendar', 'They show it', 'Doesn\'t book on purpose. Ex: 1 invoiced job, 4 paid MFs, $105 ETF, asks when they\'ll hit MCT.'],
    ];
    for (let i = 0; i < 2; i++) {
      const x = 0.45 + i * 4.65;
      box(s, x, 1.42, 4.45, 0.95, C.white, C.border);
      await iconDot(s, x + 0.18, 1.42 + 0.29, sig[i][0], 0.38);
      T(s, [{ text: sig[i][1], options: { bold: true, fontFace: HEAD, fontSize: 12, breakLine: true } }, { text: sig[i][2], options: { fontSize: 9.5, color: C.soft } }],
        { x: x + 0.72, y: 1.42, w: 3.6, h: 0.95, valign: 'middle', paraSpaceAfter: 2 });
    }
    eyebrow(s, 'Two traps to avoid', 0.45, 2.52, 4, RED);
    headPanel(s, 0.45, 2.75, 4.45, 1.55, 'Trap 1: Over-warning about the ETF', ['They already accepted paying off the term', 'Bringing it up again just sounds naggy'], RED, PINK);
    headPanel(s, 5.1, 2.75, 4.45, 1.55, 'Trap 2: Waiving the ETF', ['"The unused MFs already cover the ETF"', 'Nothing unfair happened, and nothing is disputed', 'It sets a precedent for customers who pay as agreed'], RED, PINK, 9);
    await warn(s, 4.45, 'Don\'t waive the ETF:', 'they\'re working within terms they already agreed to.');
  }

  {
    const s = content(MCT, 'They say it: "I\'ll cancel after my MCT"', 'Confirm their plan, set a TR, and keep the probe friendly.',
      'When the customer explicitly mentions paying off the MCT. Example: "I would like to cancel my membership. The agent who signed me up said she would allow me to cancel in 4 months without penalty. I would like to cancel in the 4th month without penalty as promised." Do: 1) Don\'t remind or warn about the ETF: since they\'ve already agreed to complete the term, bringing it up again can feel unnecessary or even aggravating. 2) Acknowledge and confirm their intent to fulfill the membership term. 3) Create a Timed Reminder for the day after the MF that fulfills the MCT is charged, so you\'re prompted to follow up. 4) Advise them plainly: "We\'ve noted your request and will follow up once your minimum commitment term is fulfilled." 5) Still probe to understand what\'s driving the planned cancellation. Since the customer isn\'t actually cancelling yet, don\'t make the probe sound like a cancellation interview or an attempt to talk them out of it. Keep it conversational and caring. Sample probe: "Got it! We can definitely note that for you. Since you\'re still early in your membership and just planning ahead, can I ask what\'s making you feel you may not want to continue after the six months? And how was your cleaning with <<cleaner name>>? I just want to make sure everything went well and see if there\'s anything we can help with."');
    doRows(s, 0.45, 4.45, 1.45, [
      ['Don\'t mention the ETF.', 'They already agreed to the term.'],
      ['Acknowledge and confirm', 'their plan to finish the term.'],
      ['Set a TR:', 'the day after the MF that fulfills the MCT is charged.'],
      ['Advise:', '"We\'ve noted your request and will follow up once your MCT is fulfilled."'],
      ['Still probe, kindly.', 'Not a cancellation interview.'],
    ], 0.6, 0.08);
    bubble(s, 5.1, 1.45, 4.45, 1.2, 'I would like to cancel my membership. The agent who signed me up said she would allow me to cancel in 4 months without penalty. I would like to cancel in the 4th month without penalty as promised.');
    quote(s, 5.1, 2.8, 4.45, 2.05, 'Got it! We can definitely note that for you. Since you\'re still early in your membership and just planning ahead, can I ask what\'s making you feel you may not want to continue after the six months? And how was your cleaning with <<cleaner name>>?', 'SAMPLE PROBE');
  }

  {
    const s = content(MCT, 'One more retention attempt? (optional)', 'Worth a try, as long as nothing gets waived.',
      'Optional, if you want to try for one more retention attempt anyway. Suggested framing: "I\'ll make a note in the system to cancel your membership once [MCT is fulfilled]... I know I\'m pushing my luck here, but I\'d really love to give you a better experience. I can book you with a new cleaner and even apply a $20 credit toward that job." Alternative, if the customer is close to finishing the MCT and seems frustrated: consider offering to charge the remaining month(s) now and cancel immediately. This can de-escalate without actually waiving anything, since they\'re still paying what they owe, just faster. For reference, see the Notion article PG - Creating Manual Charges.');
    quote(s, 0.45, 1.45, 4.45, 2.0, 'I\'ll make a note in the system to cancel your membership once your MCT is fulfilled... I know I\'m pushing my luck here, but I\'d really love to give you a better experience. I can book you with a new cleaner and even apply a $20 credit toward that job.', 'NEW CLEANER + $20 CREDIT');
    headPanel(s, 5.1, 1.45, 4.45, 2.0, 'Close to MCT and frustrated?', ['Offer to charge the remaining month(s) now', 'Then cancel right away', 'De-escalates without waiving anything: they still pay what they owe, just faster'], C.gold, C.goldSoft);
    await tip(s, 3.65, 'Manual charges:', 'see the Notion article PG - Creating Manual Charges.', 'FiBookOpen');
  }

  {
    const s = content(MCT, 'When the TR triggers', 'First check: did C book again after you set the reminder?',
      'When the Timed Reminder for MCT fulfillment triggers, check first: did the customer book again between when you set the reminder and when it triggered? If they did not book again: proactively cancel the membership and send a message confirming it\'s been closed. Rationale: this typically confirms the customer genuinely intended to leave once the term was up, and was just waiting it out to avoid the ETF. If they did book again: reach out to confirm whether they still want to cancel, then create another TR at least 3 days before the next MF is charged, so the customer isn\'t charged another MF unless they agreed to continue with the membership. Rationale: continued use during the MCT period is a real signal of retainability, so don\'t assume the original cancellation intent still holds without checking. Reminder: since the ETF does not apply, clearly confirm no additional MF will be charged unless they explicitly agree to continue.');
    await flowBox(s, 3.5, 1.42, 3.0, 0.5, 'Did C book again?', 'decision');
    flowSplit(s, 5.0, 1.92, 2.12, 2.675, 7.325, 2.35, 'No', 'Yes');
    headPanel(s, 0.45, 2.35, 4.45, 1.75, 'No: they waited it out', ['Proactively cancel the membership', 'Send a message confirming it\'s closed', 'They meant to leave once the term was up'], C.teal, C.tealSoft);
    headPanel(s, 5.1, 2.35, 4.45, 1.75, 'Yes: they used it again', ['Reach out: do they still want to cancel?', 'New TR at least 3 days before the next MF', 'Continued use is a real retainability signal'], C.gold, C.goldSoft);
    await tip(s, 4.3, 'Reminder:', 'no ETF applies. Confirm no more MFs will be charged unless they agree to continue.', 'FiBell');
  }

  await knowledgeCheck(MCT, [
    ['C: "I\'ll cancel once my 6 months are up." Do you remind them about the ETF?', 'No. They already agreed to finish the term. Confirm, set a TR, probe kindly.'],
    ['When should that Timed Reminder be set for?', 'The day after the MF that fulfills the MCT is charged.'],
    ['No bookings, 4 paid MFs, ETF $105. Waive it, since they\'ve paid enough?', 'No. Nothing unfair happened; they\'re working within terms they agreed to.'],
    ['The TR triggers and C booked again during the term. What now?', 'Reach out to confirm they still want to cancel; new TR at least 3 days before the next MF.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(MCT, [
    ['Did C say it, or show it?', 'Explicit plan, or no bookings on purpose?'],
    ['What actions are you going to take?', 'TR timing, ETF, one more retention attempt?'],
    ['5-min writing challenge', 'Write your reply to the customer.'],
  ], 'TRAINER: share a live ticket where a customer plans to finish their MCT and then cancel (or asks when they\'ll hit their MCT). Give trainees a few minutes to review it on their own, then work through it together. 1) Did the customer say it, or show it by not booking? 2) What actions are you going to take? No ETF warning, confirm their intent, TR for the day after the MCT-fulfilling MF, no ETF waiver, optional retention attempt (new cleaner + $20 credit, or charge the remaining months now and cancel). 3) 5-minute writing challenge: each trainee writes the reply, including a friendly probe. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 2. BIOHAZARD =================
  const BIO = 'Biohazard Report';
  await topic('Biohazard & Unsanitary Home Reports', 'CP reported C\'s home as biohazardous or unsanitary. What you do on the customer side once it\'s validated.',
    ['What counts, and what doesn\'t', 'What happened before it reached you', '1st Response: the warning', '2nd Response: evaluate and resolve'],
    'ill7.png',
    'CP Reported C\'s Home as Biohazardous or Unsanitary (C-Side Action). Category: Retention and Membership Cancellation. This covers what happens on the customer side once a CP has reported that a customer\'s home is biohazardous or extremely unsanitary: a report serious enough to potentially get the customer\'s account suspended from booking. It picks up after the CP-side investigation (handled by the Bilingual team) has already validated the CP\'s evidence.');

  {
    const s = content(BIO, 'What is this?', 'A validated CP report that can suspend the customer from booking.',
      'This page picks up after the CP-side investigation (handled by the Bilingual team) has already validated the CP\'s evidence. If you\'re on a customer-facing ticket and see an action item referencing a biohazard/unsanitary report, these instructions only apply if a CP-side agent already created that action item based on validated CP evidence. Don\'t act on a raw customer complaint about this without that upstream step having happened. Why this matters: it\'s a safety issue (a CP shouldn\'t have to work in a hazardous environment), but it\'s the customer\'s account being restricted, not the CP\'s, and the customer hasn\'t necessarily done anything malicious. So the tone and offers need to reflect "let\'s help you fix this," not "you\'re being punished." Getting the sequencing wrong (e.g. cancelling FC and waiving the ETF before giving the customer a chance to resolve it) skips a real opportunity to retain a customer who just needs time to address a hoarding or pest situation. The process is shared across Legacy DHJ and FCF: a biohazard is a biohazard regardless of which membership model is billing the customer. The only place they diverge is refund mechanics, because FCF customers are charged the MF even with zero completed jobs (the membership starts at sign-up), which a DHJ customer can never be (DHJ doesn\'t charge any MF until the first completed cleaning).');
    await cards(s, 1.45, 1.55, [
      { ico: 'FiShield', title: 'Safety first', body: 'A CP shouldn\'t have to work in a hazardous home.' },
      { ico: 'FiHeart', title: 'Help, don\'t punish', body: '"Let\'s help you fix this." C may just need time to deal with hoarding or pests.' },
      { ico: 'FiLayers', title: 'DHJ and FCF alike', body: 'One process for both. Only the refund mechanics differ.' },
    ]);
    await warn(s, 3.2, 'Validated reports only:', 'a CP-side agent must have created the action item. Never act on a raw complaint.');
    await tip(s, 3.77, 'Order matters:', 'don\'t cancel FC or waive the ETF before C has had a chance to fix it.', 'FiList');
  }

  {
    const s = content(BIO, 'What counts?', 'It has to be excessive and a direct risk to the CP.',
      'Quick reference. Biohazard conditions: presence of human/animal waste, bodily fluids, mold infestations, hazardous materials, vermin/insect infestations. These must be excessive and pose a direct risk to the CP. Does NOT qualify: normal shower mold; bugs that don\'t interfere with the CP finishing the job. Hoarder conditions: excessive clutter that blocks access to critical areas (kitchen, bathroom, exits) or creates fire, structural or unsanitary hazards. Does NOT qualify: general clutter that doesn\'t impact the CP\'s ability to clean.');
    headPanel(s, 0.45, 1.45, 4.45, 1.9, 'Biohazard', ['Human or animal waste, bodily fluids', 'Mold infestations', 'Hazardous materials', 'Vermin or insect infestations'], RED, PINK);
    headPanel(s, 5.1, 1.45, 4.45, 1.9, 'Hoarder', ['Clutter that blocks the kitchen, bathroom or exits', 'Clutter that creates fire, structural or unsanitary hazards'], C.gold, C.goldSoft);
    chip(s, 0.45, 3.5, 4.45, '✕  Doesn\'t count: normal shower mold, bugs that don\'t stop the CP', C.white, C.soft, 0.42);
    chip(s, 5.1, 3.5, 4.45, '✕  Doesn\'t count: general clutter that doesn\'t affect cleaning', C.white, C.soft, 0.42);
    await tip(s, 4.15, 'Rule of thumb:', 'excessive, and a direct risk to the CP.', 'FiAlertCircle');
  }

  {
    const s = content(BIO, 'Before it reaches you', 'A CP-side Bilingual agent has already done this. Background only.',
      'CP-side context (background only; this part already happened before the ticket reaches you). A CP-side Bilingual agent has already: escalated the CP\'s report; requested photo/video evidence (with a 72-hour Timed Reminder if it wasn\'t initially provided); verified the evidence isn\'t fake (reverse image search); and, if valid, added a known_customer_safety_issue flag and created a Tier 3 action item. The screenshot is a sample T3 action item (cp_biohazard_report): CP ID, job ID and date/time, link to the CP\'s message with the evidence, and the outcome of the investigation. Everything from here on (1st Response, 2nd Response, scenario notes) is what you actually do on the customer side.');
    const steps4 = [['Escalated', 'the CP\'s report'], ['Requested evidence', 'photos/videos (72-hr TR if missing)'], ['Verified it', 'reverse image search'], ['Valid?', 'safety flag + T3 action item']];
    const sw = (9.1 - 0.45) / 4;
    for (let i = 0; i < 4; i++) {
      const x = 0.45 + i * (sw + 0.15);
      box(s, x, 1.45, sw, 0.8, C.white, C.border);
      badge(s, x + 0.12, 1.45 + 0.25, i + 1);
      T(s, [{ text: steps4[i][0], options: { bold: true, fontFace: HEAD, breakLine: true } }, { text: steps4[i][1], options: { color: C.soft, fontSize: 8.5 } }],
        { x: x + 0.52, y: 1.45, w: sw - 0.6, h: 0.8, fontSize: 9.5, valign: 'middle' });
      if (i < 3) flowLine(s, [[x + sw, 1.85], [x + sw + 0.15, 1.85]], true);
    }
    eyebrow(s, 'Sample T3 action item (what you\'ll see on the ticket)', 0.45, 2.42, 6, C.soft);
    await shot(s, 'bio_ticket.png', 0.45, 2.65, 6.5, 2.2);
    headPanel(s, 7.1, 2.65, 2.45, 2.2, 'cp_biohazard_report', ['CP ID', 'Job ID + date/time', 'Link to the CP\'s evidence', 'Outcome of the investigation'], C.teal, C.tealSoft, 9);
  }

  {
    const s = content(BIO, '1st Response: the warning', 'Restrict the account, then respond based on membership status.',
      '1st Response, warning. 1) Cancel the customer\'s RC and all upcoming jobs, using Escalated Safety Issue as the cancellation reason code. 2) Add the known_customer_fraud flag to prevent further bookings, with the comment format: cp_biohazard_report | C ID | CP ID | J ID. 3) Create an internal note explaining why the account is being restricted: biohazard_1st_response — <CP ID> reported that C\'s home poses biohazard risks — <link to CP-side action item>. 4) Determine the customer\'s membership status and respond accordingly (next slide). Note the two flags: known_customer_safety_issue was added on the CP side and stays as a historical flag; known_customer_fraud is the one you add, and it\'s what blocks booking.');
    await checkRows(s, 1.45, [
      ['Cancel the RC + all upcoming jobs', 'Cancellation reason code: Escalated Safety Issue'],
      ['Add the known_customer_fraud flag', 'Blocks bookings. Comment: cp_biohazard_report | C ID | CP ID | J ID'],
      ['Create an internal note', 'biohazard_1st_response — <CP ID> reported that C\'s home poses biohazard risks — <link to CP-side action item>'],
      ['Check the membership status', 'Then respond accordingly (next slide).'],
    ], 0.6, 0.08);
    await tip(s, 4.3, 'Two flags:', 'known_customer_safety_issue (CP side, stays) and known_customer_fraud (yours, blocks booking).', 'FiFlag');
  }

  {
    const s = content(BIO, '1st Response by membership status', 'Give them time to fix it.',
      'Active FC (applies to both DHJ and FCF customers whose FC has started): issue 1 free month to give them time to address the situation, and set a Timed Reminder to trigger the 2nd Response check. Action on: the day before the free month expires. Who should act: any CS. Action: biohazard_followup — CP <CP_ID> reported C <C_ID> for unsanitary/biohazardous home conditions. Follow this article > 2nd Response. Unused DHJ voucher (no FC triggered yet): do not issue any DHJ retention offers. Related macros: "Bio Hazard" (covers Active FC, meaning both DHJ and FCF customers whose FC has started) and "Biohazard – Customer Insists on Booking Despite Safety Issue".');
    headPanel(s, 0.45, 1.45, 5.4, 1.15, 'Active FC (DHJ and FCF)', ['Issue 1 free month to give them time', 'Set a Timed Reminder for the 2nd Response'], C.teal, C.tealSoft);
    box(s, 0.45, 2.72, 5.4, 1.65, C.white, C.border);
    eyebrow(s, 'Timed Reminder', 0.65, 2.85, 3);
    T(s, [
      { text: 'Action on:  ', options: { bold: true, fontFace: HEAD } }, { text: 'the day before the free month expires', options: { breakLine: true } },
      { text: 'Who:  ', options: { bold: true, fontFace: HEAD } }, { text: 'any CS', options: { breakLine: true } },
      { text: 'Action:  ', options: { bold: true, fontFace: HEAD } }, { text: 'biohazard_followup — CP <CP_ID> reported C <C_ID> for unsanitary/biohazardous home conditions. Follow this article > 2nd Response.' },
    ], { x: 0.65, y: 3.1, w: 5.0, h: 1.2, fontSize: 9.5, paraSpaceAfter: 4 });
    headPanel(s, 6.0, 1.45, 3.55, 1.15, 'Unused DHJ voucher', ['No FC triggered yet', 'Don\'t issue any DHJ retention offers'], C.gold, C.goldSoft);
    box(s, 6.0, 2.72, 3.55, 1.65, C.goldSoft, C.gold);
    s.addImage({ data: await icon('FiMessageSquare', C.gold), x: 6.18, y: 2.88, w: 0.18, h: 0.18 });
    T(s, [
      { text: 'Macros', options: { bold: true, fontFace: HEAD, color: C.gold, breakLine: true } },
      { text: 'Bio Hazard', options: { bold: true, fontFace: HEAD, breakLine: true } },
      { text: 'Active FC (DHJ or FCF)', options: { color: C.soft, breakLine: true } },
      { text: 'Biohazard – Customer Insists on Booking Despite Safety Issue', options: { bold: true, fontFace: HEAD } },
    ], { x: 6.45, y: 2.82, w: 2.95, h: 1.45, fontSize: 9.5, paraSpaceAfter: 3 });
  }

  {
    const s = content(BIO, '2nd Response: they say it\'s fixed', 'Take their word for it the first time.',
      '2nd Response, evaluation. Customer says the issue is resolved: invalidate known_customer_fraud (but keep known_customer_safety_issue: it stays as a historical flag). Log an internal note, then help them book a new appointment. Important: we take the customer\'s word for it the first time. However, if the CP of the new appointment reports another biohazard concern after the customer told us it had been taken care of, that\'s when we ask the customer for proof before allowing them to book again. Review whether the proof reasonably addresses what the CP actually reported. Sufficient: invalidate known_customer_fraud (keep known_customer_safety_issue), log an internal note, then help them book a new appointment. Not sufficient (e.g. the CP reported hoarding, but the photos only show one room cleaned): don\'t invalidate the flag. If FC has started, issue 1 more free month. Advise the customer the evidence doesn\'t meet safety standards yet, and ask if they need more time.');
    const w = 2.1, gap = 0.233;
    const flow = [['C says it\'s resolved', 'end'], ['Invalidate known_customer_fraud', 'action'], ['Log an internal note', 'action'], ['Help them book', 'action']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.45, w, 0.6, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 1.75], [x + w + gap, 1.75]], true);
    }
    chip(s, 0.45 + w + gap - 0.3, 2.13, w + 0.6, 'Keep known_customer_safety_issue', C.goldSoft, C.gold, 0.3);
    eyebrow(s, 'The new CP reports it again?', 0.45, 2.7, 6, RED);
    box(s, 0.45, 2.95, 2.75, 1.95, C.goldSoft, C.gold);
    T(s, [
      { text: 'Ask C for proof', options: { bold: true, fontFace: HEAD, fontSize: 12, color: C.gold, breakLine: true } },
      { text: 'before they can book again. Does it reasonably address what the CP reported?', options: { fontSize: 9.5 } },
    ], { x: 0.62, y: 2.95, w: 2.45, h: 1.95, valign: 'middle', paraSpaceAfter: 4 });
    headPanel(s, 3.35, 2.95, 3.0, 1.95, 'Sufficient', ['Invalidate known_customer_fraud', 'Log a note, help them book'], C.teal, C.tealSoft, 9);
    headPanel(s, 6.5, 2.95, 3.05, 1.95, 'Not sufficient', ['Keep the flag', 'FC started? 1 more free month', 'Not up to safety standards yet. Need more time?'], RED, PINK, 9);
  }

  {
    const s = content(BIO, '2nd Response: they can\'t fix it', 'What you do depends on the membership status.',
      'Customer says they cannot resolve the issue. Active FC: proactively cancel FC and waive the ETF. If they request a refund, we can refund each Monthly Fee where the customer did not receive a cleaning within that MF\'s paid-through coverage. If they did have a cleaning within that MF\'s paid-through coverage, there\'s no refund unless another criterion under the Transformation guidelines makes them eligible (e.g. a Material Legal Threat). Unused DHJ voucher: cancel FC if the table has already triggered; refund and invalidate the voucher via CRM. FCF, no completed job: cancel FC via CRM. The ETF does not apply, since with no completed job it was never live. Refund the $19 base MF and any add-on charge (4hr/6hr upgrade) in full: the customer received zero cleanings, so both charges are refunded regardless of fault. Non-FC member: tell them you can\'t assist with their request right now, and invite them back once conditions improve. Customer doesn\'t respond and the Timed Reminder triggers: same as "cannot resolve the issue".');
    table(s, ['Membership status', 'Action', 'If C asks for a refund'], [
      ['Active FC', 'Proactively cancel FC and waive the ETF', 'Refund MFs with no cleaning in their paid-through coverage. Had a cleaning? No refund, unless another guideline applies (e.g. Material Legal Threat)'],
      ['Unused DHJ voucher', 'Cancel FC if the table already triggered', 'Refund and invalidate the voucher via CRM'],
      ['FCF, no completed job', 'Cancel FC via CRM. No ETF: it was never live', 'Refund the $19 base MF + any 4hr/6hr add-on in full'],
      ['Non-FC member', 'Can\'t assist right now. Invite them back once conditions improve', 'N/A'],
    ], { y: 1.45, colW: [1.75, 3.05, 4.3], fontSize: 9, rowH: 0.55 });
    await tip(s, 4.45, 'No reply when the TR triggers?', 'handle it the same as "can\'t fix it".', 'FiBellOff');
  }

  {
    const s = content(BIO, 'Needs more time, or won\'t cooperate?', 'Keep the booking block in place and give them room.',
      'Customer needs more time: if the customer\'s FC membership has already started, issue free weeks/months equivalent to what they requested. Advise them to reach out again once their home condition improves, so we can help them resume services. Customer refuses to provide evidence or won\'t cooperate: don\'t invalidate the known_customer_fraud flag. If their FC has already started, issue 1 more free month to give them more time to sort out their home conditions. Explain that we take this seriously and need proof the home meets safety standards. Create another Timed Reminder so the customer isn\'t charged another MF. If they come back and still refuse after that, treat it the same as "cannot resolve the issue". Customer doesn\'t respond and the Timed Reminder triggers: same as "cannot resolve the issue".');
    await cards(s, 1.45, 2.3, [
      { ico: 'FiClock', title: 'Needs more time', body: 'FC started? Issue free weeks or months matching what they asked for. Ask them to reach out once conditions improve.' },
      { ico: 'FiXCircle', title: 'Won\'t give evidence', body: 'Keep known_customer_fraud. FC started? 1 more free month. Explain we need proof the home is safe. New TR so no MF is charged.' },
      { ico: 'FiBellOff', title: 'No response', body: 'The TR triggers with no reply: handle it as "can\'t fix it".' },
    ]);
    await warn(s, 3.95, 'Still refusing after the extra time?', 'treat it as "can\'t fix it".');
  }

  await knowledgeCheck(BIO, [
    ['C says a CP called their home dirty. No CP-side action item. Restrict the account?', 'No. Only act once a CP-side agent validated the evidence and created the action item.'],
    ['Which cancellation reason code do you use for the RC and upcoming jobs?', 'Escalated Safety Issue.'],
    ['Active FC, 1st Response. What do you give C, and when does the TR fire?', '1 free month. TR the day before the free month expires (any CS).'],
    ['C says it\'s resolved. Which flag do you invalidate?', 'known_customer_fraud only. known_customer_safety_issue stays.'],
    ['FCF, no completed job, can\'t fix it, wants a refund. What happens?', 'Cancel FC (no ETF). Refund the $19 base MF + any add-on in full.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(BIO, [
    ['Is it a valid biohazard or hoarder report?', 'Is there a CP-side T3 action item?'],
    ['Which response are you on?', '1st or 2nd? What\'s the membership status?'],
    ['What actions are you going to take?', 'Jobs, flags, note, free month, TR, refund.'],
  ], 'TRAINER: share a live ticket with a biohazard/unsanitary T3 action item (or a customer replying after the 1st Response). Give trainees a few minutes to review it on their own, then work through it together. 1) Is it a validated report: is there a CP-side T3 action item, and does it meet the biohazard or hoarder definition? 2) Which response are you on, and what\'s the membership status (Active FC, unused DHJ voucher, FCF with no completed job, non-FC member)? 3) What actions are you going to take? 1st Response: cancel the RC and upcoming jobs (Escalated Safety Issue), add known_customer_fraud with the comment format, internal note, 1 free month + TR (Active FC), no DHJ retention offers (unused voucher). 2nd Response: resolved, can\'t fix it, needs more time, or won\'t cooperate. Then open the floor for questions.');

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Thank you!', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Close the module. Recap: Customer Agreed to Keep MCT (don\'t warn about the ETF when they\'ve already agreed to finish the term; confirm their intent; TR for the day after the MF that fulfills the MCT; probe kindly; never waive the ETF just because they didn\'t book; optional retention attempt; when the TR triggers, cancel if they didn\'t book again, or check in and set a new TR if they did). CP Reported C\'s Home as Biohazardous or Unsanitary (validated reports only; 1st Response: cancel the RC and jobs, known_customer_fraud flag, internal note, 1 free month + TR for Active FC; 2nd Response: resolved, can\'t fix it, needs more time, won\'t cooperate; refunds depend on membership status).');
  }

  if (topicSlides.length * 3 !== agendaLinks.length) throw new Error(`agenda has ${agendaLinks.length / 3} rows but the deck has ${topicSlides.length} topics`);
  agendaLinks.forEach((l, i) => { agendaSlide._rels.find(r => r.rId === l._rId).Target = String(topicSlides[Math.floor(i / 3)]); });
  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
