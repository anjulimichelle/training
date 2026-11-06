# Ticket Assessment Hub (staging prototype)

A standalone app for Care Specialist training. Trainers build ticket-based quizzes, generate a trainee form, collect submissions, score them against a per-ticket rubric with Claude, review and override the scores, and export the results.

This is a **staging prototype**. It sends nothing to trainees, publishes no forms, and touches no existing dashboards, sheets or forms. It ships with fictional demo data only.

## claude.ai artifact version

The simple trainer version is a claude.ai artifact in the Training Agenda Library style: https://claude.ai/artifact/FgJxWnL73dezrTf77zNJPs (private until shared from its Share menu). Three tabs:

1. **Quiz Generation**: title, date, pasted ticket links and trainee emails. "Create Google Form" opens the trainer's Form Maker (`apps-script/FormMaker.gs`, deployed once as a web app that runs as the trainer) and builds the form in their Drive with the instructions and, per ticket, a section with the link and the four questions. It does not publish, share or create response sheets. "Find my form and response sheet" then finds both in Drive by title (read only).
2. **Responses**: pick a quiz, pull answers from the response sheet (read only, through the trainer's Google Sheets connector) or a CSV, and see each trainee's raw answers by ticket.
3. **Scoring**: points and scoring guidelines per ticket, scores and errors per trainee. Claude scores one ticket for all trainees per call. Changing a ticket's points or guidelines marks its scores out of date until re-scored; trainer adjustments survive re-scoring.

Source: `artifact/page.src.html` and `artifact/app.src.js`. Build with `node artifact/build.js`, which writes `artifact/dist/ticket-assessment-hub.html`.

## Run it

Requires Node 20 or later.

```bash
cd ticket-assessment-hub
```

```bash
npm install
```

```bash
npm start
```

Open http://127.0.0.1:4317, click **Setup**, then **Load fictional demo data**.

Run the automated staging checks (the 15 tests from the spec plus extras, all against local mocks):

```bash
npm test
```

## Configuration

All secrets are server-side environment variables. None reach the browser or the trainee form.

| Variable | Purpose | Default |
|---|---|---|
| `ANTHROPIC_API_KEY` | Turns on Claude grading. Without it, an offline word-overlap heuristic (not AI, clearly labelled) scores staging tests. | unset |
| `CLAUDE_MODEL` | Model used for grading. | `claude-opus-5-5` |
| `CLAUDE_EFFORT` | Effort level for grading. | `high` |
| `CLAUDE_FALLBACKS` | Set to `off` to disable the server-side refusal fallback. | on |
| `EVALUATOR` | Set to `heuristic` to force the offline scorer even with a key. | auto |
| `APPS_SCRIPT_URL` / `APPS_SCRIPT_TOKEN` | Connects the Google Forms and Sheets connector (see below). | unset |
| `GOOGLE_EXTERNAL_ACTIONS_APPROVED` | Must be `yes` before the app may create a Google Form or write to Google Sheets. Reading responses does not need it. | unset (blocked) |
| `TRAINER_PASSCODE` | Requires a passcode for every trainer API call. Set it before running anywhere other than your own machine. | unset |
| `PORT` / `HOST` / `DATA_FILE` | Server binding and the staging JSON database path. | `4317` / `127.0.0.1` / `data/staging-db.json` |

## Connecting Google

The app reaches Google through a small Apps Script web app (`apps-script/Code.gs`), because Apps Script has native Forms and Sheets access and keeps the Google credentials inside the trainer's Google account.

1. In the Google account that should own the forms, create a new Apps Script project. Paste `apps-script/Code.gs` and `apps-script/appsscript.json` (enable "Show appsscript.json" in project settings).
2. Run `setupDatabase()` once. It creates a spreadsheet named "Ticket Assessment Hub DB (staging)" with the six tabs (Quiz Registry, Raw Responses, Evaluation Results, Score Breakdown, Areas for Improvement, Cohort Performance) and a random API token.
3. Run `showToken()` and copy the token from the log into `APPS_SCRIPT_TOKEN`.
4. Deploy as a web app: execute as **Me**, access **Anyone**. Copy the `/exec` URL into `APPS_SCRIPT_URL`. The token check is what keeps it private; rotate it with `rotateToken()`.
5. Only after you approve real Google actions, set `GOOGLE_EXTERNAL_ACTIONS_APPROVED=yes`.

What the connector does:

- **Create form**: one Google Form per quiz. Name, email and cohort first, then one section per ticket titled "Ticket N" with the ticket link (clickable) and description, then the four standard long-answer questions. Responses go to a new "Form Responses" tab in the database spreadsheet. The form is created **closed and unpublished**; a trainer opens it by hand.
- **Fetch responses**: reads form responses and maps them back to tickets by item ID. Re-syncing never creates duplicates (keyed by Google response ID).
- **Sync sheets**: rewrites the six hub tabs. Google's own "Form Responses" tabs are never modified.

The connector never sends email, never shares files, and never receives rubrics, reference material or expected answers.

## How it works

- `server/core.js`: shared domain logic (questions, validation, trainee-safe view, scoring math, heuristic scorer, Claude prompt and JSON schema, export tables). Runs in Node and the browser.
- `server/server.js`: HTTP server and JSON API. Staging data is a JSON file written atomically.
- `server/evaluator.js`: Claude grading, one call per ticket in parallel, with structured JSON output. Scores are clamped to each criterion's maximum and totals are recomputed on the server, so model arithmetic is never trusted. Missing criteria are flagged for trainer review.
- `server/exporter.js`: CSV and Excel exports.
- `public/`: the trainer app (four tabs) and the staging trainee form.
- `apps-script/`: the Google connector.

Data rules the app enforces:

- Raw submissions are stored once with a SHA-256 hash of the answers. Evaluation never writes to them, and a failed evaluation leaves them untouched.
- One current evaluation per submission. A second "Evaluate" is refused; "Request reevaluation" creates a new version and keeps the old one (marked superseded) for the audit trail.
- The AI evaluation and the trainer evaluation are stored separately. Trainer edits, approvals and overrides carry timestamps and the reviewer's name.
- Pass or fail is shown only when the quiz has a passing threshold.
- The dashboard's improvement summary states how many evaluated submissions it is based on, and shows nothing when there are none.

---

## Implementation report

### What has been built

A standalone Node app in `ticket-assessment-hub/` with the four requested tabs, a staging trainee form, a Google Apps Script connector, and an automated end-to-end test. It does not modify anything else in this repository.

### Fully functional (tested in this session)

- **Quiz Builder**: quiz details, cohort or individual assignment, statuses (Draft, Ready, Sent, Closed; "Sent" is a label only), add, remove, duplicate and reorder tickets, per-ticket URL, description, context, attachments (links or uploads), per-ticket rubric with any criteria, points and deductions, critical errors, alternatives and auto-fail conditions, per-ticket reference materials, validation (points must add up, URLs required, duplicate URLs flagged), preview of the trainee form beside the trainer-only rubric, explicit approval before generation, copy link and a "Send" dialog that only drafts a message.
- **Staging trainee form**: looks like the Google Form layout (identity page, then one section per ticket with the clickable link and the four questions). It receives only the trainee-safe view; the test confirms no rubric or reference text leaks.
- **Quiz Responses**: response table with every requested column and the five statuses, full submission view grouped by ticket in question order, fictional sample submissions.
- **Evaluation & Scoring**: per-ticket score, breakdown, strengths, areas for improvement, coaching and critical errors; overall totals, strengths, recurring weaknesses and coaching focus; trainer edits per criterion, added or removed improvement items, comments, result override, approve, and reevaluation with version history.
- **Performance Dashboard**: the six summary metrics, all six filters, trainee table, per-trainee history, quiz comparison matrix, improvement summary with occurrence counts and trainee names.
- **Exports**: Excel and CSV for raw responses, evaluation results, score breakdown, areas for improvement, cohort performance and the quiz registry; a six-tab full workbook; individual trainee reports. CSV cells are escaped against spreadsheet formula injection.

### Integrations

| Integration | Status in this session |
|---|---|
| Claude API | Code complete and tested against a local mock of the Messages API. **Not run against the real API**: this session has no Anthropic API key (the API returned 401). |
| Google Forms and Sheets | Apps Script connector written and its contract tested against a local mock. **Not deployed**: this container's network allows only Google Drive paths, not the Forms API, and no Google writes were authorized. |
| Google Drive | Available to this session but not used, since nothing was approved for creation. |

### Needs authorization or credentials

1. An Anthropic API key in `ANTHROPIC_API_KEY` on the server.
2. A trainer deploying `apps-script/Code.gs` in the right Google account, then setting `APPS_SCRIPT_URL` and `APPS_SCRIPT_TOKEN`.
3. Your explicit approval, then `GOOGLE_EXTERNAL_ACTIONS_APPROVED=yes`, before any real form or sheet is created.
4. A hosting decision and `TRAINER_PASSCODE` (or SSO in front) before trainers outside your machine use it.

### How to test

1. `npm test` runs the 15 staging checks with mocks and writes sample exports to a temp folder.
2. `npm start`, then Setup, then Load fictional demo data. In Quiz Builder, open the rubric sections, click **Preview & generate form**, tick the approval box and generate the staging form.
3. Open the form link, submit as `charlie.trainee@example.com`. In Quiz Responses, add the fictional sample submissions.
4. In Evaluation & Scoring, run evaluations, change a criterion score, add your name, approve. Check that the AI box still shows the original score.
5. In Performance Dashboard, filter, open a trainee's history, and download the full workbook and a CSV.

### Known limitations

- Without an API key the scorer is a word-overlap heuristic. It is useful only for testing the workflow, does not check critical errors or auto-fail conditions, and flags everything for review.
- Claude grades only the text the trainer provides. It cannot open ticket URLs or screenshots, so the rubric and references must describe what matters.
- Staging storage is one JSON file and suits a single trainer team. Production would move the store to the Google Sheet (via the connector) or a database.
- Trainer access is a shared passcode, not per-person login. Reviewer names are self-entered.
- Uploaded attachments are stored on the staging server. In Google mode, only `https://` attachment links are added to the form.
- The passing rate counts AI scores that are not yet approved; the metric's caption says how many results it includes.
- Changing a quiz after its form is generated marks the form out of date; regenerating creates a new form rather than editing the old one.
