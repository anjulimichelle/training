# Agent Macro Finder (prototype)

The prototype is a Google Sheet in Anjuli's Drive. Care agents type a customer's concern in plain words and get the best-matching macro, its full wording, the CRM steps and the policy reminders. It hasn't been shared with agents.

- Sheet: [Agent Macro Finder — PROTOTYPE](https://docs.google.com/spreadsheets/d/1CItA-i07xSegblnDUheBf-GKRh7p5av1oyqZE2UkLBQ/edit)
- Macro wording comes from the [Care Transformation Comms Kit (CS Copy)](https://docs.google.com/spreadsheets/d/1T2Q7uBOMW6zwBAG8EZ3_e8RMmjqF9qeRgDe4tPB9i-Y/edit), tab "All Macros — New taxonomy". The Knowledge Library names macros under "Related Macros" and points to this Comms Kit for the text.
- Scenarios, policy reminders and 12 inline sample scripts come from the [Homeaglow Care Knowledge Library](https://docs.google.com/document/d/1igeJMGuPe2t4ajDmQTOOitb5amy68gY6q4ImWVoaNf8/edit).

## What's in the sheet

| Tab | Purpose |
|---|---|
| 🔎 MACRO FINDER | Search box, four filters, full macro view, top 25 results |
| 📚 MASTER MACRO LIBRARY | 181 macros (169 Comms Kit, 12 Knowledge Library samples), the 15 required columns plus search helpers. Protected. |
| ⚠️ MACRO REVIEW | 29 items: duplicates, conflicts, outdated wording, missing responses, missing sources, suggested categories |
| 🗂 CATEGORY MAP | 7 main categories, 58 subcategories (20 marked "Recommended new"), keywords, customer phrases, policy reminders, scenario list |
| 🔤 SEARCH SYNONYMS | Word triggers and the extra terms they search for, plus ignored words |
| 🛠 UPDATE PROCESS | Refresh steps, automatic-sync proposal, agent access options, update log |
| ℹ️ ABOUT | How it works, sources, glossary |

Every macro is set to Pending Review. The Finder's "Showing" box is set to "Approved + Pending (prototype)" so the prototype has results; switch it to "Approved only" before launch.

## Files here

- `taxonomy.py`: category tree, scenarios, synonyms, ignored words, policy reminders.
- `assignments.py`: category, scenario, "when to use", extra keywords and Knowledge Library article for each Comms Kit row, plus the 12 Knowledge Library samples.
- `review.py`: the review items.
- `build.py`: turns a Comms Kit snapshot and a Knowledge Library text export into the cell data the sheet was built from. Comms Kit wording is never retyped: Master rows point at the copied Comms Kit tab with `=INDEX(SRC_BODY, row)` and the sheet freezes those to values.
- `apps-script/WebApp.gs` and `apps-script/Index.html`: the recommended agent-facing web app (Option A below).
- `apps-script/Sync.gs`: proposed daily sync from the Comms Kit. It never overwrites an Approved macro; changes become review rows.

Rebuild the data:

```bash
python3 build.py <commskit.json> <kb.txt> <out_dir>
```

## How search ranks results

Each typed word, plus its synonyms, scores 13 when it appears in the macro name, 12 in the keywords, phrases, category, scenario or KB article, and 11 in the response text. Words match partially, so "cancel" finds "cancellation". A macro that matches more of the typed words ranks higher. Ignored words such as "the" and "customer" don't count.

## Giving agents access

A view-only user can't type in a Google Sheet's search box, so a sheet-only Finder can't serve view-only agents. The options are on the 🛠 UPDATE PROCESS tab:

- Option A, which Claude recommends, is to deploy `WebApp.gs` and `Index.html` as an Apps Script web app that executes as the owner. Each agent gets a private search page with a Copy button, sees only Approved macros, and needs no access to the sheet.
- Option B is a personal copy per agent. Copies go stale after each update.
- Option C is edit access to the Finder tab only. Agents searching at the same time overwrite each other's search.

To install Option A: open the sheet, go to Extensions › Apps Script, add `WebApp.gs`, `Index.html` and (optionally) `Sync.gs`, then Deploy › New deployment › Web app with "Execute as: Me" and "Who has access: Anyone in your organization". For the sync, run `syncFromCommsKit` once to authorize, then add a daily time-driven trigger.
