"""Build the cell data for the Macro Finder spreadsheet.

Inputs:  a JSON snapshot of the Comms Kit tab "All Macros — New taxonomy" (rows 1–194)
         and a plain-text export of the Knowledge Library Google Doc.
Outputs: out/*.json, each a 2D array ready for a Sheets values/formulas update.

Comms Kit wording is never typed out here: Master rows point at the copied source tab
with =INDEX(SRC_*, row) and the sheet freezes those to values, so the text is the
server's own copy. Only the 12 KB sample entries carry text from this script.

Usage: python3 build.py <commskit.json> <kb.txt> <out_dir>
"""
import json
import os
import re
import sys

from assignments import ASSIGN, KB_SAMPLES
from review import ITEMS, ETF_ROWS
from taxonomy import CATEGORIES, SCENARIOS, STOPWORDS, SYNONYMS

TICKET = {"DHJ": "DHJ", "FCF": "FCF", "DHJ, FCF": "Both"}
ETF_WARNING = ("⚠ This response mentions the ETF, minimum term or cancellation fee. Don't send it as-is "
               "to Washington State customers (ETF $0, MCT 0) or M2M/Prepaid plans (no ETF). "
               "(KB: M2M & Prepaid Plans; Washington State — Temporary FC Cancellation Handling)")


def clean_sample(text):
    text = re.sub(r"^Sample:\s*", "", text.strip())
    if text[:1] in "“\"" and text[-1:] in "”\"":
        text = text[1:-1]
    return text.strip()


def main(ck_path, kb_path, out):
    ck = json.load(open(ck_path))
    rows = {r[0]: r for r in ck["rows"] if r[2] == "Customer"}
    kb = open(kb_path, encoding="utf-8-sig").read().split("\n")
    order = {(c[0], c[1]): i for i, c in enumerate(CATEGORIES)}

    entries = []
    for n, (main_, sub, scen, when, kw, art) in ASSIGN.items():
        r = rows[n]
        entries.append(dict(src=n, name=None, tt=TICKET.get(r[9], "Other / Undefined"), main=main_, sub=sub,
                            scen=SCENARIOS[scen], when=when, kw=kw, art=art,
                            path=f"{r[3]} › {r[4]}", etf=n in ETF_ROWS))
    for name, main_, sub, scen, when, kw, tt, art, (a, b) in KB_SAMPLES:
        entries.append(dict(src=None, name=name, tt=tt, main=main_, sub=sub, scen=SCENARIOS[scen], when=when,
                            kw=kw, art=art, path="", etf=False, body=clean_sample("\n".join(kb[a - 1:b]))))
    entries.sort(key=lambda e: (order[(e["main"], e["sub"])], e["src"] or 10_000))
    for i, e in enumerate(entries, 1):
        e["id"] = f"MF-{i:03d}"

    by_src = {e["src"]: e for e in entries if e["src"]}
    by_name = {e["name"]: e for e in entries if e["name"]}

    def ref(x):
        return by_src[x] if isinstance(x, int) else by_name[x]

    review = []
    for i, (kind, refs, details, evidence, action) in enumerate(ITEMS, 1):
        rid = f"RV-{i:03d}"
        if refs == "NO_KB":
            refs = [e["src"] for e in entries if e["src"] and not e["art"]]
            details = details.format(n=len(refs))
        es = [ref(x) for x in refs]
        for e in es:
            e.setdefault("reviews", []).append(rid)
        ids = ", ".join(sorted(e["id"] for e in es))
        review.append([rid, kind, ids, len(es) or "", details, evidence, action, "Pending Review", "", "", ""])

    master = [[
        "Macro ID", "Macro Name", "Ticket Type", "Main Category", "Subcategory", "Customer Scenario",
        "Search Keywords", "Common Customer Phrases", "When to Use", "Full Macro Response",
        "Required CRM Actions", "Policy Reminders", "Knowledge Library Source", "Last Reviewed",
        "Review Status", "Wording Source", "Source Row", "Comms Kit Category Path", "Review Items",
        "Preview", "_name", "_meta", "_body", "KB Link", "KB Article(s)", "_extra_kw", "_policy_add",
    ]]
    for e in entries:
        s = e["src"]
        master.append([
            e["id"],
            f"=INDEX(SRC_NAME,{s})" if s else e["name"],
            e["tt"], e["main"], e["sub"], e["scen"], "", "", e["when"],
            f"=INDEX(SRC_BODY,{s})" if s else e["body"],
            f"=INDEX(SRC_NOTES,{s})" if s else "",
            "", "", "", "Pending Review",
            "Comms Kit › All Macros — New taxonomy" if s else "Knowledge Library (inline sample)",
            s or "", e["path"], ", ".join(e.get("reviews", [])),
            "", "", "", "", "", e["art"], e["kw"], "ETF" if e["etf"] else "",
        ])

    cat = [["Main Category", "Subcategory", "Status", "Base Search Keywords", "Common Customer Phrases",
            "Policy Reminder", "Macros in Library"]]
    for main_, sub, new, kws, phrases, policy in CATEGORIES:
        cat.append([main_, sub, "Recommended new" if new else "Starting list", kws,
                    phrases.replace(" | ", "\n"), policy, ""])
    # The sheet appends this to Policy Reminders for rows flagged "ETF" in _policy_add.
    cat[1] += ["", "", "", ETF_WARNING]

    scen = [["Customer Scenario"]] + [[v] for v in SCENARIOS.values()]
    syn = [["Trigger (start of a typed word, regex)", "Also search for", "", "", "Ignored words"]]
    stop = sorted(set(STOPWORDS) | {"to", "of", "in", "on", "is", "it", "an", "or", "me", "we", "be",
                                    "do", "so", "up", "at", "by", "as", "if", "my", "am", "im"})
    for i in range(max(len(SYNONYMS), len(stop))):
        t, x = SYNONYMS[i] if i < len(SYNONYMS) else ("", "")
        syn.append([t, x.replace("'", ""), "", "", stop[i] if i < len(stop) else ""])

    os.makedirs(out, exist_ok=True)
    for name, data in [("master", master), ("review", review), ("category_map", cat),
                       ("scenarios", scen), ("synonyms", syn)]:
        json.dump(data, open(f"{out}/{name}.json", "w"), ensure_ascii=False)
    new = sum(1 for c in CATEGORIES if c[2])
    print(f"{len(entries)} macros ({len(by_src)} Comms Kit, {len(by_name)} KB samples), "
          f"{len(review)} review items, {len(CATEGORIES)} subcategories ({new} recommended new)")


if __name__ == "__main__":
    main(*sys.argv[1:4])
