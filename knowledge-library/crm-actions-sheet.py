"""Build the "Knowledge Library CRM actions" workbook from the extracted inventory.

Inputs: crm_actions.csv (one row per CRM action per article per trigger, with the source line in the
Care Knowledge Library export) and crm_actions_notes.txt (inconsistencies and gaps found while extracting).
Output: an .xlsx with three tabs (CRM actions, Inconsistencies, About), ready to upload as a Google Sheet.

Usage: python3 crm-actions-sheet.py crm_actions.csv crm_actions_notes.txt out.xlsx "<extraction date>"
"""
import csv
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

src_csv, notes_txt, out, extracted = sys.argv[1:5]
TEAL, PALE = "0B7A6F", "E3F4F1"
head_font = Font(bold=True, color="FFFFFF")
head_fill = PatternFill("solid", fgColor=TEAL)
wrap = Alignment(wrap_text=True, vertical="top")

rows = list(csv.DictReader(open(src_csv, encoding="utf-8")))
wb = Workbook()

ws = wb.active
ws.title = "CRM actions"
cols = ["#", "Section", "Article", "When (trigger)", "CRM action", "Where in CRM", "Path stated in KL?", "Detail",
        "Moved to Notion", "Source line"]
widths = [5, 24, 30, 40, 40, 34, 12, 60, 24, 10]
ws.append(cols)
for i, r in enumerate(rows, 1):
    where = r["Where in CRM"].strip()
    ws.append([i, r["Section"], r["Article"], r["When (trigger)"], r["CRM action"], where,
               "Yes" if where else "No", r["Detail"], r["Moved to Notion"], r["Source line"]])
for j, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(j)].width = w
for c in ws[1]:
    c.font, c.fill, c.alignment = head_font, head_fill, Alignment(vertical="center", wrap_text=True)
for row in ws.iter_rows(min_row=2):
    for c in row:
        c.alignment = wrap
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

wn = wb.create_sheet("Inconsistencies")
wn.append(["Inconsistencies and gaps found in the Knowledge Library while building this list (line numbers refer to the extracted export)"])
wn["A1"].font, wn["A1"].fill = head_font, head_fill
for line in open(notes_txt, encoding="utf-8").read().rstrip().split("\n"):
    wn.append([line])
wn.column_dimensions["A"].width = 140
for row in wn.iter_rows(min_row=2):
    row[0].alignment = Alignment(wrap_text=True, vertical="top")
    if row[0].value and row[0].value[:2].strip().rstrip(".").isdigit():
        row[0].font = Font(bold=True)

wa = wb.create_sheet("About")
about = [
    ("What this is", "Every CRM action the Care Knowledge Library describes, one row per action per article per trigger, in document order."),
    ("Source", "Care Knowledge Library Google Doc (1igeJMGuPe2t4ajDmQTOOitb5amy68gY6q4ImWVoaNf8), text extracted on " + extracted + "."),
    ("Rows", f"{len(rows)} actions from {len({r['Article'] for r in rows})} articles."),
    ("Path stated in KL?", "Yes when the article gives a CRM path or screen. No when it says what to do but not where (e.g. add credits, set a TR). Filter on it to see only actions with a stated path."),
    ("Moved to Notion", "Tabs that say they moved to Notion are no longer updated in the Knowledge Library, so those rows can be out of date. Check the Notion article."),
    ("Not included", "Steps done only outside the CRM (Stripe, Slack, Google forms and trackers, email) and actions shown only in screenshots."),
    ("Source line", "Line in the extracted text where the action appears, for checking the wording."),
]
wa.append(["Field", "Notes"])
for a in about:
    wa.append(list(a))
for c in wa[1]:
    c.font, c.fill = head_font, head_fill
wa.column_dimensions["A"].width = 22
wa.column_dimensions["B"].width = 120
for row in wa.iter_rows(min_row=2):
    row[0].font = Font(bold=True, color=TEAL)
    for c in row:
        c.alignment = wrap

wb.save(out)
print("saved", out, len(rows), "rows")
