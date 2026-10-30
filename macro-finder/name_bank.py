"""Build the trainee Macro Search Bank: exact macro names only, split into DHJ and FCF tabs.

Each Comms Kit customer macro goes to the DHJ tab, the FCF tab, or both (Membership
Model "DHJ, FCF"). Categories are the three the Care team asked for; a macro can sit
under more than one. When the category is a judgment call it carries a review note,
and macros that fit none of the three go under "Needs Category Review".

Usage: python3 name_bank.py <commskit.json> <out_dir>
"""
import json
import os
import sys

from assignments import ASSIGN

FC = "ForeverClean Membership"
JS = "Job and Service Issues"
JC = "Job-Related Charges"
RV = "Needs Category Review"
UV = "Unused DHJ Voucher"       # DHJ tab only
PI = "FCF Pre-Invoice"          # FCF tab only
ORDER = [FC, UV, PI, JS, JC, RV]
UNCLEAR = "⚠ Category unclear: please review"

# Subcategory (from assignments.py) -> list of (category, needs_review)
BY_SUB = {
    "Membership Inquiry": [(FC, False)], "Membership Cancellation": [(FC, False)],
    "Cancellation Closure (Graceful Close)": [(FC, False)], "Early Termination Fee (ETF)": [(FC, False)],
    "ETF Waiver": [(FC, False)], "Retention Strategy & Offers": [(FC, False)],
    "Membership Fee Refund": [(FC, False)], "Membership Fee": [(FC, False)],
    "Scheduling / Rescheduling": [(JS, False)], "Cleaner Matching & Requests": [(JS, False)],
    "Cleaner No-Show": [(JS, False)], "Cleaner Cancellation": [(JS, False)], "Cleaning Quality": [(JS, False)],
    "Job Refund Request": [(JS, False)], "Unauthorized Reschedule": [(JS, False)],
    "Service Complaints": [(JS, False)], "Recovery Offers": [(JS, False)],
    "Lockout": [(JS, False), (JC, False)],
    "Overcharged Hours": [(JS, False), (JC, False)],
    "Unauthorized Addition of Hours": [(JS, False), (JC, False)],
    "False Invoice": [(JS, False), (JC, False)],
    "Last-Minute Cancellation (LMC)": [(JC, False)],
    "Cleaning Cost": [(JC, False)], "Priority Fee": [(JC, False)], "Processing Fee": [(JC, False)],
    "Pre-Authorization Fee": [(JC, False)], "Invoices & Receipts": [(JC, False)],
    "Refund Confirmation": [(JC, False)], "Platform Fee & Sales Tax": [(JC, False)], "Tips": [(JC, False)],
    "Booking Help & Service FAQs": [(JS, True)], "Service Coverage": [(JS, True)],
    "Biohazard / Unsafe Home": [(JS, True)], "Other T&S Concerns": [(JS, True)],
    "Contacting the Cleaner": [(JS, True)],
    "Chargebacks & Bank Disputes": [(FC, True)], "Vouchers & Redemption": [(FC, True)],
    "Legal & Social Media Threats": [(FC, True)],
    "Failed Payments": [(FC, False)], "Card on File": [(RV, True)], "Callbacks": [(RV, True)], "Dashboard Links": [(RV, True)],
    "Account Updates": [(RV, True)], "Unsubscribe Requests": [(RV, True)],
    "Account Deactivation": [(RV, True)], "Reviews & Feedback": [(RV, True)],
    "Customer Follow-Up": [(RV, True)], "General Responses": [(RV, True)],
}

# Per-row overrides where one macro answers more than one concern.
BY_ROW = {
    4: [(JS, False), (FC, False)],     # FCF poor first cleaning with cancellation intent
    5: [(FC, False), (JS, False)],     # ETF waiver after two consecutive serious-issue jobs
    65: [(FC, False), (JS, False)],    # ETF waiver based on next service outcome
    90: [(JS, False), (FC, False)],    # second poor cleaning, membership cancelled at no charge
    143: [(FC, False)],                # card failed on the FC charge
    144: [(FC, False)],
    145: [(JC, False)],                # card failed on an invoiced job
    108: [(JC, True)],                 # move-out extra hours beyond the voucher
}

TABS = {"DHJ": ("DHJ",), "FCF": ("FCF",), "DHJ, FCF": ("DHJ", "FCF")}

NOT_IN_ARTICLE = "⚠ Not named in the KB article: please review"

# Extra rows for the scenario categories the Care team added. Each macro also keeps its
# rows elsewhere. KB = "Unused DHJ Voucher" and "FCF Pre-Invoice Cancellation and Refund
# Requests" articles, Related Macros sections.
# (tab, category, customer concern, Comms Kit row, note)
SCENARIO_ROWS = [
    ("DHJ", UV, "Wants the voucher refunded", 110, ""),
    ("DHJ", UV, "Wants the voucher refunded", 109, ""),
    ("DHJ", UV, "Wants to try the service first", 112, ""),
    ("DHJ", UV, "Only wanted a one-time cleaning", 111, ""),
    ("DHJ", UV, "Wants their account closed", 166, ""),
    ("DHJ", UV, "Asks why a card is needed", 114, NOT_IN_ARTICLE),
    ("DHJ", UV, "Asks whether the voucher is active", 116, NOT_IN_ARTICLE),
    ("FCF", PI, "Wants a refund before the first cleaning", 187, ""),
    ("FCF", PI, "Wants a refund before the first cleaning", 188, ""),
    ("FCF", PI, "Wants a refund before the first cleaning (service issue)", 189, ""),
    ("FCF", PI, "Charged 2 or more membership fees, no cleaning yet", 190, ""),
    ("FCF", PI, "Wants to cancel before the first cleaning", 3, NOT_IN_ARTICLE),
]
SHARED_NAME = "Two macros share this exact name; check the wording in the CRM"


def main(ck_path, out):
    rows = {r[0]: r for r in json.load(open(ck_path))["rows"] if r[2] == "Customer"}
    names = [rows[n][5] for n in ASSIGN]
    tabs = {"DHJ": [], "FCF": []}
    for n, (_main, sub, _scen, _when, _kw, _art) in ASSIGN.items():
        name = rows[n][5]
        for tab in TABS.get(rows[n][9], ()):
            for cat, unclear in BY_ROW.get(n, BY_SUB[sub]):
                note = UNCLEAR if unclear else ""
                if names.count(name) > 1:
                    note = (note + "; " if note else "") + SHARED_NAME
                tabs[tab].append([cat, sub, name, note])
    for tab, cat, concern, n, note in SCENARIO_ROWS:
        tabs[tab].append([cat, concern, rows[n][5], note])
    os.makedirs(out, exist_ok=True)
    for tab, data in tabs.items():
        data.sort(key=lambda r: (ORDER.index(r[0]), r[1], r[2].lower()))
        json.dump(data, open(f"{out}/{tab}.json", "w"), ensure_ascii=False)
        counts = {c: sum(1 for r in data if r[0] == c) for c in ORDER}
        print(tab, len(data), "rows,", len({r[2] for r in data}), "distinct names", counts,
              "flagged", sum(1 for r in data if r[3].startswith("⚠")))
        new = [r for r in data if r[0] in (UV, PI)]
        print(json.dumps(new, ensure_ascii=False))


if __name__ == "__main__":
    main(*sys.argv[1:3])
