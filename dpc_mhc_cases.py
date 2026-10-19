#!/usr/bin/env python3
"""
Find High Court cases (courts.ie) where the Data Protection Commission is a party
and Mason Hayes & Curran is listed as a legal representative, and write them to Excel.

Usage:
    python3 dpc_mhc_cases.py [output.xlsx]

Cases with no concluding order go in the "Open cases" sheet; every matching case
(with its inferred status) is written to a second sheet so nothing is hidden.

Requires: openpyxl  (pip install openpyxl)
"""
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = "https://www.courts.ie"
SEARCH_URL = BASE + "/high-court-search/GetCases/"
DETAILS_URL = BASE + "/HCS/GetCaseRefDetails/"
PARTY = "data protection commission"   # substring match also catches "...COMMISSIONER"
FIRM = "mason hayes"
HEADERS = {"User-Agent": "Mozilla/5.0", "X-Requested-With": "XMLHttpRequest"}
DORMANT_DAYS = 365  # open cases with no activity for longer than this are flagged dormant
DELAY = 1.5  # seconds between requests; the server refuses connections if hit too fast

DPC_RE = re.compile(r"DATA PROTECTION COMMISSION", re.I)
# Roles on the DPC's own side vs the other side
PLAINTIFF_ROLES = {"plaintiff", "applicant", "appellant", "petitioner"}
DEFENDANT_ROLES = {"defendant", "respondent"}
# Order / result wording that indicates proceedings have been finally disposed of
CONCLUDED_RE = re.compile(
    r"\bfinal\b|struck out|strike out|discontinu|dismiss|settled|withdrawn|"
    r"judgment for|judgement for|order of certiorari|refused.*final|perfected final",
    re.I,
)


def request(url, data=None, tries=6):
    body = urllib.parse.urlencode(data).encode() if data is not None else None
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, body, HEADERS)
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.load(r)
        except Exception as e:
            wait = 5 * (attempt + 1)
            print(f"  request failed ({e}); retrying in {wait}s", file=sys.stderr)
            time.sleep(wait)
    raise RuntimeError(f"Giving up on {url}")


def search_all():
    results, page = [], 1
    while True:
        data = dict(
            page=page, proceeding="", caseRef="", caseNumber="", year="",
            caseparty1=PARTY, caseparty1checked="Both",
            caseparty2="", caseparty2checked="Both",
            setDownVenue="", setDownType="", setdownNumFrom="", courtdate="",
            listType="", appealcasereference="", supremecasereference="", sortBy="",
        )
        d = request(SEARCH_URL, data)
        batch = d.get("results") or []
        results += batch
        print(f"Search page {page}: {len(batch)} cases (total {d.get('total')})", file=sys.stderr)
        if not batch or page * 25 >= (d.get("total") or 0):
            return results
        page += 1
        time.sleep(DELAY)


def firm(cp):
    return ((cp.get("arch_legalfirm") or {}).get("arch_name") or "").strip()


def role(cp):
    return ((cp.get("arch_caserole") or {}).get("arch_name") or "").strip()


def side(r):
    r = r.lower()
    if r in PLAINTIFF_ROLES:
        return "P"
    if r in DEFENDANT_ROLES:
        return "D"
    return "O"


def d(s):
    """ISO datetime string -> date (or None)."""
    if not s or s.startswith("/Date(") or s.startswith("0001"):
        return None
    try:
        return datetime.fromisoformat(s[:19]).date()
    except ValueError:
        return None


def analyse(case, details):
    parties = details.get("case_parties") or case.get("case_parties") or []
    dpc = [p for p in parties if DPC_RE.search(p["arch_name"])]
    dpc_sides = {side(role(p)) for p in dpc}
    dpc_role = ", ".join(sorted({role(p) for p in dpc}))
    if "D" in dpc_sides:
        dpc_position = "Defendant / Respondent"
    elif "P" in dpc_sides:
        dpc_position = "Plaintiff / Applicant"
    else:
        dpc_position = dpc_role or "Other"
    dpc_side = "D" if "D" in dpc_sides else ("P" if "P" in dpc_sides else "O")
    opposite = {"D": "P", "P": "D"}.get(dpc_side)

    against = [p["arch_name"] for p in parties
               if opposite and side(role(p)) == opposite and not DPC_RE.search(p["arch_name"])]
    others = [f'{p["arch_name"]} ({role(p)})' for p in parties
              if side(role(p)) == "O" and not DPC_RE.search(p["arch_name"])]
    mhc = [f'{p["arch_name"]} ({role(p)})' for p in parties if FIRM in firm(p).lower()]
    dpc_firms = sorted({firm(p) for p in dpc if firm(p)})

    orders = sorted(details.get("orders") or [], key=lambda o: o.get("arch_orderdate") or "")
    listings = sorted(details.get("listings") or [], key=lambda l: l.get("arch_courtdate") or "")
    apps = details.get("applications") or []
    judgments = details.get("judgements") or []

    order_text = [((o.get("arch_ordertype") or {}).get("arch_ordertypename") or "") for o in orders]
    concluding = [(d(o.get("arch_orderdate")), t) for o, t in zip(orders, order_text) if CONCLUDED_RE.search(t)]
    listing_results = [((l.get("arch_result") or {}).get("arch_name") or "") for l in listings]
    concluding += [(d(l.get("arch_courtdate")), r) for l, r in zip(listings, listing_results) if CONCLUDED_RE.search(r)]

    today = datetime.now().date()
    future = [d(l.get("arch_courtdate")) for l in listings if d(l.get("arch_courtdate")) and d(l.get("arch_courtdate")) >= today]
    activity = [x for x in
                [d(o.get("arch_orderdate")) for o in orders]
                + [d(l.get("arch_courtdate")) for l in listings]
                + [d(a.get("arch_datereceived")) for a in apps]
                + [d(j.get("arch_delivereddate")) for j in judgments]
                + [d(case.get("arch_dateofissue"))] if x]
    last_activity = max(activity) if activity else None
    last_conclusion = max(concluding, key=lambda c: c[0] or today) if concluding else None

    # Active unless a concluding order exists and nothing has happened since it
    if future:
        status = "Active (future listing)"
    elif last_conclusion and (last_activity is None or last_conclusion[0] is None or last_activity <= last_conclusion[0]):
        status = "Concluded"
    elif last_activity and (today - last_activity).days > DORMANT_DAYS:
        status = (f"Open – likely stayed or dormant: no concluding order on record, but no filings, "
                  f"orders or listings since {last_activity:%d/%m/%Y}")
    elif last_conclusion:
        status = "Active (activity after final order)"
    else:
        status = "Active"

    if details.get("arch_appealcasereference") or details.get("supreme_court_ref"):
        status += " – appealed"

    return {
        "Case reference": case["arch_name"],
        "Title": case.get("arch_publiccasetitle", ""),
        "Date of issue": d(case.get("arch_dateofissue")),
        "DPC is": dpc_position,
        "DPC role (as recorded)": dpc_role,
        "Against": "; ".join(against),
        "Other parties": "; ".join(others),
        "MHC acts for": "; ".join(mhc),
        "DPC legal rep": "; ".join(dpc_firms),
        "Status (inferred)": status,
        "Last activity": last_activity,
        "Latest order": (f"{d(orders[-1].get('arch_orderdate'))}: {order_text[-1]}" if orders else ""),
        "Concluding order/result": (f"{last_conclusion[0]}: {last_conclusion[1]}" if last_conclusion else ""),
        "Next listing": min(future) if future else None,
        "Appeal ref": details.get("arch_appealcasereference") or details.get("supreme_court_ref") or "",
        "Link": f"{BASE}/high-court-search?caseRecordId={case['arch_name']}",
    }


def write_sheet(ws, rows):
    cols = list(rows[0].keys()) if rows else []
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(vertical="top", wrap_text=True)
    for r in rows:
        ws.append([r[c] for c in cols])
    widths = {"Case reference": 20, "Title": 50, "Date of issue": 12, "DPC is": 20, "Against": 40,
              "Other parties": 30, "MHC acts for": 40, "Status (inferred)": 45, "Latest order": 40,
              "Concluding order/result": 40, "Link": 30}
    for i, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 16)
        for cell in ws.iter_rows(min_row=2, min_col=i, max_col=i):
            cell[0].alignment = Alignment(vertical="top", wrap_text=True)
            if isinstance(cell[0].value, type(datetime.now().date())):
                cell[0].number_format = "DD/MM/YYYY"
    ws.freeze_panes = "B2"
    if rows:
        ws.auto_filter.ref = ws.dimensions


def main():
    out = next((a for a in sys.argv[1:] if not a.startswith("--")), "dpc_mason_hayes_curran_cases.xlsx")

    cases = search_all()
    mhc_cases = [c for c in cases if any(FIRM in firm(p).lower() for p in c.get("case_parties") or [])]
    print(f"{len(cases)} DPC cases; {len(mhc_cases)} list Mason Hayes & Curran", file=sys.stderr)

    rows = []
    for c in mhc_cases:
        time.sleep(DELAY)
        print(f"  details {c['arch_name']}", file=sys.stderr)
        rows.append(analyse(c, request(DETAILS_URL + c["arch_name"])))
    rows.sort(key=lambda r: r["Date of issue"] or datetime.min.date(), reverse=True)

    active = [r for r in rows if r["Status (inferred)"].startswith(("Active", "Open"))]
    wb = Workbook()
    ws = wb.active
    ws.title = "Open cases"
    write_sheet(ws, active)
    write_sheet(wb.create_sheet("All MHC cases"), rows)
    note = wb.create_sheet("Notes")
    for line in [
        f"Generated {datetime.now():%d/%m/%Y %H:%M} from {BASE}/high-court-search",
        f"Search: party contains '{PARTY}' (any role) -> {len(cases)} cases; {len(rows)} list Mason Hayes & Curran; {len(active)} open (no concluding order).",
        "courts.ie has no explicit open/closed flag. 'Concluded' = an order or listing result containing "
        "Final / Struck Out / Discontinued / Dismissed / Settled / Withdrawn with no later activity.",
        f"'Open – likely stayed or dormant' = no concluding order but no filings/orders/listings in the last {DORMANT_DAYS} days (often stayed, e.g. pending CJEU or appeal).",
        "'Against' = parties on the opposite side to the DPC (excluding notice parties, shown under Other parties).",
    ]:
        note.append([line])
    note.column_dimensions["A"].width = 140
    wb.save(out)

    print(f"\nWrote {out}  ({len(active)} open / {len(rows)} total)\n")
    for r in active:
        print(f"{r['Case reference']:<22} {r['Date of issue']}  DPC={r['DPC is']:<24} {r['Status (inferred)'][:40]:<41} vs {r['Against']}")


if __name__ == "__main__":
    main()
