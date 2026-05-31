"""
app.py  -  PCAOB Rotation Build Service
Kreit & Chiu CPA LLP

Endpoints:
  GET  /health   - liveness check
  POST /build    - runs full build, returns Excel + HTML + email alert as JSON

Environment variables (set in Railway - never in code):
  BIZINTA_SUBDOMAIN   e.g. "benjaminllp"
  BIZINTA_TOKEN       Bearer token from Bizinta API Access page
  FILTER_SERVICE_URL  URL of the PCAOB filter service /filter endpoint
  BUILD_API_KEY       Secret key - Power Automate sends this in X-Api-Key header

Changes vs previous version:
  Task 3  - EP start year column made visually prominent (bold, coloured) in Dashboard
  Task 4  - Clickable rows in HTML dashboard showing drill-down filing history
  Task 5  - EQR assumption note added to Excel Sheet 1 banner and HTML dashboard note panel
  Task 6  - Bizinta client status column added to All Filings tab
  Task 7  - Issuer ID + PCAOB name columns added to All Filings tab
  Task 8  - Partner Summary tab splits active vs departed partners
  Task 9  - Drill-down modal in HTML shows calculation logic per row
  Task 10 - Client Reconciliation tab added (filed-but-gone / active-but-no-filing)
  Task 11 - Name Changes tab added using PCAOB Issuer ID as primary key
  Task 12 - Gap-year caveat flag added for human review (flagged rows highlighted)
  Task 13 - Raw PCAOB Filings tab added
  Task 14 - Raw Bizinta Client Data tab added
"""

import os, io, csv, json, base64, urllib.request, urllib.error
from datetime import date, datetime as _dt
from collections import defaultdict
from functools import wraps
from flask import Flask, jsonify, request, abort
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# -- Environment
BIZINTA_SUBDOMAIN  = os.environ.get("BIZINTA_SUBDOMAIN", "")
BIZINTA_TOKEN      = os.environ.get("BIZINTA_TOKEN", "")
FILTER_SERVICE_URL = os.environ.get("FILTER_SERVICE_URL",
                     "https://pcaob-filter-service-production.up.railway.app/filter")
BUILD_API_KEY      = os.environ.get("BUILD_API_KEY", "")

# -- Auth decorator
def require_api_key(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if BUILD_API_KEY and request.headers.get("X-Api-Key") != BUILD_API_KEY:
            abort(401, "Invalid or missing API key")
        return f(*args, **kwargs)
    return decorated

# -- Style helpers (openpyxl)
F         = "Arial"
DARK_BLUE = "1F4E79"
MED_BLUE  = "2E75B6"
WHITE     = "FFFFFF"
ORANGE    = "FF8C00"
GREEN     = "70AD47"
PURPLE    = "7030A0"
TEAL      = "00B0F0"

def _side(c="CCCCCC"): return Side(style="thin", color=c)
def tbdr(c="CCCCCC"):  s=_side(c); return Border(left=s,right=s,top=s,bottom=s)
def hdr(cell, bg=DARK_BLUE, fg=WHITE, sz=10):
    cell.font      = Font(name=F, bold=True, color=fg, size=sz)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = tbdr("AAAAAA")
def val(cell, v, sz=10, bold=False, color="000000", ha="left", wrap=False, bg=None):
    cell.value     = v
    cell.font      = Font(name=F, size=sz, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    cell.border    = tbdr("DDDDDD")
    if bg: cell.fill = PatternFill("solid", fgColor=bg)
def sc_style(cell, sc, sfg):
    bgs = {5:"FF0000", 4:"FF8C00", 3:"FFD700", 2:"92D050", 1:"00B050"}
    cell.fill      = PatternFill("solid", fgColor=bgs[min(sc,5)])
    cell.font      = Font(name=F, bold=(sc>=4), color=sfg, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")

# -- Name helpers
def flip(name):
    if not name: return ""
    name = str(name).strip()
    if "," in name:
        p = name.split(",", 1)
        return f"{p[1].strip()} {p[0].strip()}"
    return name

EP_FIXES = {
    "Huang James":      "James Huang",
    "Benjmain Chung":   "Benjamin Chung",
    "Benjamin J Chung": "Benjamin Chung",
    "Chung Benjamin":   "Benjamin Chung",
}
def fx(n): return EP_FIXES.get(n.strip(), n.strip())

ISSUER_NORM = {
    "Muscle Maker, Inc.":                         "Sadot Group Inc.",
    "Sadot Group, Inc.":                          "Sadot Group Inc.",
    "Phaos Technology (Cayman) Holdings Limited": "Phaos Technology Holdings (Cayman) Ltd",
    "Maison Solutions Inc.":                      "Maison Solutions, Inc.",
    "Vyome Therapeutics Inc.":                    "Vyome Holdings, Inc.",
    "Vyome Holdings, Inc.":                       "Vyome Holdings, Inc.",
    "IMMRSIV Inc.":                               "Immrsiv Inc.",
    "Datasea, Inc.":                              "Datasea Inc.",
    "Sigyn Therapeutics, Inc.":                   "Sigyn Therapeutics, Inc",
    "VPR Brands, LP":                             "VPR Brands LP",
    "VPR Brands, LP.":                            "VPR Brands LP",
    "VPR Brands, L.P.":                           "VPR Brands LP",
    "Franklin Wireless Corp.":                    "Franklin Wireless Corp",
    "I-ON Digital Corp.":                         "I-ON Digital Corp",
    "Boomer Holdings, Inc":                       "Boomer Holdings Inc",
    "Thunder Energies Corporation":               "Thunder Energies Corp",
}
def fxi(n): n=n.replace("\xa0"," ").strip(); return ISSUER_NORM.get(n, n)

BIZINTA_TO_PCAOB = {
    "AlphaTime Acquisition Corp.":                       None,
    "Alternus Clean Energy":                             "Alternus Clean Energy, Inc.",
    "American Senior Association Holding Group, Inc.":   None,
    "Baird Medical":                                     "Baird Medical Investment Holdings Ltd",
    "Baiya":                                             "Baiya International Group Inc.",
    "Brava Acquisition Corp":                            "Brava Acquisition Corp",
    "Btab Ecommerce Group, Inc.":                        "Btab Ecommerce Group, Inc.",
    "Concorde International Group Ltd":                  "Concorde International Group Ltd.",
    "CRK":                                               None,
    "Cuprina":                                           "Cuprina Holdings (Cayman) LTD",
    "Curanex Pharmaceuticals":                           "Curanex Pharmaceuticals Inc",
    "Datasea":                                           "Datasea Inc.",
    "Durango Gold Corp":                                 None,
    "EDETEK Inc.":                                       None,
    "Followone Inc.":                                    None,
    "Fonon Corporation":                                 None,
    "FOXO Technologies Inc.":                            "FOXO Technologies Inc.",
    "FutureCrest Acquisition Corp. II":                  None,
    "FutureCrest Acquisition Corp III":                  None,
    "Galle Technology Limited":                          None,
    "GATC Health Corp.":                                 None,
    "Geoswift Digital Group Limited":                    None,
    "G-mango Inc.":                                      None,
    "Goodvision AI Inc.":                                None,
    "Graphjet Technology":                               "Graphjet Technology",
    "HDEDUCATION":                                       None,
    "Inspire Veterinary Partners":                       "Inspire Veterinary Partners, Inc.",
    "IntelliStem":                                       None,
    "I-ON Digital Corp":                                 "I-ON Digital Corp",
    "Iveda Solutions":                                   None,
    "IWAC Holding Company Inc.":                         "IWAC Holding Co Inc.",
    "iZooto":                                            None,
    "JTS - Outsourcing.":                                None,
    "Kandi Technologies Group, Inc.":                    "Kandi Technologies Group, Inc.",
    "Kat Fabricators":                                   None,
    "Knorex":                                            "Knorex Ltd.",
    "Laser Photonics Corporation":                       None,
    "LDR":                                               None,
    "Lithium & Boron Technology, Inc.":                  "Lithium & Boron Technology, Inc.",
    "Lokahi Therapeutics (fka Apimeds Pharmaceuticals)": "Apimeds Pharmaceuticals US, Inc.",
    "Maison Solutions, Inc.":                            "Maison Solutions, Inc.",
    "Medera Inc.":                                       "Medera Inc.",
    "Meihua International Medical Technologies Co., Ltd.":"Meihua International Medical Technologies Co., Ltd.",
    "Non-Client Projects":                               None,
    "NovelStem International Corp.":                     "NovelStem International Corp.",
    "Novusterra":                                        "Novusterra Inc.",
    "Olayan America":                                    None,
    "PCAOB":                                             None,
    "Phaos":                                             "Phaos Technology Holdings (Cayman) Ltd",
    "Porche Capital SEE 1 Acquisition Corp.":            None,
    "Precision Aerospace Group Inc":                     "Precision Aerospace & Defense Group, Inc.",
    "PT Transcoal":                                      None,
    "RainRock Acquisition Corp":                         None,
    "Restake Technologies Ltd.":                         None,
    "Ryde Technology":                                   "Ryde Group Ltd",
    "Sadot Group Inc.":                                  "Sadot Group Inc.",
    "SDE":                                               None,
    "Shanghai Yihang Network Technology Co., Ltd":       None,
    "Shreya Acquisition Group":                          "Shreya Acquisition Group",
    "Sigyn Therapeutic":                                 "Sigyn Therapeutics, Inc",
    "Skubbs Pte. Ltd.":                                  "Skubbs Holdings Ltd",
    "Test Company":                                      None,
    "TJGC Group Limited (fka CTRL Media)":               "Ctrl Group Limited",
    "TrueTribe":                                         None,
    "Unicoin Inc.":                                      "Unicoin Inc.",
    "VPR Brands LP":                                     "VPR Brands LP",
    "Vyome Holdings, Inc.":                              "Vyome Holdings, Inc.",
    "Zooming (Korea) Co., Ltd.":                         None,
}

# Reverse lookup: PCAOB name -> Bizinta name (for reconciliation)
PCAOB_TO_BIZINTA = {v: k for k, v in BIZINTA_TO_PCAOB.items() if v}

EQR_PROP_ID = "45173303"

# Ghost/deleted clients to exclude from rotation dashboard
GHOST_CLIENTS = {"Epione Health", "Fuse Group Holding Inc", "2022 Audit"}

# -- Bizinta API
def fetch_bizinta():
    endpoint = f"https://{BIZINTA_SUBDOMAIN}.bizinta.com/graphql"
    query = """
    {
      organizations(filters: {}) {
        nodes {
          displayName
          pipelineStatus { displayName }
          clientManager { displayName }
          genericPropValues {
            genericProp { id }
            genericPropValues { displayName }
          }
        }
      }
    }
    """
    payload = json.dumps({"query": query}).encode("utf-8")
    req = urllib.request.Request(endpoint, data=payload, headers={
        "Authorization":   f"Bearer {BIZINTA_TOKEN}",
        "Content-Type":    "application/json",
        "Accept":          "application/json",
        "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
        "Origin":          f"https://{BIZINTA_SUBDOMAIN}.bizinta.com",
        "Referer":         f"https://{BIZINTA_SUBDOMAIN}.bizinta.com/graphql/ide",
    }, method="POST")
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    clients = data["data"]["organizations"]["nodes"]
    out = {}
    for c in clients:
        nm     = c["displayName"]
        if nm in GHOST_CLIENTS:
            continue
        status = (c.get("pipelineStatus") or {}).get("displayName", "")
        ep     = flip((c.get("clientManager") or {}).get("displayName", ""))
        eqr    = ""
        for info in c.get("genericPropValues") or []:
            if (info.get("genericProp") or {}).get("id") == EQR_PROP_ID:
                vals = info.get("genericPropValues") or []
                if vals: eqr = vals[0].get("displayName", "") or ""
                break
        out[nm] = {"status": status, "ep": ep, "eqr": eqr}
    return out

# -- PCAOB data from filter service
def fetch_pcaob():
    with urllib.request.urlopen(FILTER_SERVICE_URL, timeout=60) as resp:
        raw = resp.read().decode("utf-8")
    reader = csv.DictReader(io.StringIO(raw))
    all_bulk = []
    raw_rows = []
    for row in reader:
        raw_rows.append(dict(row))
        if row.get("Latest Form AP Filing", "").strip() != "1": continue
        fn = row.get("Engagement Partner First Name", "").strip()
        mn = row.get("Engagement Partner Middle Name", "").strip()
        ln = row.get("Engagement Partner Last Name", "").strip()
        ep_full = fx((f"{fn} {mn} {ln}" if mn else f"{fn} {ln}").strip())
        fpe_raw = row.get("Fiscal Period End Date", "").strip().split()[0]
        try:
            fpe_dt  = _dt.strptime(fpe_raw, "%m/%d/%Y")
            fye_str = fpe_dt.strftime("%Y-%m-%d")
            yr      = fpe_dt.year
        except:
            fye_str = fpe_raw; yr = 0
        filed_raw = row.get("Filing Date", "").strip().split()[0]
        try:    filed = _dt.strptime(filed_raw, "%m/%d/%Y").strftime("%Y-%m-%d")
        except: filed = filed_raw
        signer = fx(f"{row.get('Signed First Name','').strip()} {row.get('Signed Last Name','').strip()}".strip())
        issuer_raw  = row.get("Issuer Name", "")
        issuer_norm = fxi(issuer_raw)
        issuer_id   = row.get("Issuer CIK", "").strip()
        all_bulk.append({
            "year": yr, "ep": ep_full,
            "issuer": issuer_norm,
            "issuer_raw": issuer_raw.strip(),
            "issuer_id": issuer_id,
            "signer": signer, "filed": filed, "fye": fye_str,
        })
    # Dedup: per (ep, issuer_norm, fye) keep most recently filed
    groups = defaultdict(list)
    for r in all_bulk:
        groups[(r["ep"], r["issuer"], r["fye"])].append(r)
    records = sorted(
        [max(g, key=lambda x: x["filed"]) for g in groups.values()],
        key=lambda r: (r["ep"], r["issuer"], r["year"])
    )
    return records, raw_rows

# -- Gap-year detection (Task 12)
def detect_gap_years(enriched):
    """
    Returns a set of (ep, issuer, year) tuples where a gap-year return pattern
    is detected: the partner has a filing this year but had NO filing in the
    immediately preceding year, yet DID have a filing before that gap.
    These rows need human review - a deliberate gap to reset the clock would be
    viewed as circumvention by PCAOB.
    """
    group_years = defaultdict(set)
    for r in enriched:
        group_years[(r["ep"], r["issuer"])].add(r["year"])

    gap_flags = set()
    for (ep, issuer), years in group_years.items():
        sorted_years = sorted(years)
        for i, yr in enumerate(sorted_years):
            if i == 0:
                continue
            prev = sorted_years[i-1]
            if yr - prev > 1:
                # Gap detected - all records AFTER the gap are flagged
                for flagged_yr in sorted_years[i:]:
                    gap_flags.add((ep, issuer, flagged_yr))
    return gap_flags

# -- Rotation logic
def build_rotation(records):
    group_years = defaultdict(set)
    for r in records: group_years[(r["ep"], r["issuer"])].add(r["year"])

    def consec(ep, issuer, yr):
        ys = group_years[(ep, issuer)]; c = 1; y = yr - 1
        while y in ys: c += 1; y -= 1
        return c

    def rot_status(c):
        if c >= 5: return (5, "CRITICAL - Year 5+ (Rotate Now)", "FF0000", "FFFFFF")
        if c == 4: return (4, "WARNING - Year 4 (Plan Rotation)", "FF8C00", "FFFFFF")
        if c == 3: return (3, "MONITOR - Year 3",                 "FFD700", "000000")
        if c == 2: return (2, "OK - Year 2",                      "92D050", "000000")
        return          (1, "OK - Year 1",                        "00B050", "FFFFFF")

    enriched = []
    for r in records:
        c = consec(r["ep"], r["issuer"], r["year"])
        sc, sl, sbg, sfg = rot_status(c)
        all_yrs = sorted(group_years[(r["ep"], r["issuer"])])
        max_yr  = max(all_yrs)
        start_yr = r["year"] - c + 1

        # Build calculation explanation (Task 9)
        consec_chain = []
        y = r["year"]
        while y in group_years[(r["ep"], r["issuer"])]:
            consec_chain.append(y)
            y -= 1
        consec_chain = sorted(consec_chain)

        calc_note = (
            f"Years on file: {', '.join(str(y) for y in all_yrs)}. "
            f"Consecutive chain ending {r['year']}: {' > '.join(str(y) for y in consec_chain)} = {c} year(s). "
            f"Rotation limit: 5 years. Years remaining: {max(0, 5 - c)}."
        )
        if len(all_yrs) > len(consec_chain):
            gap_yrs = []
            for i in range(1, len(all_yrs)):
                if all_yrs[i] - all_yrs[i-1] > 1:
                    gap_yrs.append(f"{all_yrs[i-1]}-{all_yrs[i]}")
            calc_note += f" NOTE: Gap(s) detected in filing history ({', '.join(gap_yrs)}) - count restarted after gap."

        enriched.append({**r, "consec":c, "sc":sc, "sl":sl, "sbg":sbg, "sfg":sfg,
                         "yrs_left":max(0,5-c), "max_yr":max_yr,
                         "start_yr":start_yr,
                         "all_yrs": all_yrs,
                         "calc_note": calc_note})

    latest_map = {}
    for r in enriched:
        k = (r["ep"], r["issuer"])
        if k not in latest_map or r["year"] > latest_map[k]["year"]:
            latest_map[k] = r
    dashboard = sorted(latest_map.values(), key=lambda r: (-r["sc"], r["ep"], r["issuer"]))
    return enriched, dashboard

# -- Excel builder
def build_excel(enriched, dashboard, pcaob_to_eqr, bizinta_data, raw_pcaob_rows, run_date_str):
    wb = Workbook()
    gap_flags = detect_gap_years(enriched)

    # ---- Sheet 1: Rotation Dashboard ----
    ws1 = wb.active; ws1.title = "Rotation Dashboard"
    ws1.sheet_properties.tabColor = DARK_BLUE
    ws1.merge_cells("A1:M1")
    ws1["A1"] = "KREIT & CHIU CPA LLP - PCAOB Partner Rotation Tracker (AS 1201)"
    ws1["A1"].font      = Font(name=F, bold=True, size=14, color=DARK_BLUE)
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32
    ws1.merge_cells("A2:M2")
    ws1["A2"] = (f"Source: PCAOB Form AP Filings (Firm 6651) | EQR: Bizinta API (live) | "
                 f"Generated: {run_date_str} | {len(enriched)} filings | {len(dashboard)} active engagements")
    ws1["A2"].font      = Font(name=F, italic=True, size=9, color="595959")
    ws1["A2"].alignment = Alignment(horizontal="center")
    ws1.row_dimensions[2].height = 15

    # Task 5: EQR assumption note
    ws1.merge_cells("A3:M3")
    ws1["A3"] = ("EQR column pre-filled from Bizinta live API. Green cell = EQR confirmed in Bizinta. Yellow = needs manual entry. "
                 "ASSUMPTION (B - CRITICAL): EQR tenure assumed to start same year as EP. PCAOB Form AP does not record EQR "
                 "appointment dates. Human review required where firm has internal records showing a different EQR start date.")
    ws1["A3"].font      = Font(name=F, italic=True, size=8, color="7F3F00")
    ws1["A3"].fill      = PatternFill("solid", fgColor="FFF2CC")
    ws1["A3"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws1.row_dimensions[3].height = 36

    # Task 12: Gap-year caveat note
    ws1.merge_cells("A4:M4")
    ws1["A4"] = ("GAP-YEAR ASSUMPTION (A - CRITICAL): A year with no Form AP filing breaks the consecutive chain and resets "
                 "the count to Year 1 on return (literal interpretation of SEC Rule 2-01(c)(6)). Rows marked [GAP] in orange "
                 "indicate a gap-then-return pattern - human review required. An engineered gap to reset the clock is "
                 "circumvention under PCAOB rules.")
    ws1["A4"].font      = Font(name=F, italic=True, size=8, color="7F0000")
    ws1["A4"].fill      = PatternFill("solid", fgColor="FFE4E1")
    ws1["A4"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws1.row_dimensions[4].height = 36

    # Task 3: Start Year is now column C (was implicit) - made visually prominent with bold blue
    DASH_HDRS = ["Engagement\nPartner", "Issuer / Client", "EP Start\nYear",
                 "Latest\nAudit Yr", "Consec.\nYears", "Yrs\nLeft",
                 "Rotation Status", "Signer\n(Form AP)", "EQR\n(Bizinta)",
                 "Fiscal\nYear End", "Last Filed", "Rotation\nDeadline",
                 "Gap\nFlag"]
    ws1.row_dimensions[5].height = 44
    for c, h in enumerate(DASH_HDRS, 1): hdr(ws1.cell(5, c, h))
    CW = [22, 36, 11, 12, 11, 10, 32, 22, 26, 14, 14, 16, 10]
    for i, w in enumerate(CW, 1): ws1.column_dimensions[get_column_letter(i)].width = w

    for ri, r in enumerate(dashboard, 6):
        eqr_val  = pcaob_to_eqr.get(r["issuer"], "")
        left_str = "ROTATE NOW" if r["yrs_left"] == 0 else str(r["yrs_left"])
        fye      = r["fye"][:10] if r["fye"] else ""
        deadline = str(r["start_yr"] + 4) if r["start_yr"] else ""
        has_gap  = any((r["ep"], r["issuer"], yr) in gap_flags for yr in r["all_yrs"])
        gap_flag = "[GAP]" if has_gap else ""

        row_vals = [r["ep"], r["issuer"], r["start_yr"], r["year"], r["consec"],
                    left_str, r["sl"], r["signer"], eqr_val, fye, r["filed"][:10],
                    deadline, gap_flag]
        for ci, v in enumerate(row_vals, 1):
            cell = ws1.cell(ri, ci, v)
            cell.font      = Font(name=F, size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=(ci in [2, 7]))
            cell.border    = tbdr("DDDDDD")
            if ci in (5, 7):
                sc_style(cell, r["sc"], r["sfg"])
            elif ci == 6:
                if r["yrs_left"] == 0:
                    cell.fill = PatternFill("solid", fgColor="FF0000")
                    cell.font = Font(name=F, bold=True, color="FFFFFF", size=10)
                elif r["yrs_left"] == 1:
                    cell.fill = PatternFill("solid", fgColor="FF8C00")
                    cell.font = Font(name=F, bold=True, color="FFFFFF", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif ci == 3:
                # Task 3: EP Start Year - bold, dark blue, centred, prominent
                cell.font  = Font(name=F, bold=True, size=11, color=DARK_BLUE)
                cell.fill  = PatternFill("solid", fgColor="DEEAF1")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif ci == 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif ci == 9:
                bg_eqr = "E2EFDA" if eqr_val else "FFF2CC"
                cell.fill = PatternFill("solid", fgColor=bg_eqr)
                if not eqr_val:
                    cell.font  = Font(name=F, size=9, italic=True, color="AA6600")
                    cell.value = "(enter EQR)"
            elif ci == 13 and has_gap:
                # Task 12: Gap flag column - orange highlight
                cell.fill  = PatternFill("solid", fgColor="FF8C00")
                cell.font  = Font(name=F, bold=True, color="FFFFFF", size=9)
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[ri].height = 18
    ws1.freeze_panes = "A6"

    # ---- Sheet 2: All Filings (Tasks 6, 7, 12) ----
    ws2 = wb.create_sheet("All Filings"); ws2.sheet_properties.tabColor = MED_BLUE
    ws2.merge_cells("A1:N1")
    ws2["A1"] = f"ALL FORM AP FILINGS - Kreit & Chiu CPA LLP | {len(enriched)} records"
    ws2["A1"].font      = Font(name=F, bold=True, size=12, color=DARK_BLUE)
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 24
    # Task 6: added Bizinta Status; Task 7: added Issuer ID and PCAOB Registered Name
    ALL_HDRS = ["Audit Year", "Filing Date", "Engagement Partner", "Issuer / Client",
                "Issuer ID\n(CIK)", "PCAOB Registered\nName", "Consec. Year #",
                "Start Year", "Rotation Status", "Signer (Form AP)",
                "EQR (Bizinta)", "Fiscal Year End", "Yrs Left",
                "Bizinta\nClient Status"]
    ws2.row_dimensions[2].height = 40
    for c, h in enumerate(ALL_HDRS, 1): hdr(ws2.cell(2, c, h), bg=MED_BLUE)
    CW2 = [10, 13, 22, 36, 12, 36, 13, 11, 30, 22, 24, 14, 10, 16]
    for i, w in enumerate(CW2, 1): ws2.column_dimensions[get_column_letter(i)].width = w

    for ri, r in enumerate(enriched, 3):
        eqr_v       = pcaob_to_eqr.get(r["issuer"], "")
        issuer_id   = r.get("issuer_id", "")
        issuer_raw  = r.get("issuer_raw", r["issuer"])
        biz_status  = ""
        # Look up Bizinta status via reverse mapping
        biz_name = PCAOB_TO_BIZINTA.get(r["issuer"])
        if biz_name and biz_name in bizinta_data:
            biz_status = bizinta_data[biz_name].get("status", "")
        has_gap = (r["ep"], r["issuer"], r["year"]) in gap_flags

        rv = [r["year"], r["filed"][:10], r["ep"], r["issuer"],
              issuer_id, issuer_raw, r["consec"],
              r["start_yr"], r["sl"], r["signer"], eqr_v,
              r["fye"][:10] if r["fye"] else "", r["yrs_left"], biz_status]
        for ci, v in enumerate(rv, 1):
            cell = ws2.cell(ri, ci, v)
            cell.font      = Font(name=F, size=9)
            cell.alignment = Alignment(vertical="center", wrap_text=(ci in [4, 6, 9]))
            cell.border    = tbdr("EEEEEE")
            if ci in (7, 9):
                sc_style(cell, r["sc"], r["sfg"])
            elif ci == 11 and eqr_v:
                cell.fill = PatternFill("solid", fgColor="E2EFDA")
            elif ci == 14:
                if biz_status == "Active":
                    cell.fill = PatternFill("solid", fgColor="E2EFDA")
                    cell.font = Font(name=F, size=9, color="375623")
                elif biz_status == "Inactive":
                    cell.fill = PatternFill("solid", fgColor="FFF2CC")
                    cell.font = Font(name=F, size=9, color="7F3F00")
            # Task 12: highlight gap-flagged rows with light orange background
            if has_gap and ci not in (7, 9, 11, 14):
                current_fill = cell.fill.fgColor.rgb if cell.fill.patternType == "solid" else None
                if not current_fill or current_fill == "00000000":
                    cell.fill = PatternFill("solid", fgColor="FFF0E0")
        ws2.row_dimensions[ri].height = 16
    ws2.freeze_panes = "A3"

    # ---- Sheet 3: Partner Summary (Task 8: active vs departed) ----
    ws3 = wb.create_sheet("Partner Summary"); ws3.sheet_properties.tabColor = "70AD47"
    ws3.merge_cells("A1:I1")
    ws3["A1"] = "ENGAGEMENT PARTNER ACTIVITY SUMMARY"
    ws3["A1"].font = Font(name=F, bold=True, size=13, color=DARK_BLUE)
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    # Task 8: determine departed vs active partners from Bizinta employee data
    # We use the dashboard EP list - if they appear in bizinta as Inactive, mark departed
    # For now we use a status derived from whether the EP has ANY active client in Bizinta
    active_eps   = set()
    departed_eps = set()
    for biz_name, biz_info in bizinta_data.items():
        ep = biz_info.get("ep", "")
        if not ep: continue
        if biz_info.get("status") == "Active":
            active_eps.add(ep)
        else:
            departed_eps.add(ep)
    # An EP with any active client is active
    all_dashboard_eps = set(r["ep"] for r in dashboard)

    pdata = defaultdict(lambda: {"engagements": set(), "years": set(), "max_c": 0,
                                  "crit": 0, "warn": 0, "mon": 0, "is_active": False})
    for r in dashboard:
        p = pdata[r["ep"]]
        p["engagements"].add(r["issuer"])
        p["years"].add(r["year"])
        p["max_c"] = max(p["max_c"], r["consec"])
        if r["sc"] >= 5: p["crit"] += 1
        elif r["sc"] == 4: p["warn"] += 1
        elif r["sc"] == 3: p["mon"] += 1
        if r["ep"] in active_eps: p["is_active"] = True

    # Task 8: Section A - Active partners
    ws3.merge_cells("A2:I2")
    ws3["A2"] = "SECTION A: ACTIVE PARTNERS (have active clients in Bizinta)"
    ws3["A2"].font = Font(name=F, bold=True, size=11, color=WHITE)
    ws3["A2"].fill = PatternFill("solid", fgColor="375623")
    ws3["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[2].height = 22

    SUM_HDRS = ["Engagement Partner", "Status", "Active Clients", "Years Active\n(Range)",
                "Max Consec.\nYears", "Critical\n(Yr 5+)", "Warning\n(Yr 4)",
                "Monitor\n(Yr 3)", "Overall Risk"]
    ws3.row_dimensions[3].height = 40
    for c, h in enumerate(SUM_HDRS, 1): hdr(ws3.cell(3, c, h))
    CW3 = [24, 12, 14, 16, 16, 14, 14, 14, 16]
    for i, w in enumerate(CW3, 1): ws3.column_dimensions[get_column_letter(i)].width = w

    active_partners   = [(ep, pd) for ep, pd in pdata.items() if pd["is_active"]]
    inactive_partners = [(ep, pd) for ep, pd in pdata.items() if not pd["is_active"]]

    def write_partner_row(ws, ri, ep, pd, status_label, status_bg, status_fg):
        yrs  = sorted(pd["years"])
        risk = ("HIGH RISK", "FF0000", "FFFFFF") if pd["crit"] else \
               ("MEDIUM RISK", "FF8C00", "FFFFFF") if pd["warn"] else \
               ("MONITOR", "FFD700", "000000") if pd["mon"] else ("LOW RISK", "00B050", "FFFFFF")
        rv = [ep, status_label, len(pd["engagements"]),
              f"{yrs[0]}-{yrs[-1]}" if yrs else "",
              pd["max_c"], pd["crit"], pd["warn"], pd["mon"], risk[0]]
        for ci, v in enumerate(rv, 1):
            cell = ws.cell(ri, ci, v)
            cell.font      = Font(name=F, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = tbdr()
            if ci == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif ci == 2:
                cell.fill = PatternFill("solid", fgColor=status_bg)
                cell.font = Font(name=F, bold=True, color=status_fg, size=9)
            elif ci == 9:
                cell.fill = PatternFill("solid", fgColor=risk[1])
                cell.font = Font(name=F, bold=True, color=risk[2], size=10)
            elif ci == 5:
                mc = pd["max_c"]; fg = "FFFFFF" if mc >= 3 else "000000"
                sc_style(cell, min(mc, 5), fg)
        ws.row_dimensions[ri].height = 20

    current_row = 4
    for ep, pd in sorted(active_partners, key=lambda x: (-x[1]["crit"], -x[1]["warn"], x[0])):
        write_partner_row(ws3, current_row, ep, pd, "Active", "375623", "FFFFFF")
        current_row += 1

    # Task 8: Section B - Departed/Inactive partners
    current_row += 1
    ws3.merge_cells(f"A{current_row}:I{current_row}")
    ws3[f"A{current_row}"] = "SECTION B: DEPARTED / INACTIVE PARTNERS (no active clients in Bizinta)"
    ws3[f"A{current_row}"].font = Font(name=F, bold=True, size=11, color=WHITE)
    ws3[f"A{current_row}"].fill = PatternFill("solid", fgColor="595959")
    ws3[f"A{current_row}"].alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[current_row].height = 22
    current_row += 1

    hdr_row = current_row
    for c, h in enumerate(SUM_HDRS, 1): hdr(ws3.cell(hdr_row, c, h), bg="595959")
    ws3.row_dimensions[hdr_row].height = 40
    current_row += 1

    for ep, pd in sorted(inactive_partners, key=lambda x: (-x[1]["crit"], -x[1]["warn"], x[0])):
        write_partner_row(ws3, current_row, ep, pd, "Departed", "595959", "FFFFFF")
        current_row += 1

    # ---- Sheet 4: Client Reconciliation (Task 10) ----
    ws4 = wb.create_sheet("Client Reconciliation"); ws4.sheet_properties.tabColor = PURPLE

    ws4.merge_cells("A1:E1")
    ws4["A1"] = "CLIENT RECONCILIATION - Filed-but-Gone vs Active-but-No-Filing"
    ws4["A1"].font = Font(name=F, bold=True, size=13, color=DARK_BLUE)
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 28

    # Build sets for reconciliation
    pcaob_issuers  = set(r["issuer"] for r in enriched)

    # Category 1: in PCAOB filings but not active in Bizinta
    # A PCAOB issuer is "covered" if it maps to an Active Bizinta client
    active_biz_pcaob_names = set()
    for biz_name, biz_info in bizinta_data.items():
        if biz_info.get("status") == "Active":
            mapped = BIZINTA_TO_PCAOB.get(biz_name)
            if mapped:
                active_biz_pcaob_names.add(mapped)
    filed_but_gone = sorted(pcaob_issuers - active_biz_pcaob_names)

    # Category 2: Active in Bizinta but no PCAOB filing found
    # Include ALL active Bizinta clients where either:
    #   (a) their BIZINTA_TO_PCAOB mapping is None (not an issuer / no mapping defined), OR
    #   (b) their mapped PCAOB name has no filing in enriched
    # Exclude ghost clients and internal non-clients (None values already handled)
    active_no_filing = []
    for biz_name, biz_info in sorted(bizinta_data.items()):
        if biz_info.get("status") != "Active":
            continue
        mapped_pcaob = BIZINTA_TO_PCAOB.get(biz_name)
        # If mapped and the PCAOB name HAS filings, skip (this client is covered)
        if mapped_pcaob and mapped_pcaob in pcaob_issuers:
            continue
        # Otherwise: either no mapping, mapping=None, or mapped but no filing
        active_no_filing.append({
            "biz_name":    biz_name,
            "pcaob_name":  mapped_pcaob or "",
            "ep":          biz_info.get("ep", ""),
            "eqr":         biz_info.get("eqr", ""),
            "reason":      "No PCAOB mapping defined" if mapped_pcaob is None or biz_name not in BIZINTA_TO_PCAOB
                           else ("Mapped but no filing found" if mapped_pcaob else "No PCAOB mapping defined"),
        })

    # Category 1: Filed with PCAOB but not active in Bizinta
    ws4.merge_cells("A2:E2")
    ws4["A2"] = f"CATEGORY 1: Appears in PCAOB filings but NOT active in Bizinta ({len(filed_but_gone)} issuers)"
    ws4["A2"].font = Font(name=F, bold=True, size=11, color=WHITE)
    ws4["A2"].fill = PatternFill("solid", fgColor=DARK_BLUE)
    ws4["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws4.row_dimensions[2].height = 22

    hdrs4a = ["Issuer (PCAOB Name)", "Bizinta Name", "Bizinta Status", "Last EP", "Notes"]
    ws4.row_dimensions[3].height = 32
    for c, h in enumerate(hdrs4a, 1): hdr(ws4.cell(3, c, h))
    CW4 = [40, 40, 16, 22, 30]
    for i, w in enumerate(CW4, 1): ws4.column_dimensions[get_column_letter(i)].width = w

    ri = 4
    for issuer in filed_but_gone:
        biz_name = PCAOB_TO_BIZINTA.get(issuer, "")
        biz_status = ""
        if biz_name and biz_name in bizinta_data:
            biz_status = bizinta_data[biz_name].get("status", "")
        last_ep = ""
        last_yr = 0
        for r in enriched:
            if r["issuer"] == issuer and r["year"] > last_yr:
                last_yr = r["year"]; last_ep = r["ep"]
        note = "No Bizinta mapping" if not biz_name else ("Inactive in Bizinta" if biz_status == "Inactive" else "Check mapping")
        for ci, v in enumerate([issuer, biz_name, biz_status, last_ep, note], 1):
            cell = ws4.cell(ri, ci, v)
            cell.font = Font(name=F, size=10)
            cell.border = tbdr("DDDDDD")
            cell.alignment = Alignment(vertical="center", wrap_text=(ci in [1, 2, 5]))
        ws4.row_dimensions[ri].height = 16
        ri += 1

    # Category 2: Active in Bizinta but no PCAOB filing
    ri += 1
    ws4.merge_cells(f"A{ri}:E{ri}")
    ws4[f"A{ri}"] = f"CATEGORY 2: Active in Bizinta but NO PCAOB filing found ({len(active_no_filing)} clients)"
    ws4[f"A{ri}"].font = Font(name=F, bold=True, size=11, color=WHITE)
    ws4[f"A{ri}"].fill = PatternFill("solid", fgColor=MED_BLUE)
    ws4[f"A{ri}"].alignment = Alignment(horizontal="left", vertical="center")
    ws4.row_dimensions[ri].height = 22
    ri += 1

    hdrs4b = ["Bizinta Client Name", "PCAOB Mapped Name", "Bizinta EP", "Bizinta EQR", "Notes"]
    ws4.row_dimensions[ri].height = 32
    for c, h in enumerate(hdrs4b, 1): hdr(ws4.cell(ri, c, h), bg=MED_BLUE)
    ri += 1

    for item in active_no_filing:
        for ci, v in enumerate([item["biz_name"], item["pcaob_name"],
                                 item["ep"], item["eqr"], item["reason"]], 1):
            cell = ws4.cell(ri, ci, v)
            cell.font = Font(name=F, size=10)
            cell.border = tbdr("DDDDDD")
            cell.alignment = Alignment(vertical="center", wrap_text=(ci in [1, 2, 5]))
            cell.fill = PatternFill("solid", fgColor="EEF5FB")
        ws4.row_dimensions[ri].height = 16
        ri += 1

    # ---- Sheet 5: Name Changes (Task 11) ----
    ws5 = wb.create_sheet("Name Changes"); ws5.sheet_properties.tabColor = TEAL

    ws5.merge_cells("A1:D1")
    ws5["A1"] = "NAME CHANGES & NORMALISATION - PCAOB Issuer ID as Primary Key"
    ws5["A1"].font = Font(name=F, bold=True, size=13, color=DARK_BLUE)
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 28

    ws5.merge_cells("A2:D2")
    ws5["A2"] = ("This tab lists issuers where the PCAOB-registered name differs from the normalised name used in this tracker, "
                 "or where multiple raw names map to one canonical name. Issuer CIK is the authoritative key per PCAOB records.")
    ws5["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    ws5["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws5.row_dimensions[2].height = 28

    NC_HDRS = ["Issuer CIK\n(PCAOB ID)", "Raw PCAOB Name\n(as filed)", "Normalised Name\n(used in tracker)", "Mapped Bizinta Name"]
    ws5.row_dimensions[3].height = 40
    for c, h in enumerate(NC_HDRS, 1): hdr(ws5.cell(3, c, h), bg=TEAL)
    CW5 = [16, 46, 46, 40]
    for i, w in enumerate(CW5, 1): ws5.column_dimensions[get_column_letter(i)].width = w

    # Build name change data: group raw_pcaob_rows by issuer_id
    id_to_names = defaultdict(set)
    id_to_norm  = {}
    for r in enriched:
        iid = r.get("issuer_id", "")
        if iid:
            id_to_names[iid].add(r.get("issuer_raw", r["issuer"]))
            id_to_norm[iid]  = r["issuer"]

    nc_ri = 4
    shown = set()
    for r in enriched:
        iid      = r.get("issuer_id", "")
        raw_name = r.get("issuer_raw", r["issuer"])
        norm     = r["issuer"]
        if raw_name == norm and iid not in id_to_names:
            continue
        key = (iid, raw_name)
        if key in shown: continue
        shown.add(key)
        # Only show rows where there IS a normalisation or multiple names
        biz_name = PCAOB_TO_BIZINTA.get(norm, "")
        for ci, v in enumerate([iid, raw_name, norm, biz_name], 1):
            cell = ws5.cell(nc_ri, ci, v)
            cell.font = Font(name=F, size=10)
            cell.border = tbdr("DDDDDD")
            cell.alignment = Alignment(vertical="center", wrap_text=(ci in [2, 3, 4]))
            if raw_name != norm:
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
        ws5.row_dimensions[nc_ri].height = 16
        nc_ri += 1

    # Also show ISSUER_NORM entries explicitly
    ws5.cell(nc_ri, 1, "").border = tbdr()
    nc_ri += 1
    ws5.merge_cells(f"A{nc_ri}:D{nc_ri}")
    ws5[f"A{nc_ri}"] = "STATIC NORMALISATION DICTIONARY (ISSUER_NORM) - corporate rebrands and variant spellings"
    ws5[f"A{nc_ri}"].font = Font(name=F, bold=True, size=10, color=WHITE)
    ws5[f"A{nc_ri}"].fill = PatternFill("solid", fgColor=TEAL)
    ws5[f"A{nc_ri}"].alignment = Alignment(horizontal="left", vertical="center")
    ws5.row_dimensions[nc_ri].height = 20
    nc_ri += 1
    for raw_key, canonical in ISSUER_NORM.items():
        if raw_key == canonical: continue
        for ci, v in enumerate(["", raw_key, canonical, ""], 1):
            cell = ws5.cell(nc_ri, ci, v)
            cell.font = Font(name=F, size=10)
            cell.border = tbdr("DDDDDD")
            cell.alignment = Alignment(vertical="center")
            if v: cell.fill = PatternFill("solid", fgColor="E8F5E9")
        ws5.row_dimensions[nc_ri].height = 16
        nc_ri += 1

    # ---- Sheet 6: Legend & Notes ----
    ws6 = wb.create_sheet("Legend & Notes"); ws6.sheet_properties.tabColor = "595959"
    notes = [
        ("PCAOB AS 1201 - 5 Year Rule",
         "Both the Lead Engagement Partner (EP) and Engagement Quality Reviewer (EQR) must rotate after "
         "5 consecutive years of service on an audit engagement. A 5-year cooling-off period applies after mandatory rotation."),
        ("Consecutive Year Count",
         "Counted from Form AP filings. Each fiscal year end where a Form AP was filed counts as one year. "
         "The count resets if there is a gap year with no filing (Assumption A)."),
        ("EQR Assumption (B - CRITICAL)",
         "PCAOB Form AP does not record EQR appointment dates. This tracker assumes EQR tenure starts at the "
         "same time as the EP. Human review is required where the firm has internal records showing a different EQR start date."),
        ("Gap-Year Assumption (A - CRITICAL)",
         "A year with no Form AP filing breaks the consecutive chain and resets the count to Year 1 on return. "
         "This is the literal interpretation of 'consecutive' under SEC Rule 2-01(c)(6). However, an engineered gap "
         "(deliberately rotating one year to reset the clock) would be viewed as circumvention by PCAOB. "
         "Human review is required for any gap-then-return pattern. Rows marked [GAP] require review."),
        ("Cooling-Off Period",
         "The tracker monitors forward service (years toward the 5-year limit) only. It does not track whether "
         "a rotated partner has completed the 5-year cooling-off period before returning. Phase 2 item."),
        ("Ghost / Deleted Clients",
         "The following clients exist in the Bizinta API database but cannot be found in the Bizinta UI and are "
         "excluded from the tracker: Epione Health, Fuse Group Holding Inc, 2022 Audit."),
        ("Data Sources",
         "PCAOB: FirmFilings.csv bulk dataset (official PCAOB source, Firm ID 6651, Latest=1 filter). "
         "Bizinta: Live GraphQL API call at time of generation. EQR custom field Property ID: 45173303."),
        ("Deduplication Key",
         "Per (EP, issuer_normalised, fiscal_period_end_date). This correctly handles issuers with multiple "
         "fiscal year ends in the same calendar year (e.g. Brava Acquisition Corp 2025: Sep-30 AND Dec-31)."),
        ("Status Colour Coding",
         "Red = CRITICAL Year 5+ (rotate immediately). Orange = WARNING Year 4 (plan rotation). "
         "Yellow = MONITOR Year 3. Light green = OK Year 2. Green = OK Year 1."),
    ]
    ws6.merge_cells("A1:C1")
    ws6["A1"] = "LEGEND, ASSUMPTIONS & METHODOLOGY NOTES"
    ws6["A1"].font = Font(name=F, bold=True, size=13, color=DARK_BLUE)
    ws6["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws6.row_dimensions[1].height = 28
    ws6.column_dimensions["A"].width = 36
    ws6.column_dimensions["B"].width = 100
    ws6.column_dimensions["C"].width = 20
    for ri, (title, body) in enumerate(notes, 2):
        c1 = ws6.cell(ri, 1, title)
        c1.font = Font(name=F, bold=True, size=10, color=DARK_BLUE)
        c1.alignment = Alignment(vertical="top", wrap_text=True)
        c1.fill = PatternFill("solid", fgColor="DEEAF1")
        c1.border = tbdr("CCCCCC")
        c2 = ws6.cell(ri, 2, body)
        c2.font = Font(name=F, size=10)
        c2.alignment = Alignment(vertical="top", wrap_text=True)
        c2.border = tbdr("CCCCCC")
        ws6.row_dimensions[ri].height = 52

    # ---- Sheet 7: Raw PCAOB Filings (Task 13) ----
    ws7 = wb.create_sheet("Raw PCAOB Filings"); ws7.sheet_properties.tabColor = "C0C0C0"
    ws7.merge_cells("A1:J1")
    ws7["A1"] = f"RAW PCAOB FORM AP FILINGS - All rows from FirmFilings.csv for Firm 6651 | {len(raw_pcaob_rows)} total rows"
    ws7["A1"].font = Font(name=F, bold=True, size=11, color=DARK_BLUE)
    ws7["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws7.row_dimensions[1].height = 22

    if raw_pcaob_rows:
        raw_headers = list(raw_pcaob_rows[0].keys())
        ws7.row_dimensions[2].height = 32
        for c, h in enumerate(raw_headers, 1):
            hdr(ws7.cell(2, c, h), bg="595959")
            ws7.column_dimensions[get_column_letter(c)].width = max(12, min(40, len(h) + 4))
        for ri, row in enumerate(raw_pcaob_rows, 3):
            for ci, h in enumerate(raw_headers, 1):
                cell = ws7.cell(ri, ci, row.get(h, ""))
                cell.font = Font(name=F, size=9)
                cell.border = tbdr("EEEEEE")
                cell.alignment = Alignment(vertical="center")
            ws7.row_dimensions[ri].height = 14
    ws7.freeze_panes = "A3"

    # ---- Sheet 8: Raw Bizinta Client Data (Task 14) ----
    ws8 = wb.create_sheet("Raw Bizinta Data"); ws8.sheet_properties.tabColor = "FFC000"
    ws8.merge_cells("A1:E1")
    ws8["A1"] = f"RAW BIZINTA CLIENT DATA - Live API pull | {len(bizinta_data)} clients"
    ws8["A1"].font = Font(name=F, bold=True, size=11, color=DARK_BLUE)
    ws8["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws8.row_dimensions[1].height = 22

    BIZ_HDRS = ["Bizinta Display Name", "Status", "EP (Client Manager)", "EQR (Custom Field)", "PCAOB Mapped Name"]
    ws8.row_dimensions[2].height = 32
    for c, h in enumerate(BIZ_HDRS, 1): hdr(ws8.cell(2, c, h), bg="7F6000")
    CW8 = [36, 14, 24, 24, 40]
    for i, w in enumerate(CW8, 1): ws8.column_dimensions[get_column_letter(i)].width = w

    for ri, (biz_name, info) in enumerate(sorted(bizinta_data.items()), 3):
        pcaob_mapped = BIZINTA_TO_PCAOB.get(biz_name, "")
        rv = [biz_name, info.get("status",""), info.get("ep",""), info.get("eqr",""),
              pcaob_mapped if pcaob_mapped else "(no mapping)"]
        for ci, v in enumerate(rv, 1):
            cell = ws8.cell(ri, ci, v)
            cell.font = Font(name=F, size=10)
            cell.border = tbdr("DDDDDD")
            cell.alignment = Alignment(vertical="center", wrap_text=(ci in [1, 5]))
            if ci == 2:
                if v == "Active":
                    cell.fill = PatternFill("solid", fgColor="E2EFDA")
                    cell.font = Font(name=F, size=10, color="375623")
                elif v == "Inactive":
                    cell.fill = PatternFill("solid", fgColor="FFF2CC")
                    cell.font = Font(name=F, size=10, color="7F3F00")
            elif ci == 5 and not pcaob_mapped:
                cell.font = Font(name=F, size=10, italic=True, color="888888")
        ws8.row_dimensions[ri].height = 16
    ws8.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# -- HTML dashboard builder (Tasks 4, 5, 9, 12)
def build_html(enriched, dashboard, pcaob_to_eqr, run_date_str):
    gap_flags = detect_gap_years(enriched)

    def jsd(o):
        s = json.dumps(o, ensure_ascii=False)
        s = s.replace("&", "\\u0026").replace("<", "\\u003c").replace(">", "\\u003e")
        return s

    dash_js = jsd([{
        "ep":     r["ep"],
        "issuer": r["issuer"],
        "startYr":r["start_yr"],
        "yr":     r["year"],
        "consec": r["consec"],
        "left":   r["yrs_left"],
        "sc":     r["sc"],
        "sl":     r["sl"],
        "signer": r["signer"],
        "filed":  r["filed"][:10],
        "fye":    r["fye"][:10] if r["fye"] else "",
        "eqr":    pcaob_to_eqr.get(r["issuer"], ""),
        "allYrs": r["all_yrs"],
        "calcNote": r["calc_note"],
        "hasGap": any((r["ep"], r["issuer"], yr) in gap_flags for yr in r["all_yrs"]),
    } for r in dashboard])

    all_js = jsd([{
        "ep":      r["ep"],
        "issuer":  r["issuer"],
        "startYr": r["start_yr"],
        "yr":      r["year"],
        "consec":  r["consec"],
        "left":    r["yrs_left"],
        "sc":      r["sc"],
        "sl":      r["sl"],
        "signer":  r["signer"],
        "filed":   r["filed"][:10],
        "eqr":     pcaob_to_eqr.get(r["issuer"], ""),
        "calcNote": r["calc_note"],
        "hasGap":  (r["ep"], r["issuer"], r["year"]) in gap_flags,
    } for r in enriched])

    eps    = sorted(set(r["ep"] for r in dashboard))
    yrs    = sorted(set(r["year"] for r in enriched))
    ep_opts = "".join(f'<option value="{e}">{e}</option>' for e in eps)
    yr_opts = "".join(f'<option value="{y}">{y}</option>' for y in yrs)

    html = f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PCAOB Rotation Dashboard - Kreit &amp; Chiu CPA LLP</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:Arial,sans-serif;background:#f0f4f8;color:#1a1a2e;font-size:13px}}
.hdr{{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:#fff;padding:18px 28px;display:flex;justify-content:space-between;align-items:center}}
.hdr h1{{font-size:18px;font-weight:700}}.hdr .sub{{font-size:11px;opacity:.85;margin-top:3px}}
.hdr .badge{{background:rgba(255,255,255,.18);border-radius:8px;padding:8px 16px;text-align:center}}
.hdr .badge .n{{font-size:20px;font-weight:700}}
.stats{{display:flex;gap:12px;padding:14px 24px;background:#fff;border-bottom:1px solid #dde3ea;flex-wrap:wrap}}
.sc{{flex:1;min-width:110px;border-radius:8px;padding:12px 14px;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.07)}}
.sc .n{{font-size:26px;font-weight:700}}.sc .l{{font-size:10px;text-transform:uppercase;letter-spacing:.5px;margin-top:3px;opacity:.75}}
.s5c{{background:#fff0f0;color:#cc0000;border-left:4px solid #FF0000}}
.s4c{{background:#fff4e5;color:#cc5500;border-left:4px solid #FF8C00}}
.s3c{{background:#fffbe5;color:#806600;border-left:4px solid #FFD700}}
.okc{{background:#f0fff4;color:#006600;border-left:4px solid #00B050}}
.ttc{{background:#f0f4ff;color:#1F4E79;border-left:4px solid #1F4E79}}
.ctrl{{padding:11px 24px;background:#fff;border-bottom:1px solid #dde3ea;display:flex;gap:12px;align-items:center;flex-wrap:wrap}}
.ctrl label{{font-size:11px;font-weight:700;color:#555}}
.ctrl select,.ctrl input{{border:1px solid #ccc;border-radius:4px;padding:5px 10px;font-size:12px;background:#fff}}
.tabs{{margin-left:auto;display:flex;gap:6px}}
.tb{{padding:6px 14px;border-radius:4px;border:1px solid #ccc;font-size:12px;cursor:pointer;background:#fff;font-weight:500}}
.tb.active{{background:#1F4E79;color:#fff;border-color:#1F4E79}}
.main{{padding:14px 24px}}
.note{{background:#e8f4fd;border-left:4px solid #2E75B6;border-radius:0 6px 6px 0;padding:10px 16px;margin-bottom:10px;font-size:11px;color:#1a3a5c;line-height:1.6}}
.note-warn{{background:#fff4e5;border-left:4px solid #FF8C00;border-radius:0 6px 6px 0;padding:10px 16px;margin-bottom:10px;font-size:11px;color:#7F3F00;line-height:1.6}}
.note-danger{{background:#fff0f0;border-left:4px solid #FF0000;border-radius:0 6px 6px 0;padding:10px 16px;margin-bottom:14px;font-size:11px;color:#7F0000;line-height:1.6}}
.leg{{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:12px;padding:9px 14px;background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.li{{display:flex;align-items:center;gap:6px;font-size:11px}}
.ld{{width:13px;height:13px;border-radius:3px;flex-shrink:0}}
.stit{{font-size:13px;font-weight:700;color:#1F4E79;margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid #2E75B6;display:flex;align-items:center;gap:8px}}
.cnt{{background:#2E75B6;color:#fff;font-size:10px;padding:2px 7px;border-radius:10px}}
table{{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.08);margin-bottom:18px}}
th{{background:#1F4E79;color:#fff;padding:10px 12px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;white-space:nowrap}}
td{{padding:9px 12px;border-bottom:1px solid #edf0f5;font-size:12px;vertical-align:middle}}
tr:last-child td{{border-bottom:none}}tr:hover td{{background:#f7f9fc}}
.pill{{display:inline-block;padding:3px 10px;border-radius:12px;font-size:10px;font-weight:700;white-space:nowrap}}
.p5{{background:#FF0000;color:#fff}}.p4{{background:#FF8C00;color:#fff}}
.p3{{background:#FFD700;color:#333}}.p2{{background:#92D050;color:#333}}.p1{{background:#00B050;color:#fff}}
.bar{{display:flex;gap:3px;align-items:center}}
.bk{{width:17px;height:17px;border-radius:3px}}
.f5{{background:#FF0000}}.f4{{background:#FF8C00}}.f3{{background:#FFD700}}
.f2{{background:#92D050}}.f1{{background:#00B050}}.fe{{background:#e0e0e0}}
.eqr-biz{{display:inline-block;background:#E2EFDA;border:1px solid #70AD47;border-radius:4px;padding:3px 8px;font-size:11px;color:#375623;font-weight:600}}
.hidden{{display:none}}.nr{{text-align:center;padding:36px;color:#888;font-style:italic}}
.gen{{font-size:10px;color:#888;padding:8px 24px;text-align:right}}
/* Task 3: Start Year prominent styling */
.start-yr{{font-size:13px;font-weight:700;color:#1F4E79;background:#DEEAF1;padding:2px 7px;border-radius:4px;display:inline-block}}
/* Task 4 & 9: Clickable rows + drill-down modal */
.clickable{{cursor:pointer}}
.clickable:hover td{{background:#EBF5FB !important}}
/* Task 12: Gap flag styling */
.gap-row td{{background:#FFF0E0 !important}}
.gap-badge{{display:inline-block;background:#FF8C00;color:#fff;font-size:9px;font-weight:700;padding:1px 5px;border-radius:3px;margin-left:4px}}
/* Modal */
.modal-overlay{{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:1000;align-items:center;justify-content:center}}
.modal-overlay.open{{display:flex}}
.modal{{background:#fff;border-radius:10px;padding:28px;max-width:680px;width:90%;max-height:80vh;overflow-y:auto;box-shadow:0 8px 32px rgba(0,0,0,.25)}}
.modal-hdr{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:18px;padding-bottom:14px;border-bottom:2px solid #1F4E79}}
.modal-hdr h2{{font-size:15px;color:#1F4E79;font-weight:700}}
.modal-close{{background:none;border:none;font-size:20px;cursor:pointer;color:#888;line-height:1;padding:2px 6px}}
.modal-close:hover{{color:#333}}
.modal-section{{margin-bottom:16px}}
.modal-section h3{{font-size:12px;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}}
.modal-grid{{display:grid;grid-template-columns:140px 1fr;gap:6px 12px;font-size:12px}}
.modal-label{{color:#888;font-weight:600}}
.modal-value{{color:#1a1a2e}}
.calc-box{{background:#f8f9fa;border-left:3px solid #2E75B6;padding:10px 14px;border-radius:0 6px 6px 0;font-size:11px;color:#1a3a5c;line-height:1.7;margin-top:8px}}
.hist-table{{width:100%;border-collapse:collapse;font-size:11px;margin-top:8px}}
.hist-table th{{background:#f0f4f8;color:#555;padding:6px 10px;text-align:left;font-weight:700;border-bottom:2px solid #dde3ea}}
.hist-table td{{padding:5px 10px;border-bottom:1px solid #f0f0f0;vertical-align:middle}}
.gap-warn{{background:#fff4e5;border:1px solid #FF8C00;border-radius:6px;padding:10px 14px;font-size:11px;color:#7F3F00;margin-top:10px}}
</style></head><body>
<div class="hdr">
  <div><h1>PCAOB Partner Rotation Dashboard</h1>
  <div class="sub">Kreit &amp; Chiu CPA LLP &nbsp;|&nbsp; Firm ID: 6651 &nbsp;|&nbsp; PCAOB AS 1201 - 5-Year Rule &nbsp;|&nbsp; Generated: {run_date_str}</div></div>
  <div class="badge"><div class="n" id="hdr-date"></div><div style="font-size:10px;opacity:.8">Run Date</div></div>
</div>
<div class="stats" id="stats"></div>
<div class="ctrl">
  <label>Partner:</label>
  <select id="ep-f" onchange="filter()"><option value="">All Partners</option>{ep_opts}</select>
  <label>Status:</label>
  <select id="st-f" onchange="filter()">
    <option value="">All</option><option value="5">Critical (Yr 5+)</option>
    <option value="4">Warning (Yr 4)</option><option value="3">Monitor (Yr 3)</option>
    <option value="1">OK (Yr 1-2)</option>
  </select>
  <label>Year:</label>
  <select id="yr-f" onchange="filter()"><option value="">All Years</option>{yr_opts}</select>
  <label>Search:</label>
  <input id="srch" type="text" placeholder="Issuer or partner..." oninput="filter()" style="width:180px">
  <label style="margin-left:8px">
    <input type="checkbox" id="gap-only" onchange="filter()"> Gap flags only
  </label>
  <div class="tabs">
    <button class="tb active" onclick="tab('dash')">Rotation Dashboard</button>
    <button class="tb" onclick="tab('hist')">Full Filing History</button>
  </div>
</div>
<div class="main">
<!-- Task 5: EQR assumption note -->
<div class="note">
  <strong>PCAOB AS 1201:</strong> Lead EP and EQR must rotate after <strong>5 consecutive years</strong>, followed by a 5-year cooling-off period.
  &nbsp;<span style="background:#E2EFDA;border:1px solid #70AD47;border-radius:3px;padding:1px 6px;font-size:10px;color:#375623;font-weight:600">Green</span> EQR badge = live from Bizinta.
  &nbsp;<strong>EQR Assumption B (CRITICAL):</strong> EQR tenure assumed to start same year as EP. PCAOB Form AP does not record EQR appointment dates. Human review required where firm holds internal EQR appointment records.
</div>
<!-- Task 12: Gap-year assumption note -->
<div class="note-danger">
  <strong>Gap-Year Assumption A (CRITICAL):</strong> A year with no Form AP filing resets the consecutive count to Year 1 on return.
  Rows marked <span class="gap-badge">GAP</span> indicate a gap-then-return pattern requiring human review.
  An engineered gap to reset the clock constitutes circumvention under PCAOB rules.
</div>
<div id="t-dash">
  <div class="leg">
    <span style="font-size:11px;font-weight:700;color:#555;margin-right:4px">Legend:</span>
    <div class="li"><div class="ld" style="background:#FF0000"></div>Year 5+ CRITICAL</div>
    <div class="li"><div class="ld" style="background:#FF8C00"></div>Year 4 WARNING</div>
    <div class="li"><div class="ld" style="background:#FFD700"></div>Year 3 MONITOR</div>
    <div class="li"><div class="ld" style="background:#92D050"></div>Year 2 OK</div>
    <div class="li"><div class="ld" style="background:#00B050"></div>Year 1 OK</div>
    <div class="li" style="margin-left:12px"><span style="font-size:10px;color:#777">Click any row for drill-down</span></div>
  </div>
  <div class="stit">Current Engagement Status <span class="cnt" id="dc">0</span></div>
  <table><thead><tr>
    <th>Engagement Partner</th><th>Issuer / Client</th>
    <th style="text-align:center">EP Start Yr</th>
    <th style="text-align:center">Audit Yr</th><th style="text-align:center">Consecutive Yrs</th>
    <th style="text-align:center">Yrs Left</th><th>Rotation Status</th>
    <th>Signer</th><th>EQR</th>
  </tr></thead><tbody id="db"></tbody></table>
</div>
<div id="t-hist" class="hidden">
  <div class="stit">All Form AP Filings <span class="cnt" id="hc">0</span></div>
  <table><thead><tr>
    <th>Filed</th><th style="text-align:center">Yr</th>
    <th>Engagement Partner</th><th>Issuer / Client</th>
    <th style="text-align:center">Rotation Yr #</th><th>Status</th>
    <th>Signer</th><th>EQR</th>
  </tr></thead><tbody id="hb"></tbody></table>
</div>
</div>


<!-- Task 4 & 9: Drill-down modal -->
<div class="modal-overlay" id="modal" onclick="closeModal(event)">
  <div class="modal">
    <div class="modal-hdr">
      <h2 id="modal-title">Engagement Detail</h2>
      <button class="modal-close" onclick="closeModalBtn()">x</button>
    </div>
    <div id="modal-body"></div>
  </div>
</div>

<script>
const D={dash_js};const A={all_js};

/* Index maps - avoids JSON-in-HTML-attribute entirely */
const DM={{}};D.forEach((r,i)=>DM[i]=r);
const AM={{}};A.forEach((r,i)=>AM[i]=r);

function pc(sc){{return sc>=5?'p5':sc==4?'p4':sc==3?'p3':sc==2?'p2':'p1'}}
function bar(c){{let h='<div class="bar">';for(let i=1;i<=5;i++)h+=`<div class="bk ${{i<=c?'f'+Math.min(c,5):'fe'}}"></div>`;return h+'</div>';}}
function slabel(sc){{const m={{5:'CRITICAL Yr 5+',4:'WARNING Yr 4',3:'MONITOR Yr 3',2:'OK Yr 2',1:'OK Yr 1'}};return m[Math.min(sc,5)];}}
function esc(s){{const d=document.createElement('div');d.textContent=s;return d.innerHTML;}}

function renderStats(){{
  const c5=D.filter(r=>r.sc>=5).length,c4=D.filter(r=>r.sc==4).length,c3=D.filter(r=>r.sc==3).length,ok=D.filter(r=>r.sc<=2).length;
  document.getElementById('stats').innerHTML=
    `<div class="sc s5c"><div class="n">${{c5}}</div><div class="l">Critical (Yr 5+)</div></div>`+
    `<div class="sc s4c"><div class="n">${{c4}}</div><div class="l">Warning (Yr 4)</div></div>`+
    `<div class="sc s3c"><div class="n">${{c3}}</div><div class="l">Monitor (Yr 3)</div></div>`+
    `<div class="sc okc"><div class="n">${{ok}}</div><div class="l">OK (Yr 1-2)</div></div>`+
    `<div class="sc ttc"><div class="n">${{D.length}}</div><div class="l">Engagements</div></div>`;
}}

/* Task 4 & 9: drill-down modal - receives row object directly, no HTML attribute encoding */
function openModal(r){{
  const hist = A.filter(x=>x.ep===r.ep && x.issuer===r.issuer).sort((a,b)=>a.yr-b.yr);
  const histRows = hist.map(x=>`
    <tr>
      <td>${{x.yr}}</td>
      <td>${{x.filed}}</td>
      <td style="text-align:center">${{bar(x.consec)}} Yr ${{x.consec}}</td>
      <td><span class="pill ${{pc(x.sc)}}">${{slabel(x.sc)}}</span></td>
      <td>${{x.hasGap?'<span class="gap-badge">GAP</span>':''}}</td>
    </tr>`).join('');
  const gapWarn = r.hasGap
    ? `<div class="gap-warn"><strong>GAP FLAG:</strong> A gap year was detected in this partner's filing history for this issuer. The consecutive count restarted after the gap. Human review required to confirm this was not an engineered reset.</div>`
    : '';
  const eqrHtml = r.eqr
    ? `<span class="eqr-biz">${{esc(r.eqr)}}</span>`
    : 'Not in Bizinta';
  const leftHtml = r.left===0
    ? '<strong style="color:#cc0000">ROTATE NOW</strong>'
    : `<strong style="color:${{r.left===1?'#cc5500':'#006600'}}">${{r.left}} year(s)</strong>`;
  document.getElementById('modal-title').textContent = r.ep + ' \u2013 ' + r.issuer;
  document.getElementById('modal-body').innerHTML = `
    <div class="modal-section">
      <h3>Engagement Summary</h3>
      <div class="modal-grid">
        <span class="modal-label">Partner</span><span class="modal-value"><strong>${{esc(r.ep)}}</strong></span>
        <span class="modal-label">Issuer</span><span class="modal-value">${{esc(r.issuer)}}</span>
        <span class="modal-label">EP Start Year</span><span class="modal-value"><span class="start-yr">${{r.startYr||r.yr}}</span></span>
        <span class="modal-label">Latest Audit Year</span><span class="modal-value">${{r.yr}}</span>
        <span class="modal-label">EQR</span><span class="modal-value">${{eqrHtml}}</span>
        <span class="modal-label">Signer</span><span class="modal-value">${{esc(r.signer||'-')}}</span>
        <span class="modal-label">Status</span><span class="modal-value"><span class="pill ${{pc(r.sc)}}">${{slabel(r.sc)}}</span></span>
        <span class="modal-label">Years Remaining</span><span class="modal-value">${{leftHtml}}</span>
      </div>
    </div>
    <div class="modal-section">
      <h3>Calculation Logic</h3>
      <div class="calc-box">${{esc(r.calcNote)}}</div>
      ${{gapWarn}}
    </div>
    <div class="modal-section">
      <h3>Full Filing History (${{hist.length}} year(s))</h3>
      <table class="hist-table">
        <thead><tr><th>Audit Yr</th><th>Filed</th><th>Consecutive</th><th>Status</th><th></th></tr></thead>
        <tbody>${{histRows}}</tbody>
      </table>
    </div>`;
  document.getElementById('modal').classList.add('open');
}}
function closeModal(e){{if(e.target===document.getElementById('modal'))closeModalBtn();}}
function closeModalBtn(){{document.getElementById('modal').classList.remove('open');}}
document.addEventListener('keydown',e=>{{if(e.key==='Escape')closeModalBtn();}});

/* Attach click handlers after render using data-idx, not inline JSON */
function attachClicks(tbodyId, map){{
  document.getElementById(tbodyId).querySelectorAll('tr[data-idx]').forEach(tr=>{{
    tr.addEventListener('click',()=>openModal(map[+tr.dataset.idx]));
  }});
}}

function renderDash(data){{
  const tb=document.getElementById('db');document.getElementById('dc').textContent=data.length;
  if(!data.length){{tb.innerHTML='<tr><td colspan="9" class="nr">No records match.</td></tr>';return;}}
  tb.innerHTML=data.map((r,i)=>`<tr data-idx="${{i}}" class="clickable${{r.hasGap?' gap-row':''}}">
    <td><strong>${{esc(r.ep)}}</strong></td>
    <td>${{esc(r.issuer)}}${{r.hasGap?'<span class="gap-badge">GAP</span>':''}}</td>
    <td style="text-align:center"><span class="start-yr">${{r.startYr||r.yr}}</span></td>
    <td style="text-align:center">${{r.yr}}</td>
    <td style="text-align:center">${{bar(r.consec)}} <span style="font-size:10px;color:#666">Yr ${{r.consec}}</span></td>
    <td style="text-align:center"><strong style="color:${{r.left==0?'#cc0000':r.left==1?'#cc5500':'#006600'}}">${{r.left==0?'ROTATE NOW':r.left}}</strong></td>
    <td><span class="pill ${{pc(r.sc)}}">${{slabel(r.sc)}}</span></td>
    <td style="color:#555;font-size:11px">${{esc(r.signer)}}</td>
    <td>${{r.eqr?`<span class="eqr-biz">${{esc(r.eqr)}}</span>`:'<span style="color:#aaa;font-size:11px">-</span>'}}</td>
  </tr>`).join('');
  /* store filtered data in closure for click lookup */
  const localMap={{}};data.forEach((r,i)=>localMap[i]=r);
  document.getElementById('db').querySelectorAll('tr[data-idx]').forEach(tr=>{{
    tr.addEventListener('click',()=>openModal(localMap[+tr.dataset.idx]));
  }});
}}

function renderHist(data){{
  const tb=document.getElementById('hb');document.getElementById('hc').textContent=data.length;
  if(!data.length){{tb.innerHTML='<tr><td colspan="8" class="nr">No records match.</td></tr>';return;}}
  tb.innerHTML=data.map((r,i)=>`<tr data-idx="${{i}}" class="clickable${{r.hasGap?' gap-row':''}}">
    <td>${{r.filed}}</td><td style="text-align:center">${{r.yr}}</td>
    <td><strong>${{esc(r.ep)}}</strong></td>
    <td>${{esc(r.issuer)}}${{r.hasGap?'<span class="gap-badge">GAP</span>':''}}</td>
    <td style="text-align:center">${{bar(r.consec)}}</td>
    <td><span class="pill ${{pc(r.sc)}}">${{slabel(r.sc)}}</span></td>
    <td style="color:#555;font-size:11px">${{esc(r.signer)}}</td>
    <td style="font-size:11px;color:#375623">${{esc(r.eqr||'')}}</td>
  </tr>`).join('');
  const localMap={{}};data.forEach((r,i)=>localMap[i]=r);
  document.getElementById('hb').querySelectorAll('tr[data-idx]').forEach(tr=>{{
    tr.addEventListener('click',()=>openModal(localMap[+tr.dataset.idx]));
  }});
}}

function filter(){{
  const ep=document.getElementById('ep-f').value,
        st=document.getElementById('st-f').value,
        yr=document.getElementById('yr-f').value,
        q=document.getElementById('srch').value.toLowerCase(),
        gapOnly=document.getElementById('gap-only').checked;
  const fd=r=>{{
    if(ep&&r.ep!==ep)return false;
    if(st){{const c=parseInt(st);if(c==1&&r.sc>2)return false;if(c>1&&r.sc!==c)return false;}}
    if(yr&&String(r.yr)!==yr)return false;
    if(q&&!r.ep.toLowerCase().includes(q)&&!r.issuer.toLowerCase().includes(q))return false;
    if(gapOnly&&!r.hasGap)return false;
    return true;
  }};
  renderDash(D.filter(fd));renderHist(A.filter(fd));
}}

function tab(t){{
  document.getElementById('t-dash').classList.toggle('hidden',t!=='dash');
  document.getElementById('t-hist').classList.toggle('hidden',t!=='hist');
  document.querySelectorAll('.tb').forEach((b,i)=>b.classList.toggle('active',(i==0&&t=='dash')||(i==1&&t=='hist')));
}}

document.getElementById('hdr-date').textContent=new Date().toLocaleDateString('en-GB',{{day:'2-digit',month:'short',year:'numeric'}});
renderStats();renderDash(D);renderHist(A);
</script></body></html>"""
    return html


# -- Email alert builder
def build_email_alert(dashboard, run_date_str):
    critical = [r for r in dashboard if r["sc"] >= 5]
    warning  = [r for r in dashboard if r["sc"] == 4]
    monitor  = [r for r in dashboard if r["sc"] == 3]

    subject = (f"PCAOB Rotation Alert - {run_date_str} | "
               f"{len(critical)} Critical, {len(warning)} Warning, {len(monitor)} Monitor")

    def rows_html(items, bg, fg, label):
        if not items:
            return f'<tr><td colspan="5" style="color:#888;font-style:italic;padding:10px 12px">No {label} engagements</td></tr>'
        return "".join(f"""
        <tr>
          <td style="padding:9px 12px;border-bottom:1px solid #eee;font-weight:600">{r['ep']}</td>
          <td style="padding:9px 12px;border-bottom:1px solid #eee">{r['issuer']}</td>
          <td style="padding:9px 12px;border-bottom:1px solid #eee;text-align:center">{r['year']}</td>
          <td style="padding:9px 12px;border-bottom:1px solid #eee;text-align:center">
            <span style="background:{bg};color:{fg};padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700">
              Year {r['consec']}
            </span>
          </td>
          <td style="padding:9px 12px;border-bottom:1px solid #eee;text-align:center">
            {'<strong style="color:#cc0000">ROTATE NOW</strong>' if r['yrs_left']==0 else f'{r["yrs_left"]} yr{"s" if r["yrs_left"]!=1 else ""}'}
          </td>
        </tr>""" for r in items)

    th = 'style="background:#1F4E79;color:#fff;padding:9px 12px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase"'

    html_body = f"""
<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#f4f6f9;margin:0;padding:0">
<div style="max-width:700px;margin:24px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.1)">
  <div style="background:linear-gradient(135deg,#1F4E79,#2E75B6);padding:24px 28px;color:#fff">
    <h1 style="margin:0;font-size:20px">PCAOB Partner Rotation Alert</h1>
    <p style="margin:6px 0 0;font-size:12px;opacity:.85">Kreit &amp; Chiu CPA LLP &nbsp;|&nbsp; Firm ID 6651 &nbsp;|&nbsp; Generated: {run_date_str}</p>
  </div>
  <div style="display:flex;gap:0;border-bottom:3px solid #eee">
    <div style="flex:1;text-align:center;padding:16px;border-right:1px solid #eee">
      <div style="font-size:28px;font-weight:700;color:#cc0000">{len(critical)}</div>
      <div style="font-size:11px;color:#666;margin-top:3px;text-transform:uppercase;letter-spacing:.5px">Critical (Yr 5+)</div>
    </div>
    <div style="flex:1;text-align:center;padding:16px;border-right:1px solid #eee">
      <div style="font-size:28px;font-weight:700;color:#cc5500">{len(warning)}</div>
      <div style="font-size:11px;color:#666;margin-top:3px;text-transform:uppercase;letter-spacing:.5px">Warning (Yr 4)</div>
    </div>
    <div style="flex:1;text-align:center;padding:16px;border-right:1px solid #eee">
      <div style="font-size:28px;font-weight:700;color:#806600">{len(monitor)}</div>
      <div style="font-size:11px;color:#666;margin-top:3px;text-transform:uppercase;letter-spacing:.5px">Monitor (Yr 3)</div>
    </div>
    <div style="flex:1;text-align:center;padding:16px">
      <div style="font-size:28px;font-weight:700;color:#1F4E79">{len(dashboard)}</div>
      <div style="font-size:11px;color:#666;margin-top:3px;text-transform:uppercase;letter-spacing:.5px">Total Engagements</div>
    </div>
  </div>
  <div style="padding:24px 28px">
    <h2 style="font-size:14px;color:#cc0000;margin:0 0 10px;padding-bottom:6px;border-bottom:2px solid #FF0000">
      CRITICAL - Rotate Immediately (Year 5+)
    </h2>
    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;font-size:12px">
      <tr><th {th}>Partner</th><th {th}>Client / Issuer</th><th {th} style="text-align:center">Audit Yr</th><th {th} style="text-align:center">Status</th><th {th} style="text-align:center">Years Left</th></tr>
      {rows_html(critical, '#FF0000', '#fff', 'critical')}
    </table>
    <h2 style="font-size:14px;color:#cc5500;margin:0 0 10px;padding-bottom:6px;border-bottom:2px solid #FF8C00">
      WARNING - Plan Rotation (Year 4 - 1 Year Remaining)
    </h2>
    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;font-size:12px">
      <tr><th {th}>Partner</th><th {th}>Client / Issuer</th><th {th} style="text-align:center">Audit Yr</th><th {th} style="text-align:center">Status</th><th {th} style="text-align:center">Years Left</th></tr>
      {rows_html(warning, '#FF8C00', '#fff', 'warning')}
    </table>
    <h2 style="font-size:14px;color:#806600;margin:0 0 10px;padding-bottom:6px;border-bottom:2px solid #FFD700">
      MONITOR - On Watch (Year 3 - 2 Years Remaining)
    </h2>
    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;font-size:12px">
      <tr><th {th}>Partner</th><th {th}>Client / Issuer</th><th {th} style="text-align:center">Audit Yr</th><th {th} style="text-align:center">Status</th><th {th} style="text-align:center">Years Left</th></tr>
      {rows_html(monitor, '#FFD700', '#333', 'monitor')}
    </table>
    <div style="background:#e8f4fd;border-left:4px solid #2E75B6;border-radius:0 6px 6px 0;padding:12px 16px;font-size:11px;color:#1a3a5c;line-height:1.7">
      <strong>PCAOB AS 1201:</strong> Both the lead engagement partner (EP) and the engagement quality reviewer (EQR)
      must rotate after 5 consecutive years. A 5-year cooling-off period applies after mandatory rotation.<br>
      <strong>EQR Assumption:</strong> EQR tenure assumed to start same year as EP. Human review required where firm holds different internal records.<br>
      Full Excel tracker and interactive HTML dashboard are attached to this email.
    </div>
  </div>
  <div style="background:#f4f6f9;padding:12px 28px;font-size:10px;color:#999;text-align:center">
    Auto-generated by PCAOB Rotation Build Service &nbsp;|&nbsp; Kreit &amp; Chiu CPA LLP &nbsp;|&nbsp;
    This email and its attachments contain confidential firm data. Do not forward externally.
  </div>
</div>
</body></html>"""

    return subject, html_body, {
        "critical_count": len(critical),
        "warning_count":  len(warning),
        "monitor_count":  len(monitor),
        "total":          len(dashboard),
        "critical": [{"ep":r["ep"],"issuer":r["issuer"],"year":r["year"],"consec":r["consec"]} for r in critical],
        "warning":  [{"ep":r["ep"],"issuer":r["issuer"],"year":r["year"],"consec":r["consec"]} for r in warning],
    }


# -- Routes
@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "pcaob-build-service"})

@app.route("/build", methods=["POST"])
@require_api_key
def build():
    run_date      = date.today()
    run_date_str  = run_date.strftime("%d %b %Y")
    filename_date = run_date.strftime("%Y-%m-%d")

    try:
        bizinta_data = fetch_bizinta()
    except Exception as e:
        return jsonify({"error": f"Bizinta API failed: {e}"}), 502

    try:
        records, raw_pcaob_rows = fetch_pcaob()
    except Exception as e:
        return jsonify({"error": f"PCAOB filter service failed: {e}"}), 502

    # Build EQR lookup: pcaob_issuer_name -> eqr
    pcaob_to_eqr = {}
    biz_all_names = set(bizinta_data.keys())
    for biz_name, biz in bizinta_data.items():
        mapped_pcaob = {v for k, v in BIZINTA_TO_PCAOB.items() if v and k in biz_all_names}
        pcaob_name = BIZINTA_TO_PCAOB.get(biz_name)
        if pcaob_name and biz.get("eqr"):
            pcaob_to_eqr[pcaob_name] = biz["eqr"]

    enriched, dashboard = build_rotation(records)

    excel_bytes   = build_excel(enriched, dashboard, pcaob_to_eqr, bizinta_data, raw_pcaob_rows, run_date_str)
    html_str      = build_html(enriched, dashboard, pcaob_to_eqr, run_date_str)
    email_subject, email_html, summary = build_email_alert(dashboard, run_date_str)

    return jsonify({
        "excel_b64":       base64.b64encode(excel_bytes).decode("utf-8"),
        "html_b64":        base64.b64encode(html_str.encode("utf-8")).decode("utf-8"),
        "excel_filename":  f"PCAOB_Rotation_Tracker_KreitChiu_{filename_date}.xlsx",
        "html_filename":   f"PCAOB_Rotation_Dashboard_KreitChiu_{filename_date}.html",
        "email_subject":   email_subject,
        "email_html":      email_html,
        "summary":         summary,
        "run_date":        filename_date,
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
