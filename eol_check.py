"""
eol_check.py
─────────────
Flags hardware and software running a platform version that is out of
vendor support, using the public endoflife.date API (no key needed).
Unsupported software is a classic ISO 27001 finding (A.8.8, management of
technical vulnerabilities): it no longer receives security patches.

Each asset names its platform in two columns of the Asset_Inventory sheet:
  "Platform (endoflife.date)"  product id, e.g. windows-server, cisco-ios-xe, ubuntu
  "Platform Version"           release cycle, e.g. 2019, 17.3, 22.04

Usage:
    python eol_check.py                  # report to reports/EOL_Check_<date>.xlsx
    python eol_check.py --raise-issues   # also add A.8.8 issues to the repository
    python eol_check.py --offline        # use cached API responses only
"""

import argparse
import json
import os
import urllib.request
from datetime import date, timedelta

import pandas as pd
from openpyxl import Workbook, load_workbook

from compliance_rules import refresh_lifecycle
from config import CACHE_FOLDER, EOL_API_URL, EOL_WARNING_DAYS, REMEDIATION_SLA_DAYS, REPO_FILE, REPORTS_FOLDER, SHEETS

PLATFORM_COL = "Platform (endoflife.date)"
VERSION_COL = "Platform Version"

EOL_STATES = ["End of Life", f"EOL within {EOL_WARNING_DAYS} days", "Supported", "Unknown"]


# ─── endoflife.date client with a local cache ─────────────────────────────────

def fetch_cycles(product, offline=False, cache_dir=os.path.join(CACHE_FOLDER, "eol")):
    """Release cycles of a product. Cached on disk so repeated runs work offline."""
    os.makedirs(cache_dir, exist_ok=True)
    cache_file = os.path.join(cache_dir, f"{product}.json")
    if not offline:
        try:
            request = urllib.request.Request(EOL_API_URL.format(product=product),
                                             headers={"Accept": "application/json", "User-Agent": "asset-compliance-tracker"})
            with urllib.request.urlopen(request, timeout=20) as response:
                cycles = json.loads(response.read().decode("utf-8"))
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(cycles, f)
            return cycles
        except Exception as exc:
            print(f"  Could not reach endoflife.date for '{product}' ({exc}); trying cache.")
    if os.path.exists(cache_file):
        with open(cache_file, encoding="utf-8") as f:
            return json.load(f)
    return None


# ─── Classification ───────────────────────────────────────────────────────────

def find_cycle(cycles, version):
    """Exact cycle match first, else the longest cycle the version starts with (e.g. 17.3.4 -> 17.3)."""
    version = str(version).strip().lower()
    if version.endswith(".0") and version[:-2].replace(".", "").isdigit():
        version = version[:-2]
    by_name = {str(c.get("cycle", "")).lower(): c for c in cycles}
    if version in by_name:
        return by_name[version]
    prefixes = [name for name in by_name if version.startswith(name + ".") or version.startswith(name + "-")]
    return by_name[max(prefixes, key=len)] if prefixes else None


def classify_eol(cycle, today):
    """(state, eol_date or None) for one release cycle."""
    if cycle is None:
        return "Unknown", None
    eol = cycle.get("eol")
    if eol is True:
        return "End of Life", None
    if eol is False or eol is None:
        return "Supported", None
    eol_date = date.fromisoformat(str(eol)[:10])
    if eol_date < today:
        return "End of Life", eol_date
    if eol_date <= today + timedelta(days=EOL_WARNING_DAYS):
        return f"EOL within {EOL_WARNING_DAYS} days", eol_date
    return "Supported", eol_date


def check_assets(assets, today, cycles_for=None, offline=False):
    """
    One row per asset with a platform: state, end-of-life date, latest release.
    cycles_for(product) can be injected for testing; defaults to the cached API client.
    """
    cycles_for = cycles_for or (lambda product: fetch_cycles(product, offline=offline))
    cache = {}
    rows = []
    active = assets[assets["Lifecycle Status"] != "Retired"]
    for _, asset in active.iterrows():
        product = str(asset.get(PLATFORM_COL) or "").strip().lower()
        version = asset.get(VERSION_COL)
        if not product or product == "nan" or pd.isna(version):
            continue
        if product not in cache:
            cache[product] = cycles_for(product)
        cycles = cache[product]
        cycle = find_cycle(cycles, version) if cycles else None
        state, eol_date = classify_eol(cycle, today)
        rows.append({
            "Asset ID": asset["Asset ID"],
            "Asset Name": asset["Asset Name"],
            "Owner ID": asset.get("Owner ID"),
            "Platform": product,
            "Version": str(version),
            "EOL State": state,
            "End of Support": eol_date,
            "Days Left": (eol_date - today).days if eol_date else None,
            "Latest Release": (cycle or {}).get("latest"),
            "Note": "" if cycle else ("Product not found on endoflife.date" if not cycles else "Version not found"),
        })
    return pd.DataFrame(rows, columns=["Asset ID", "Asset Name", "Owner ID", "Platform", "Version", "EOL State",
                                       "End of Support", "Days Left", "Latest Release", "Note"])


# ─── Outputs ──────────────────────────────────────────────────────────────────

def write_report(results, today):
    from generate_reports import CENTER, TITLE_FILL, TITLE_FONT, data_cell, h, set_widths, RED_FILL, YELLOW_FILL, GREEN_FILL

    os.makedirs(REPORTS_FOLDER, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "EOL Check"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:J1")
    ws["A1"].value = f"Platform End-of-Life Check  |  {today.strftime('%d %b %Y')}  |  Source: endoflife.date"
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = TITLE_FONT, TITLE_FILL, CENTER
    ws.row_dimensions[1].height = 36

    for i, col in enumerate(results.columns, 1):
        h(ws.cell(3, i, col))
    fills = {"End of Life": RED_FILL, f"EOL within {EOL_WARNING_DAYS} days": YELLOW_FILL, "Supported": GREEN_FILL}
    for r, (_, row) in enumerate(results.iterrows(), 4):
        for c, col in enumerate(results.columns, 1):
            value = row[col]
            cell = ws.cell(r, c, None if (not isinstance(value, str) and pd.isna(value)) else value)
            data_cell(cell, alt=(r % 2 == 0), fill=fills.get(row["EOL State"]) if col == "EOL State" else None)
            if isinstance(value, date):
                cell.number_format = "DD-MMM-YYYY"
    set_widths(ws, {"A": 10, "B": 28, "C": 10, "D": 16, "E": 10, "F": 20, "G": 15, "H": 10, "I": 15, "J": 30})
    ws.freeze_panes = "A4"
    path = os.path.join(REPORTS_FOLDER, f"EOL_Check_{today.isoformat()}.xlsx")
    wb.save(path)
    return path


def raise_issues(results, today, repo_file=REPO_FILE):
    """Adds an A.8.8 compliance issue for each unsupported platform that has no open issue yet."""
    issues = pd.read_excel(repo_file, sheet_name=SHEETS["issues"])
    open_eol = set(issues[(issues["Status"] != "Closed") &
                          (issues["Issue Description"].astype(str).str.startswith("Unsupported platform"))]["Asset ID"])
    flagged = results[results["EOL State"].isin(["End of Life", f"EOL within {EOL_WARNING_DAYS} days"])]
    new = flagged[~flagged["Asset ID"].isin(open_eol)]
    if new.empty:
        return 0

    wb = load_workbook(repo_file)
    ws = wb[SHEETS["issues"]]
    headers = [c.value for c in ws[1]]
    next_id = len(issues) + 1
    for _, row in new.iterrows():
        severity = "High" if row["EOL State"] == "End of Life" else "Medium"
        when = row["End of Support"].strftime("%d %b %Y") if isinstance(row["End of Support"], date) else "date unknown"
        values = {
            "Issue ID": f"ISS-{next_id:03d}",
            "Asset ID": row["Asset ID"],
            "Issue Description": f"Unsupported platform: {row['Platform']} {row['Version']} (end of support {when}). Upgrade or replace.",
            "Severity": severity,
            "Identified Date": today,
            "Remediation Owner": row["Owner ID"],
            "Due Date": today + timedelta(days=REMEDIATION_SLA_DAYS[severity]),
            "Status": "Open",
            "Resolution Notes": "Raised automatically by eol_check.py",
            "Last Updated": today,
            "ISO 27001 Control": "A.8.8",
        }
        ws.append([values.get(hd, "") for hd in headers])
        for c, hd in enumerate(headers, 1):
            if hd in ("Identified Date", "Due Date", "Last Updated"):
                ws.cell(ws.max_row, c).number_format = "DD-MMM-YYYY"
        next_id += 1
    wb.save(repo_file)
    return len(new)


def main():
    parser = argparse.ArgumentParser(description="Check asset platforms against endoflife.date")
    parser.add_argument("--raise-issues", action="store_true", help="Add A.8.8 issues for unsupported platforms")
    parser.add_argument("--offline", action="store_true", help="Use cached API responses only")
    args = parser.parse_args()

    today = date.today()
    assets = refresh_lifecycle(pd.read_excel(REPO_FILE, sheet_name=SHEETS["assets"]), today)
    results = check_assets(assets, today, offline=args.offline)
    path = write_report(results, today)

    counts = results["EOL State"].value_counts()
    print(f"\nChecked {len(results)} assets with a platform:")
    for state in EOL_STATES:
        print(f"  {state:<22} {counts.get(state, 0)}")
    print(f"Report saved: {path}")
    if args.raise_issues:
        print(f"New A.8.8 issues raised: {raise_issues(results, today)}")


if __name__ == "__main__":
    main()
