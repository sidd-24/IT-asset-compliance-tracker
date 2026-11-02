"""
compliance_checker.py
──────────────────────
Scans the asset repository, auto-flags lifecycle status changes,
and prints a summary of what needs attention.

Usage:
    python compliance_checker.py             # Summary only
    python compliance_checker.py --update    # Also updates Lifecycle Status in repo
"""

import argparse
import sys
from datetime import date

import pandas as pd
from openpyxl import load_workbook

from compliance_rules import add_deadline_state, refresh_lifecycle
from config import REPO_FILE, SHEETS, NEAR_EXPIRY_DAYS


def load_assets():
    df = pd.read_excel(REPO_FILE, sheet_name=SHEETS["assets"])
    if "Expiry / Warranty Date" in df.columns:
        df["Expiry / Warranty Date"] = pd.to_datetime(df["Expiry / Warranty Date"], errors="coerce").dt.date
    return df


def run_check(update=False):
    today   = date.today()
    stored  = load_assets()
    assets  = refresh_lifecycle(stored, today)     # statuses as of today, not as last saved
    issues  = add_deadline_state(pd.read_excel(REPO_FILE, sheet_name=SHEETS["issues"]), today)

    changed = []
    for (_, old), (_, new) in zip(stored.iterrows(), assets.iterrows()):
        if str(old.get("Lifecycle Status", "")).strip() != new["Lifecycle Status"]:
            changed.append({
                "Asset ID":    new["Asset ID"],
                "Asset Name":  new["Asset Name"],
                "Old Status":  str(old.get("Lifecycle Status", "")).strip(),
                "New Status":  new["Lifecycle Status"],
                "Expiry Date": new["Expiry / Warranty Date"],
            })

    # ── Print Summary ────────────────────────────────────────────────────────
    print("\n" + "="*60)
    print(f"  Compliance Check  |  {today.strftime('%d %b %Y')}")
    print("="*60)

    total    = len(assets)
    expired  = assets[assets["Lifecycle Status"] == "Expired"]
    near_exp = assets[assets["Lifecycle Status"] == "Near Expiry"]
    unknown  = assets[assets["Lifecycle Status"] == "Unknown Expiry"]
    non_comp = assets[assets["Compliance Status"] == "Non-Compliant"]
    pending  = assets[assets["Compliance Status"] == "Pending Review"]
    overdue  = issues[issues["Deadline State"] == "Overdue"].sort_values("Days Overdue", ascending=False)
    due_soon = issues[issues["Deadline State"] == "Due Soon"]

    print(f"\n  Total Assets      : {total}")
    print(f"  Expired           : {len(expired)}")
    print(f"  Near Expiry       : {len(near_exp)}")
    print(f"  Unknown Expiry    : {len(unknown)}")
    print(f"  Non-Compliant     : {len(non_comp)}")
    print(f"  Pending Review    : {len(pending)}")
    print(f"  Overdue Issues    : {len(overdue)}")
    print(f"  Issues Due Soon   : {len(due_soon)}")

    if overdue.shape[0]:
        print("\n  ── OVERDUE REMEDIATION ──────────────────────────────")
        for _, r in overdue.iterrows():
            print(f"  [{r['Issue ID']}] {r['Severity']:<6} {r['Days Overdue']:>3} days overdue  |  "
                  f"{r['Asset ID']}  |  owner {r['Remediation Owner']}")

    if expired.shape[0]:
        print("\n  ── EXPIRED ASSETS ──────────────────────────────────")
        for _, r in expired.iterrows():
            print(f"  [{r['Asset ID']}] {r['Asset Name']}  |  Expired: {r['Expiry / Warranty Date']}")

    if near_exp.shape[0]:
        print(f"\n  ── NEAR EXPIRY (within {NEAR_EXPIRY_DAYS} days) ─────────────────")
        for _, r in near_exp.iterrows():
            print(f"  [{r['Asset ID']}] {r['Asset Name']}  |  Expires: {r['Expiry / Warranty Date']}")

    if non_comp.shape[0]:
        print("\n  ── NON-COMPLIANT ASSETS ─────────────────────────────")
        for _, r in non_comp.iterrows():
            print(f"  [{r['Asset ID']}] {r['Asset Name']}  |  Status: {r['Compliance Status']}")

    if changed:
        print(f"\n  ── LIFECYCLE STATUS CHANGES DETECTED ({len(changed)}) ──────")
        for c in changed:
            print(f"  [{c['Asset ID']}] {c['Asset Name']}  |  {c['Old Status']} → {c['New Status']}")
        if update:
            print("\n  Saving updated statuses to repository...")
            _write_lifecycle_updates(assets)
            print("  Done.")
        else:
            print("\n  Run with --update to apply these changes to the repository.")
    else:
        print("\n  No lifecycle status changes detected.")

    print("\n" + "="*60 + "\n")


def _write_lifecycle_updates(updated_df):
    wb = load_workbook(REPO_FILE)
    ws = wb[SHEETS["assets"]]

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    try:
        id_col     = headers.index("Asset ID") + 1
        status_col = headers.index("Lifecycle Status") + 1
    except ValueError:
        print("  Column not found. Check sheet headers.")
        return

    id_to_status = dict(zip(updated_df["Asset ID"], updated_df["Lifecycle Status"]))

    for row in ws.iter_rows(min_row=2):
        asset_id = row[id_col - 1].value
        if asset_id in id_to_status:
            row[status_col - 1].value = id_to_status[asset_id]

    wb.save(REPO_FILE)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")   # box-drawing characters crash cp1252 Windows consoles
    parser = argparse.ArgumentParser()
    parser.add_argument("--update", action="store_true", help="Write status changes back to the repo")
    args = parser.parse_args()
    run_check(update=args.update)
