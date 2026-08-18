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
from datetime import date, timedelta

import pandas as pd
from openpyxl import load_workbook

from config import REPO_FILE, SHEETS, NEAR_EXPIRY_DAYS


def load_assets():
    df = pd.read_excel(REPO_FILE, sheet_name=SHEETS["assets"])
    if "Expiry / Warranty Date" in df.columns:
        df["Expiry / Warranty Date"] = pd.to_datetime(df["Expiry / Warranty Date"], errors="coerce").dt.date
    return df


def compute_lifecycle_status(expiry_date, today):
    if pd.isna(expiry_date) or expiry_date is None:
        return "Active"
    if expiry_date < today:
        return "Expired"
    if expiry_date <= today + timedelta(days=NEAR_EXPIRY_DAYS):
        return "Near Expiry"
    return "Active"


def run_check(update=False):
    today  = date.today()
    assets = load_assets()

    changed = []
    for idx, row in assets.iterrows():
        expiry     = row.get("Expiry / Warranty Date")
        current    = str(row.get("Lifecycle Status", "")).strip()
        computed   = compute_lifecycle_status(expiry, today)

        if computed != current and current != "Retired":
            changed.append({
                "Asset ID":    row.get("Asset ID"),
                "Asset Name":  row.get("Asset Name"),
                "Old Status":  current,
                "New Status":  computed,
                "Expiry Date": expiry,
            })
            if update:
                assets.at[idx, "Lifecycle Status"] = computed

    # ── Print Summary ────────────────────────────────────────────────────────
    print("\n" + "="*60)
    print(f"  Compliance Check  |  {today.strftime('%d %b %Y')}")
    print("="*60)

    total    = len(assets)
    expired  = assets[assets["Lifecycle Status"] == "Expired"]
    near_exp = assets[assets["Lifecycle Status"] == "Near Expiry"]
    non_comp = assets[assets["Compliance Status"] == "Non-Compliant"]
    pending  = assets[assets["Compliance Status"] == "Pending Review"]

    print(f"\n  Total Assets      : {total}")
    print(f"  Expired           : {len(expired)}")
    print(f"  Near Expiry       : {len(near_exp)}")
    print(f"  Non-Compliant     : {len(non_comp)}")
    print(f"  Pending Review    : {len(pending)}")

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
    parser = argparse.ArgumentParser()
    parser.add_argument("--update", action="store_true", help="Write status changes back to the repo")
    args = parser.parse_args()
    run_check(update=args.update)
