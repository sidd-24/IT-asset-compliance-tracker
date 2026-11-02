"""
license_reconciliation.py
──────────────────────────
Compares seats purchased with seats actually assigned for every software
license in the License_Entitlements sheet.

  Under-licensed  more seats in use than purchased: audit risk, the vendor can
                  bill the shortfall (true-up) at list price
  Over-licensed   utilisation below LICENSE_UNDERUSE_THRESHOLD: money spent on
                  unused seats that could be cut at renewal
  Balanced        everything else

ISO 27001 maps license compliance to A.5.32 (intellectual property rights).

Usage:
    python license_reconciliation.py      # report to reports/License_Reconciliation_<date>.xlsx
"""

import os
from datetime import date

import pandas as pd
from openpyxl import Workbook

from config import LICENSE_UNDERUSE_THRESHOLD, REPO_FILE, REPORTS_FOLDER, SHEETS


def reconcile(licenses: pd.DataFrame) -> pd.DataFrame:
    df = licenses.copy()
    purchased = pd.to_numeric(df["Seats Purchased"], errors="coerce").fillna(0)
    assigned = pd.to_numeric(df["Seats Assigned"], errors="coerce").fillna(0)
    cost = pd.to_numeric(df["Annual Cost per Seat (EUR)"], errors="coerce").fillna(0)

    df["Utilisation"] = (assigned / purchased.where(purchased > 0)).round(3)
    df["Unused Seats"] = (purchased - assigned).clip(lower=0).astype(int)
    df["Shortfall Seats"] = (assigned - purchased).clip(lower=0).astype(int)
    df["Annual Waste (EUR)"] = (df["Unused Seats"] * cost).round(2)
    df["True-up Exposure (EUR)"] = (df["Shortfall Seats"] * cost).round(2)

    def status(row):
        if row["Shortfall Seats"] > 0:
            return "Under-licensed"
        if pd.notna(row["Utilisation"]) and row["Utilisation"] < LICENSE_UNDERUSE_THRESHOLD:
            return "Over-licensed"
        return "Balanced"

    df["Reconciliation Status"] = df.apply(status, axis=1)
    return df


def summarise(result: pd.DataFrame) -> dict:
    return {
        "licenses": len(result),
        "under_licensed": int((result["Reconciliation Status"] == "Under-licensed").sum()),
        "over_licensed": int((result["Reconciliation Status"] == "Over-licensed").sum()),
        "annual_waste_eur": float(result["Annual Waste (EUR)"].sum()),
        "true_up_exposure_eur": float(result["True-up Exposure (EUR)"].sum()),
    }


def write_report(result: pd.DataFrame, today) -> str:
    from generate_reports import (CENTER, GREEN_FILL, RED_FILL, TITLE_FILL, TITLE_FONT, YELLOW_FILL,
                                  data_cell, h, set_widths, write_kpi_row, write_section_title)

    os.makedirs(REPORTS_FOLDER, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "License Reconciliation"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:L1")
    ws["A1"].value = f"Software License Reconciliation  |  {today.strftime('%d %b %Y')}"
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = TITLE_FONT, TITLE_FILL, CENTER
    ws.row_dimensions[1].height = 36

    s = summarise(result)
    write_section_title(ws, 3, "SUMMARY", col_span=12)
    rows = [
        ("Licenses reviewed", s["licenses"], ""),
        ("Under-licensed", s["under_licensed"], "Audit risk: buy seats or remove installations"),
        ("Over-licensed", s["over_licensed"], f"Utilisation below {LICENSE_UNDERUSE_THRESHOLD:.0%}"),
        ("Annual spend on unused seats (EUR)", f"{s['annual_waste_eur']:,.0f}", "Savings potential at renewal"),
        ("True-up exposure (EUR)", f"{s['true_up_exposure_eur']:,.0f}", "What a vendor audit could bill today"),
    ]
    for i, (label, value, note) in enumerate(rows, 4):
        write_kpi_row(ws, i, label, value, note)

    columns = ["License ID", "Asset ID", "Product", "Vendor", "License Metric", "Seats Purchased", "Seats Assigned",
               "Utilisation", "Unused Seats", "Shortfall Seats", "Annual Waste (EUR)", "True-up Exposure (EUR)",
               "Reconciliation Status"]
    start = 4 + len(rows) + 2
    write_section_title(ws, start, "DETAIL BY LICENSE", col_span=len(columns))
    for i, col in enumerate(columns, 1):
        h(ws.cell(start + 1, i, col))
    fills = {"Under-licensed": RED_FILL, "Over-licensed": YELLOW_FILL, "Balanced": GREEN_FILL}
    for r, (_, row) in enumerate(result.iterrows(), start + 2):
        for c, col in enumerate(columns, 1):
            value = row.get(col, "")
            cell = ws.cell(r, c, None if (not isinstance(value, str) and pd.isna(value)) else value)
            data_cell(cell, alt=(r % 2 == 0), fill=fills.get(value) if col == "Reconciliation Status" else None)
            if col == "Utilisation":
                cell.number_format = "0%"
            if "EUR" in col:
                cell.number_format = "#,##0"
    set_widths(ws, {"A": 34, "B": 10, "C": 24, "D": 14, "E": 14, "F": 11, "G": 11, "H": 11,
                    "I": 10, "J": 10, "K": 14, "L": 16, "M": 20})
    path = os.path.join(REPORTS_FOLDER, f"License_Reconciliation_{today.isoformat()}.xlsx")
    wb.save(path)
    return path


def main():
    today = date.today()
    result = reconcile(pd.read_excel(REPO_FILE, sheet_name=SHEETS["licenses"]))
    path = write_report(result, today)
    s = summarise(result)
    print(f"\nLicenses reviewed: {s['licenses']}")
    print(f"  Under-licensed: {s['under_licensed']}   True-up exposure: EUR {s['true_up_exposure_eur']:,.0f}")
    print(f"  Over-licensed:  {s['over_licensed']}   Annual spend on unused seats: EUR {s['annual_waste_eur']:,.0f}")
    print(f"Report saved: {path}")


if __name__ == "__main__":
    main()
