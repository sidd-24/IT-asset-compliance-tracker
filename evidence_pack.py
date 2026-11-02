"""
evidence_pack.py
─────────────────
Builds an ISO/IEC 27001:2022 audit evidence pack from the asset repository:

  Statement of Applicability   Annex A controls for asset management, applicability,
                               justification and implementation status
  Control Evidence             what was checked automatically and which items failed
  Findings by Control          open compliance issues with deadline state
  EOL Check                    platforms out of vendor support (A.8.8)
  License Reconciliation       seats purchased vs assigned (A.5.32)

Usage:
    python evidence_pack.py              # reports/ISO27001_Evidence_Pack_<date>.xlsx
    python evidence_pack.py --offline    # use cached endoflife.date responses
"""

import argparse
import os
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from compliance_rules import add_deadline_state, refresh_lifecycle
from config import REPO_FILE, REPORTS_FOLDER, SHEETS
from eol_check import check_assets
from generate_reports import (CENTER, GREEN_FILL, RED_FILL, TITLE_FILL, TITLE_FONT, YELLOW_FILL,
                              data_cell, h, write_kpi_row, write_section_title)
from iso27001 import assess_controls
from license_reconciliation import reconcile

STATUS_FILLS = {"Effective": GREEN_FILL, "Needs Improvement": YELLOW_FILL, "Not Effective": RED_FILL}
DEADLINE_FILLS = {"Overdue": RED_FILL, "Due Soon": YELLOW_FILL, "On Track": GREEN_FILL}


def build_inputs(today, offline=False, repo_file=REPO_FILE, cycles_for=None):
    assets = refresh_lifecycle(pd.read_excel(repo_file, sheet_name=SHEETS["assets"]), today)
    issues = add_deadline_state(pd.read_excel(repo_file, sheet_name=SHEETS["issues"]), today)
    licenses = reconcile(pd.read_excel(repo_file, sheet_name=SHEETS["licenses"]))
    eol = check_assets(assets, today, cycles_for=cycles_for, offline=offline)
    return assets, issues, licenses, eol


def _title(ws, text, width):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(f"A1:{get_column_letter(width)}1")
    ws["A1"].value = text
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = TITLE_FONT, TITLE_FILL, CENTER
    ws.row_dimensions[1].height = 36


def _table(ws, df, start_row, columns, widths, fills=None):
    fills = fills or {}
    for i, col in enumerate(columns, 1):
        h(ws.cell(start_row, i, col))
    for r, (_, row) in enumerate(df.iterrows(), start_row + 1):
        ws.row_dimensions[r].height = 30
        for c, col in enumerate(columns, 1):
            value = row.get(col, "")
            if not isinstance(value, str) and pd.isna(value):
                value = None
            cell = ws.cell(r, c, value)
            data_cell(cell, alt=(r % 2 == 0), fill=fills.get(col, {}).get(value))
            if isinstance(value, (date, pd.Timestamp)):
                cell.number_format = "DD-MMM-YYYY"
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(start_row + 1, 1)


def write_pack(soa, assets, issues, licenses, eol, today, path):
    wb = Workbook()

    # ── Statement of Applicability
    ws = wb.active
    ws.title = "Statement of Applicability"
    _title(ws, f"ISO/IEC 27001:2022 Statement of Applicability (asset management scope)  |  {today.strftime('%d %b %Y')}", 9)
    write_section_title(ws, 3, "SUMMARY", col_span=9)
    counts = soa["Implementation Status"].value_counts()
    for i, status in enumerate(["Effective", "Needs Improvement", "Not Effective", "Manual Review"], 4):
        write_kpi_row(ws, i, status, int(counts.get(status, 0)), "controls")
    _table(ws, soa, 9,
           ["Control", "Title", "Theme", "Applicable", "Justification", "Implementation Status",
            "Open Issues", "Overdue Issues", "Issue IDs"],
           [9, 34, 14, 11, 48, 20, 10, 10, 24],
           {"Implementation Status": STATUS_FILLS})

    # ── Control evidence
    ws = wb.create_sheet("Control Evidence")
    _title(ws, "Automated evidence per control", 4)
    _table(ws, soa, 3, ["Control", "Implementation Status", "Automated Evidence", "Affected Items"],
           [9, 20, 80, 50], {"Implementation Status": STATUS_FILLS})

    # ── Findings by control
    ws = wb.create_sheet("Findings by Control")
    _title(ws, "Open compliance issues by control", 9)
    open_issues = issues[issues["Status"] != "Closed"].sort_values(["ISO 27001 Control", "Due Date"])
    _table(ws, open_issues, 3,
           ["ISO 27001 Control", "Issue ID", "Asset ID", "Issue Description", "Severity", "Remediation Owner",
            "Due Date", "Deadline State", "Days Overdue"],
           [12, 10, 10, 50, 10, 14, 14, 14, 10],
           {"Deadline State": DEADLINE_FILLS, "Severity": {"High": RED_FILL, "Medium": YELLOW_FILL, "Low": GREEN_FILL}})

    # ── EOL check and license reconciliation
    ws = wb.create_sheet("EOL Check")
    _title(ws, "Platform end-of-life check (endoflife.date)  |  A.8.8", len(eol.columns))
    _table(ws, eol, 3, list(eol.columns), [10, 28, 10, 16, 10, 20, 14, 10, 14, 30],
           {"EOL State": {"End of Life": RED_FILL, "Supported": GREEN_FILL}})

    ws = wb.create_sheet("License Reconciliation")
    _title(ws, "Software license reconciliation  |  A.5.32", 9)
    _table(ws, licenses, 3,
           ["License ID", "Product", "Seats Purchased", "Seats Assigned", "Utilisation", "Annual Waste (EUR)",
            "True-up Exposure (EUR)", "Reconciliation Status", "Last Reconciled"],
           [34, 24, 12, 12, 11, 14, 16, 20, 14],
           {"Reconciliation Status": {"Under-licensed": RED_FILL, "Over-licensed": YELLOW_FILL, "Balanced": GREEN_FILL}})

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb.save(path)
    return path


def main():
    parser = argparse.ArgumentParser(description="Build an ISO 27001 evidence pack from the asset repository")
    parser.add_argument("--offline", action="store_true", help="Use cached endoflife.date responses only")
    args = parser.parse_args()

    today = date.today()
    assets, issues, licenses, eol = build_inputs(today, offline=args.offline)
    soa = assess_controls(assets, issues, today, licenses=licenses, eol=eol)
    path = write_pack(soa, assets, issues, licenses, eol, today,
                      os.path.join(REPORTS_FOLDER, f"ISO27001_Evidence_Pack_{today.isoformat()}.xlsx"))

    print("\nISO 27001 Statement of Applicability")
    for _, row in soa.iterrows():
        print(f"  {row['Control']:<7} {row['Implementation Status']:<18} {row['Title']}")
    print(f"\nEvidence pack saved: {path}")


if __name__ == "__main__":
    main()
