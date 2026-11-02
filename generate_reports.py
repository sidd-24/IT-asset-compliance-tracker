"""
generate_reports.py
────────────────────
Run this script weekly or monthly to produce audit reports from the repository.

Usage:
    python generate_reports.py --type weekly
    python generate_reports.py --type monthly
    python generate_reports.py --type both   (default)
"""

import os
import argparse
from datetime import date, timedelta

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from compliance_rules import add_deadline_state, refresh_lifecycle, same_month
from config import REPO_FILE, SHEETS, REPORTS_FOLDER, NEAR_EXPIRY_DAYS

# ─── Styles ───────────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="1F3864")
GREEN_FILL   = PatternFill("solid", fgColor="C6EFCE")
YELLOW_FILL  = PatternFill("solid", fgColor="FFEB9C")
RED_FILL     = PatternFill("solid", fgColor="FFC7CE")
ALT_FILL     = PatternFill("solid", fgColor="EBF3FB")
TITLE_FILL   = PatternFill("solid", fgColor="D6E4F7")

WHITE_BOLD   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
TITLE_FONT   = Font(name="Arial", bold=True, color="1F3864", size=14)
SECTION_FONT = Font(name="Arial", bold=True, color="1F3864", size=11)
NORMAL_FONT  = Font(name="Arial", size=10)
BOLD_FONT    = Font(name="Arial", bold=True, size=10)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

THIN_SIDE   = Side(style="thin",   color="BFBFBF")
THICK_SIDE  = Side(style="medium", color="1F3864")
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEAD_BORDER = Border(left=THICK_SIDE, right=THICK_SIDE, top=THICK_SIDE, bottom=THICK_SIDE)


def h(cell, merge_ws=None, merge_range=None):
    cell.font      = WHITE_BOLD
    cell.fill      = HEADER_FILL
    cell.border    = HEAD_BORDER
    cell.alignment = CENTER


def data_cell(cell, alt=False, fill=None):
    cell.font      = NORMAL_FONT
    cell.fill      = fill if fill else (ALT_FILL if alt else PatternFill("solid", fgColor="FFFFFF"))
    cell.border    = CELL_BORDER
    cell.alignment = LEFT


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ─── Load Data ────────────────────────────────────────────────────────────────

def load_data(today=None):
    """Assets with lifecycle status recomputed for today, and issues with their deadline state."""
    today  = today or date.today()
    assets = pd.read_excel(REPO_FILE, sheet_name=SHEETS["assets"])
    issues = pd.read_excel(REPO_FILE, sheet_name=SHEETS["issues"])

    date_cols_a = ["Purchase Date", "Expiry / Warranty Date", "Last Audited"]
    date_cols_i = ["Identified Date", "Due Date", "Last Updated"]

    for col in date_cols_a:
        if col in assets.columns:
            assets[col] = pd.to_datetime(assets[col], errors="coerce").dt.date

    for col in date_cols_i:
        if col in issues.columns:
            issues[col] = pd.to_datetime(issues[col], errors="coerce").dt.date

    return refresh_lifecycle(assets, today), add_deadline_state(issues, today)


# ─── Shared Helpers ───────────────────────────────────────────────────────────

def get_compliance_fill(status):
    return {
        "Compliant":     GREEN_FILL,
        "Non-Compliant": RED_FILL,
        "Pending Review":YELLOW_FILL,
    }.get(status, PatternFill("solid", fgColor="FFFFFF"))


def get_lifecycle_fill(status):
    return {
        "Active":      GREEN_FILL,
        "Near Expiry": YELLOW_FILL,
        "Expired":     RED_FILL,
    }.get(status, PatternFill("solid", fgColor="FFFFFF"))


def get_severity_fill(sev):
    return {"High": RED_FILL, "Medium": YELLOW_FILL, "Low": GREEN_FILL}.get(
        sev, PatternFill("solid", fgColor="FFFFFF")
    )


def write_section_title(ws, row, text, col_span=8):
    ws.row_dimensions[row].height = 26
    ws.merge_cells(f"A{row}:{get_column_letter(col_span)}{row}")
    cell = ws.cell(row, 1, text)
    cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    cell.fill      = PatternFill("solid", fgColor="2E75B6")
    cell.alignment = CENTER


def write_kpi_row(ws, row, label, value, note=""):
    ws.row_dimensions[row].height = 22
    lc = ws.cell(row, 1, label)
    lc.font   = BOLD_FONT
    lc.fill   = ALT_FILL
    lc.border = CELL_BORDER
    lc.alignment = LEFT

    vc = ws.cell(row, 2, value)
    vc.font   = Font(name="Arial", bold=True, size=12, color="1F3864")
    vc.fill   = PatternFill("solid", fgColor="FFFFFF")
    vc.border = CELL_BORDER
    vc.alignment = CENTER

    if note:
        nc = ws.cell(row, 3, note)
        nc.font   = Font(name="Arial", size=10, italic=True, color="595959")
        nc.fill   = PatternFill("solid", fgColor="FFFFFF")
        nc.border = CELL_BORDER
        nc.alignment = LEFT


# ─── Weekly Report ────────────────────────────────────────────────────────────

def build_weekly_report(wb, assets, issues, today):
    ws = wb.create_sheet("Weekly Audit")
    ws.sheet_view.showGridLines = False

    # Title
    ws.row_dimensions[1].height = 40
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value     = f"Weekly IT Asset Audit Report  |  Week of {today.strftime('%d %b %Y')}"
    t.font      = TITLE_FONT
    t.fill      = TITLE_FILL
    t.alignment = CENTER

    ws.row_dimensions[2].height = 18
    ws.merge_cells("A2:H2")
    ws["A2"].value     = "Generated automatically by generate_reports.py"
    ws["A2"].font      = Font(name="Arial", italic=True, size=9, color="595959")
    ws["A2"].alignment = CENTER

    # ── Section 1: Summary KPIs
    write_section_title(ws, 4, "SUMMARY")

    open_issues = issues[issues["Status"] != "Closed"]
    high_open   = len(open_issues[open_issues["Severity"] == "High"])
    near_expiry = assets[assets["Lifecycle Status"] == "Near Expiry"]
    expired     = assets[assets["Lifecycle Status"] == "Expired"]
    non_compl   = assets[assets["Compliance Status"] == "Non-Compliant"]

    kpis = [
        ("Total Assets",            len(assets),                        ""),
        ("Compliant",               len(assets[assets["Compliance Status"] == "Compliant"]),  ""),
        ("Non-Compliant",           len(non_compl),                     "Immediate action required"),
        ("Pending Review",          len(assets[assets["Compliance Status"] == "Pending Review"]), ""),
        ("Near Expiry (< 90 days)", len(near_expiry),                   "Plan renewal"),
        ("Expired",                 len(expired),                       "Action overdue"),
        ("Open Issues",             len(open_issues),                   ""),
        ("Overdue Remediation",     len(open_issues[open_issues["Deadline State"] == "Overdue"]), "Past the remediation deadline"),
        ("High Severity Open",      high_open,                          "Escalate if not actioned"),
    ]

    ws.cell(5, 1, "Metric").font  = WHITE_BOLD
    ws.cell(5, 1).fill            = HEADER_FILL
    ws.cell(5, 1).border          = HEAD_BORDER
    ws.cell(5, 1).alignment       = CENTER
    ws.cell(5, 2, "Value").font   = WHITE_BOLD
    ws.cell(5, 2).fill            = HEADER_FILL
    ws.cell(5, 2).border          = HEAD_BORDER
    ws.cell(5, 2).alignment       = CENTER
    ws.cell(5, 3, "Notes").font   = WHITE_BOLD
    ws.cell(5, 3).fill            = HEADER_FILL
    ws.cell(5, 3).border          = HEAD_BORDER
    ws.cell(5, 3).alignment       = CENTER

    for i, (label, val, note) in enumerate(kpis, 6):
        write_kpi_row(ws, i, label, val, note)

    row = 6 + len(kpis) + 2

    # ── Section 2: New & Open Compliance Issues
    write_section_title(ws, row, "OPEN COMPLIANCE ISSUES")
    row += 1

    iss_cols = ["Issue ID", "Asset ID", "Issue Description", "Severity",
                "Remediation Owner", "Due Date", "Status", "Deadline State"]

    ws.row_dimensions[row].height = 30
    for i, col in enumerate(iss_cols, 1):
        c = ws.cell(row, i, col)
        h(c)
    row += 1

    open_iss = open_issues[iss_cols] if all(c in open_issues.columns for c in iss_cols) else open_issues

    for r_idx, (_, row_data) in enumerate(open_iss.iterrows()):
        ws.row_dimensions[row].height = 22
        for c_idx, col in enumerate(iss_cols, 1):
            val  = row_data.get(col, "")
            cell = ws.cell(row, c_idx, val if not (isinstance(val, float) and pd.isna(val)) else "")
            data_cell(cell, alt=(r_idx % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"
            if col == "Severity":
                cell.fill = get_severity_fill(str(val))
            if col == "Status" and str(val) == "Open":
                cell.fill = RED_FILL
            if col == "Deadline State":
                cell.fill = {"Overdue": RED_FILL, "Due Soon": YELLOW_FILL, "On Track": GREEN_FILL}.get(str(val), cell.fill)
        row += 1

    row += 2

    # ── Section 3: Assets Near Expiry / Expired
    write_section_title(ws, row, "ASSETS REQUIRING ATTENTION (Near Expiry / Expired)")
    row += 1

    attn_cols = ["Asset ID", "Asset Name", "Asset Type", "Lifecycle Status",
                 "Expiry / Warranty Date", "Owner ID", "Compliance Status"]

    ws.row_dimensions[row].height = 30
    for i, col in enumerate(attn_cols, 1):
        c = ws.cell(row, i, col)
        h(c)
    row += 1

    attn = assets[assets["Lifecycle Status"].isin(["Near Expiry", "Expired"])]
    for r_idx, (_, row_data) in enumerate(attn.iterrows()):
        ws.row_dimensions[row].height = 22
        for c_idx, col in enumerate(attn_cols, 1):
            val  = row_data.get(col, "")
            cell = ws.cell(row, c_idx, val if not (isinstance(val, float) and pd.isna(val)) else "")
            data_cell(cell, alt=(r_idx % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"
            if col == "Lifecycle Status":
                cell.fill = get_lifecycle_fill(str(val))
            if col == "Compliance Status":
                cell.fill = get_compliance_fill(str(val))
        row += 1

    set_widths(ws, {"A":12,"B":26,"C":14,"D":18,"E":38,"F":18,"G":14,"H":16})
    ws.freeze_panes = "A3"


# ─── Monthly Report ───────────────────────────────────────────────────────────

def build_monthly_report(wb, assets, issues, today):
    ws = wb.create_sheet("Monthly Audit")
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 40
    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value     = f"Monthly IT Asset Compliance Report  |  {today.strftime('%B %Y')}"
    t.font      = TITLE_FONT
    t.fill      = TITLE_FILL
    t.alignment = CENTER

    ws.row_dimensions[2].height = 18
    ws.merge_cells("A2:I2")
    ws["A2"].value     = "Generated automatically by generate_reports.py"
    ws["A2"].font      = Font(name="Arial", italic=True, size=9, color="595959")
    ws["A2"].alignment = CENTER

    # ── Section 1: Full Asset Inventory Snapshot
    row = 4
    write_section_title(ws, row, "FULL ASSET INVENTORY SNAPSHOT", col_span=9)
    row += 1

    snap_cols = [
        "Asset ID", "Asset Name", "Asset Type", "Category",
        "Lifecycle Status", "Compliance Status", "Owner ID",
        "Expiry / Warranty Date", "Notes"
    ]

    ws.row_dimensions[row].height = 30
    for i, col in enumerate(snap_cols, 1):
        c = ws.cell(row, i, col)
        h(c)
    row += 1

    for r_idx, (_, row_data) in enumerate(assets.iterrows()):
        ws.row_dimensions[row].height = 22
        for c_idx, col in enumerate(snap_cols, 1):
            val  = row_data.get(col, "")
            cell = ws.cell(row, c_idx, val if not (isinstance(val, float) and pd.isna(val)) else "")
            data_cell(cell, alt=(r_idx % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"
            if col == "Lifecycle Status":
                cell.fill = get_lifecycle_fill(str(val))
            if col == "Compliance Status":
                cell.fill = get_compliance_fill(str(val))
        row += 1

    row += 2

    # ── Section 2: Compliance Issues Full Status
    write_section_title(ws, row, "ALL COMPLIANCE ISSUES — FULL STATUS", col_span=9)
    row += 1

    iss_all_cols = [
        "Issue ID", "Asset ID", "Issue Description", "Severity",
        "Identified Date", "Remediation Owner", "Due Date", "Status", "Resolution Notes"
    ]

    ws.row_dimensions[row].height = 30
    for i, col in enumerate(iss_all_cols, 1):
        c = ws.cell(row, i, col)
        h(c)
    row += 1

    for r_idx, (_, row_data) in enumerate(issues.iterrows()):
        ws.row_dimensions[row].height = 22
        for c_idx, col in enumerate(iss_all_cols, 1):
            val  = row_data.get(col, "")
            cell = ws.cell(row, c_idx, val if not (isinstance(val, float) and pd.isna(val)) else "")
            data_cell(cell, alt=(r_idx % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"
            if col == "Severity":
                cell.fill = get_severity_fill(str(val))
            if col == "Status":
                cell.fill = {
                    "Open": RED_FILL, "In Progress": YELLOW_FILL, "Closed": GREEN_FILL
                }.get(str(val), cell.fill)
        row += 1

    row += 2

    # ── Section 3: Monthly Stats
    write_section_title(ws, row, "MONTHLY COMPLIANCE STATISTICS", col_span=9)
    row += 1

    closed_this_month = issues[
        (issues["Status"] == "Closed") &
        (issues["Last Updated"].apply(lambda d: same_month(d, today)))
    ]
    new_this_month = issues[
        issues["Identified Date"].apply(lambda d: same_month(d, today))
    ]

    stats = [
        ("Total Assets",              len(assets)),
        ("Hardware",                  len(assets[assets["Asset Type"] == "Hardware"])),
        ("Software",                  len(assets[assets["Asset Type"] == "Software"])),
        ("Compliance Rate",           f"{(len(assets[assets['Compliance Status'] == 'Compliant']) / max(len(assets),1) * 100):.1f}%"),
        ("Issues Raised This Month",  len(new_this_month)),
        ("Issues Closed This Month",  len(closed_this_month)),
        ("Issues Still Open",         len(issues[issues["Status"] != "Closed"])),
        ("High Severity Open",        len(issues[(issues["Severity"] == "High") & (issues["Status"] != "Closed")])),
        ("Overdue Remediation",       len(issues[issues["Deadline State"] == "Overdue"])),
    ]

    ws.cell(row, 1, "Metric").font  = WHITE_BOLD
    ws.cell(row, 1).fill            = HEADER_FILL
    ws.cell(row, 1).border          = HEAD_BORDER
    ws.cell(row, 1).alignment       = CENTER
    ws.cell(row, 2, "Value").font   = WHITE_BOLD
    ws.cell(row, 2).fill            = HEADER_FILL
    ws.cell(row, 2).border          = HEAD_BORDER
    ws.cell(row, 2).alignment       = CENTER
    row += 1

    for label, val in stats:
        write_kpi_row(ws, row, label, val)
        row += 1

    set_widths(ws, {"A":12,"B":26,"C":14,"D":18,"E":16,"F":18,"G":14,"H":20,"I":36})
    ws.freeze_panes = "A3"


# ─── Update Audit Log ─────────────────────────────────────────────────────────

def update_audit_log(assets, issues, audit_type, today):
    wb  = load_workbook(REPO_FILE)
    ws  = wb[SHEETS["audit_log"]]

    open_issues  = issues[issues["Status"] != "Closed"]
    high_open    = len(open_issues[open_issues["Severity"] == "High"])
    last_row     = ws.max_row + 1
    audit_id     = f"AUD-{(last_row - 1):03d}"

    row = [
        audit_id, today, audit_type,
        len(assets),
        len(assets[assets["Compliance Status"] == "Compliant"]),
        len(assets[assets["Compliance Status"] == "Non-Compliant"]),
        len(assets[assets["Compliance Status"] == "Pending Review"]),
        len(open_issues), high_open,
        "generate_reports.py"
    ]

    ws.append(row)
    ws.cell(last_row, 2).number_format = "DD-MMM-YYYY"
    wb.save(REPO_FILE)
    print(f"Audit log updated: {audit_id}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def generate_reports(report_type="both"):
    os.makedirs(REPORTS_FOLDER, exist_ok=True)
    today  = date.today()
    assets, issues = load_data(today)

    if report_type in ("weekly", "both"):
        wb   = Workbook()
        wb.remove(wb.active)
        build_weekly_report(wb, assets, issues, today)
        fname = os.path.join(REPORTS_FOLDER, f"Weekly_Audit_{today.strftime('%Y-%m-%d')}.xlsx")
        wb.save(fname)
        update_audit_log(assets, issues, "Weekly", today)
        print(f"Weekly report saved: {fname}")

    if report_type in ("monthly", "both"):
        wb   = Workbook()
        wb.remove(wb.active)
        build_monthly_report(wb, assets, issues, today)
        fname = os.path.join(REPORTS_FOLDER, f"Monthly_Audit_{today.strftime('%Y-%m')}.xlsx")
        wb.save(fname)
        update_audit_log(assets, issues, "Monthly", today)   # previously skipped when running "both"
        print(f"Monthly report saved: {fname}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--type", choices=["weekly", "monthly", "both"], default="both")
    args = parser.parse_args()
    generate_reports(args.type)
