import os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from config import REPO_FILE, SHEETS

# ─── Style Helpers ────────────────────────────────────────────────────────────

HEADER_FILL   = PatternFill("solid", fgColor="1F3864")
SUBHEAD_FILL  = PatternFill("solid", fgColor="2E75B6")
ALT_ROW_FILL  = PatternFill("solid", fgColor="EBF3FB")
WARN_FILL     = PatternFill("solid", fgColor="FFE699")
DANGER_FILL   = PatternFill("solid", fgColor="FF7070")
GREEN_FILL    = PatternFill("solid", fgColor="C6EFCE")
YELLOW_FILL   = PatternFill("solid", fgColor="FFEB9C")
RED_FILL      = PatternFill("solid", fgColor="FFC7CE")

WHITE_BOLD  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
DARK_BOLD   = Font(name="Arial", bold=True, color="1F3864", size=10)
NORMAL_FONT = Font(name="Arial", size=10)
TITLE_FONT  = Font(name="Arial", bold=True, color="1F3864", size=14)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

THIN  = Side(style="thin",   color="BFBFBF")
THICK = Side(style="medium", color="1F3864")
CELL_BORDER  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_BORDER = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)


def style_header(cell, center=True):
    cell.font      = WHITE_BOLD
    cell.fill      = HEADER_FILL
    cell.border    = HEADER_BORDER
    cell.alignment = CENTER if center else LEFT


def style_subheader(cell):
    cell.font      = WHITE_BOLD
    cell.fill      = SUBHEAD_FILL
    cell.border    = CELL_BORDER
    cell.alignment = CENTER


def style_row(cell, alt=False):
    cell.font      = NORMAL_FONT
    cell.fill      = ALT_ROW_FILL if alt else PatternFill("solid", fgColor="FFFFFF")
    cell.border    = CELL_BORDER
    cell.alignment = LEFT


def set_col_widths(ws, widths: dict):
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def freeze(ws, cell="A2"):
    ws.freeze_panes = cell

# ─── Sheet: Ownership Register ────────────────────────────────────────────────

def build_owners_sheet(ws):
    ws.sheet_view.showGridLines = False

    headers = ["Owner ID", "Owner Name", "Email", "Department", "Role", "Manager Name", "Phone"]
    ws.row_dimensions[1].height = 30
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        style_header(c)

    owners = [
        ["OWN-001", "Alice Johnson",  "alice.j@company.com",   "IT",         "IT Manager",         "Bob Smith",    "+1-555-0101"],
        ["OWN-002", "Bob Smith",      "bob.s@company.com",     "IT",         "IT Director",        "Carol White",  "+1-555-0102"],
        ["OWN-003", "Carol White",    "carol.w@company.com",   "Finance",    "Finance Manager",    "Bob Smith",    "+1-555-0103"],
        ["OWN-004", "David Lee",      "david.l@company.com",   "HR",         "HR Manager",         "Bob Smith",    "+1-555-0104"],
        ["OWN-005", "Emma Davis",     "emma.d@company.com",    "Operations", "Ops Lead",           "Bob Smith",    "+1-555-0105"],
        ["OWN-006", "Frank Miller",   "frank.m@company.com",   "IT",         "Systems Admin",      "Alice Johnson","+1-555-0106"],
        ["OWN-007", "Grace Chen",     "grace.c@company.com",   "IT",         "Security Analyst",   "Alice Johnson","+1-555-0107"],
    ]

    for r, row in enumerate(owners, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            style_row(cell, alt=(r % 2 == 0))

    set_col_widths(ws, {"A": 12, "B": 20, "C": 28, "D": 14, "E": 20, "F": 20, "G": 16})
    freeze(ws)

# ─── Sheet: Asset Inventory ───────────────────────────────────────────────────

def build_assets_sheet(ws):
    ws.sheet_view.showGridLines = False

    headers = [
        "Asset ID", "Asset Name", "Asset Type", "Category",
        "Vendor / Manufacturer", "Model / Version",
        "Serial / License Number", "Purchase Date",
        "Expiry / Warranty Date", "Lifecycle Status",
        "Department", "Location", "Owner ID",
        "Compliance Status", "Last Audited", "Notes"
    ]

    ws.row_dimensions[1].height = 35
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        style_header(c)

    today = date.today()
    assets = [
        ["HW-001","Dell Latitude 5540",        "Hardware","Laptop",        "Dell",       "Latitude 5540","SN-DL001-2022", date(2022,3,15), date(2025,3,15), "Near Expiry","IT","HQ-Floor1","OWN-001","Compliant",     date(2025,4,1), "Annual audit done"],
        ["HW-002","HP ProLiant DL380",          "Hardware","Server",        "HP",         "DL380 Gen10",  "SN-HP002-2021", date(2021,6,1),  date(2026,6,1),  "Active",     "IT","Data Center","OWN-006","Compliant",     date(2025,4,1), "Primary app server"],
        ["HW-003","Cisco Catalyst 9300",        "Hardware","Network Device","Cisco",      "Cat 9300",     "SN-CS003-2020", date(2020,9,10), date(2025,9,10), "Near Expiry","IT","Data Center","OWN-006","Pending Review",date(2025,3,1), "Firmware update needed"],
        ["HW-004","Dell OptiPlex 7090",         "Hardware","Desktop",       "Dell",       "OptiPlex 7090","SN-DL004-2021", date(2021,1,20), date(2024,1,20), "Expired",    "Finance","HQ-Floor2","OWN-003","Non-Compliant", date(2025,2,1), "Warranty expired, needs replacement"],
        ["HW-005","HP LaserJet Pro M404",       "Hardware","Printer",       "HP",         "M404dn",       "SN-HP005-2023", date(2023,7,5),  date(2026,7,5),  "Active",     "HR","HQ-Floor3","OWN-004","Compliant",     date(2025,4,1), ""],
        ["HW-006","Lenovo ThinkPad T14",        "Hardware","Laptop",        "Lenovo",     "ThinkPad T14", "SN-LN006-2023", date(2023,2,28), date(2026,2,28), "Active",     "Operations","Remote","OWN-005","Compliant",     date(2025,4,1), "Remote worker device"],
        ["HW-007","Dell PowerEdge R740",        "Hardware","Server",        "Dell",       "PowerEdge R740","SN-DL007-2022",date(2022,11,10),date(2027,11,10),"Active",     "IT","Data Center","OWN-006","Compliant",     date(2025,4,1), "Database server"],
        ["SW-001","Microsoft 365 E3",           "Software","App License",   "Microsoft",  "365 E3",       "LK-MS001-2024", date(2024,1,1),  date(2025,12,31),"Active",     "All","N/A","OWN-001","Compliant",     date(2025,4,1), "100 seats"],
        ["SW-002","Windows Server 2022",        "Software","OS License",    "Microsoft",  "Server 2022",  "LK-MS002-2021", date(2021,6,1),  date(2026,6,1),  "Active",     "IT","Data Center","OWN-006","Compliant",     date(2025,4,1), "Datacenter license"],
        ["SW-003","CrowdStrike Falcon",         "Software","Security Tool", "CrowdStrike","Falcon Pro",   "LK-CS003-2024", date(2024,3,1),  date(2025,3,1),  "Near Expiry","IT","N/A","OWN-007","Non-Compliant", date(2025,3,15),"License renewal overdue"],
        ["SW-004","Jira Software",              "Software","SaaS Subscription","Atlassian","Jira Cloud",  "LK-AT004-2023", date(2023,5,1),  date(2026,5,1),  "Active",     "IT","N/A","OWN-001","Compliant",     date(2025,4,1), "50 users"],
        ["SW-005","AutoCAD 2024",               "Software","App License",   "Autodesk",   "AutoCAD 2024", "LK-AD005-2024", date(2024,1,15), date(2025,1,15), "Expired",    "Operations","N/A","OWN-005","Non-Compliant", date(2025,2,1), "Not renewed, usage flagged"],
        ["SW-006","Adobe Acrobat Pro",          "Software","App License",   "Adobe",      "Acrobat Pro DC","LK-AB006-2023",date(2023,8,1),  date(2026,8,1),  "Active",     "Finance","N/A","OWN-003","Compliant",     date(2025,4,1), "5 seats"],
        ["SW-007","Zoom Business",              "Software","SaaS Subscription","Zoom",    "Business Plan","LK-ZM007-2024", date(2024,6,1),  today+timedelta(days=45),"Near Expiry","All","N/A","OWN-001","Pending Review",date(2025,3,30),"Renewal review in progress"],
    ]

    status_fills = {
        "Active":        GREEN_FILL,
        "Near Expiry":   YELLOW_FILL,
        "Expired":       RED_FILL,
        "Retired":       PatternFill("solid", fgColor="D9D9D9"),
        "Compliant":     GREEN_FILL,
        "Non-Compliant": RED_FILL,
        "Pending Review":YELLOW_FILL,
    }

    for r, row in enumerate(assets, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            style_row(cell, alt=(r % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"
            col_header = headers[c - 1]
            if col_header in ("Lifecycle Status", "Compliance Status") and val in status_fills:
                cell.fill   = status_fills[val]
                cell.font   = Font(name="Arial", size=10, bold=True,
                                   color="375623" if "Compliant" in val or val == "Active"
                                   else ("9C6500" if "Near Expiry" in val or "Pending" in val else "9C0006"))

    set_col_widths(ws, {
        "A":10,"B":26,"C":12,"D":18,"E":22,"F":20,
        "G":22,"H":14,"I":18,"J":15,"K":14,"L":14,
        "M":10,"N":16,"O":14,"P":30
    })
    freeze(ws)

# ─── Sheet: Compliance Issues ─────────────────────────────────────────────────

def build_issues_sheet(ws):
    ws.sheet_view.showGridLines = False

    headers = [
        "Issue ID", "Asset ID", "Issue Description", "Severity",
        "Identified Date", "Remediation Owner", "Due Date",
        "Status", "Resolution Notes", "Last Updated"
    ]

    ws.row_dimensions[1].height = 35
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        style_header(c)

    issues = [
        ["ISS-001","HW-004","Warranty expired. Device running without support coverage.","High",   date(2024,1,21),"OWN-003",date(2025,5,30),"Open",        "",                         date(2025,4,1)],
        ["ISS-002","SW-003","CrowdStrike license not renewed. Endpoint unprotected.",  "High",   date(2025,3,16),"OWN-007",date(2025,4,30),"In Progress", "Renewal PO raised.",       date(2025,4,2)],
        ["ISS-003","SW-005","AutoCAD used without valid license. Compliance risk.",    "High",   date(2025,2,2), "OWN-005",date(2025,4,15),"In Progress", "Usage suspended pending review.", date(2025,3,20)],
        ["ISS-004","HW-003","Cisco switch firmware 2 versions behind.",               "Medium", date(2025,3,2), "OWN-006",date(2025,5,15),"Open",        "",                         date(2025,4,1)],
        ["ISS-005","SW-007","Zoom renewal not actioned. Risk of service disruption.", "Medium", date(2025,3,31),"OWN-001",date(2025,5,10),"Open",        "",                         date(2025,3,31)],
        ["ISS-006","HW-001","Laptop warranty expires in <90 days. Plan replacement.", "Low",    date(2025,4,1), "OWN-001",date(2025,6,1), "Open",        "",                         date(2025,4,1)],
    ]

    sev_fills  = {"High": RED_FILL, "Medium": YELLOW_FILL, "Low": GREEN_FILL}
    stat_fills = {"Open": RED_FILL, "In Progress": YELLOW_FILL, "Closed": GREEN_FILL}
    sev_fonts  = {
        "High":   Font(name="Arial", size=10, bold=True, color="9C0006"),
        "Medium": Font(name="Arial", size=10, bold=True, color="9C6500"),
        "Low":    Font(name="Arial", size=10, bold=True, color="375623"),
    }

    for r, row in enumerate(issues, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            style_row(cell, alt=(r % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"
            col_header = headers[c - 1]
            if col_header == "Severity" and val in sev_fills:
                cell.fill = sev_fills[val]
                cell.font = sev_fonts[val]
            if col_header == "Status" and val in stat_fills:
                cell.fill = stat_fills[val]

    set_col_widths(ws, {
        "A":10,"B":10,"C":42,"D":10,
        "E":15,"F":20,"G":13,"H":13,
        "I":36,"J":13
    })
    freeze(ws)

# ─── Sheet: Audit Log ─────────────────────────────────────────────────────────

def build_audit_log_sheet(ws):
    ws.sheet_view.showGridLines = False

    headers = [
        "Audit ID", "Audit Date", "Audit Type",
        "Total Assets", "Compliant", "Non-Compliant", "Pending Review",
        "Open Issues", "Critical Issues", "Generated By"
    ]

    ws.row_dimensions[1].height = 35
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        style_header(c)

    log = [
        ["AUD-001", date(2025,2,3),  "Weekly",  12, 9, 2, 1, 4, 2, "generate_reports.py"],
        ["AUD-002", date(2025,2,10), "Weekly",  12, 9, 2, 1, 4, 2, "generate_reports.py"],
        ["AUD-003", date(2025,2,17), "Weekly",  13, 9, 2, 2, 5, 2, "generate_reports.py"],
        ["AUD-004", date(2025,2,24), "Weekly",  13, 9, 3, 1, 5, 3, "generate_reports.py"],
        ["AUD-005", date(2025,3,1),  "Monthly", 13, 9, 3, 1, 5, 3, "generate_reports.py"],
        ["AUD-006", date(2025,3,3),  "Weekly",  14, 10,3, 1, 5, 3, "generate_reports.py"],
        ["AUD-007", date(2025,3,10), "Weekly",  14, 10,3, 1, 5, 3, "generate_reports.py"],
        ["AUD-008", date(2025,3,17), "Weekly",  14, 10,3, 1, 6, 3, "generate_reports.py"],
        ["AUD-009", date(2025,3,24), "Weekly",  14, 10,3, 1, 6, 3, "generate_reports.py"],
        ["AUD-010", date(2025,4,1),  "Monthly", 14, 9, 3, 2, 6, 3, "generate_reports.py"],
    ]

    type_fills = {
        "Weekly":  PatternFill("solid", fgColor="DAE9F8"),
        "Monthly": PatternFill("solid", fgColor="C6EFCE"),
    }

    for r, row in enumerate(log, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            style_row(cell, alt=(r % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"
            if headers[c - 1] == "Audit Type":
                cell.fill = type_fills.get(val, cell.fill)

    set_col_widths(ws, {
        "A":12,"B":14,"C":12,"D":14,"E":12,
        "F":16,"G":16,"H":13,"I":16,"J":22
    })
    freeze(ws)

# ─── Sheet: Dashboard ─────────────────────────────────────────────────────────

def build_dashboard_sheet(ws):
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = "IT Asset Repository — Compliance Dashboard"
    title_cell.font  = Font(name="Arial", bold=True, size=16, color="1F3864")
    title_cell.fill  = PatternFill("solid", fgColor="EBF3FB")
    title_cell.alignment = CENTER
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:H2")
    ws["A2"].value = f"Last Refreshed: {date.today().strftime('%d %b %Y')}  |  Run generate_reports.py to update"
    ws["A2"].font  = Font(name="Arial", italic=True, size=10, color="595959")
    ws["A2"].alignment = CENTER
    ws["A2"].fill  = PatternFill("solid", fgColor="F2F2F2")

    # KPI Header
    ws.merge_cells("A4:H4")
    ws["A4"].value = "KEY METRICS"
    ws["A4"].font  = WHITE_BOLD
    ws["A4"].fill  = HEADER_FILL
    ws["A4"].alignment = CENTER
    ws.row_dimensions[4].height = 26

    kpi_headers = ["Metric", "Count", "% of Total", "Status"]
    kpi_data = [
        ["Total Assets",           "=COUNTA(Asset_Inventory!A2:A1000)",
         "100%", ""],
        ["Hardware Assets",        "=COUNTIF(Asset_Inventory!C2:C1000,\"Hardware\")",
         '=IF(B6=0,"",TEXT(B7/B6,"0%"))', ""],
        ["Software Assets",        "=COUNTIF(Asset_Inventory!C2:C1000,\"Software\")",
         '=IF(B5=0,"",TEXT(B8/B6,"0%"))', ""],
        ["Compliant",              "=COUNTIF(Asset_Inventory!N2:N1000,\"Compliant\")",
         '=IF(B5=0,"",TEXT(B9/B6,"0%"))', "✅ Good"],
        ["Non-Compliant",          "=COUNTIF(Asset_Inventory!N2:N1000,\"Non-Compliant\")",
         '=IF(B5=0,"",TEXT(B10/B6,"0%"))', "🔴 Action Needed"],
        ["Pending Review",         "=COUNTIF(Asset_Inventory!N2:N1000,\"Pending Review\")",
         '=IF(B5=0,"",TEXT(B11/B6,"0%"))', "🟡 Review Due"],
        ["Active Assets",          "=COUNTIF(Asset_Inventory!J2:J1000,\"Active\")",
         '=IF(B5=0,"",TEXT(B12/B6,"0%"))', ""],
        ["Near Expiry (<90 days)", "=COUNTIF(Asset_Inventory!J2:J1000,\"Near Expiry\")",
         '=IF(B5=0,"",TEXT(B13/B6,"0%"))', "🟡 Plan Renewal"],
        ["Expired",                "=COUNTIF(Asset_Inventory!J2:J1000,\"Expired\")",
         '=IF(B5=0,"",TEXT(B14/B6,"0%"))', "🔴 Immediate Action"],
    ]

    ws.row_dimensions[5].height = 24
    for i, h in enumerate(kpi_headers, 1):
        c = ws.cell(5, i, h)
        style_subheader(c)

    status_map = {
        "✅ Good":         GREEN_FILL,
        "🔴 Action Needed":RED_FILL,
        "🟡 Review Due":   YELLOW_FILL,
        "🟡 Plan Renewal": YELLOW_FILL,
        "🔴 Immediate Action": RED_FILL,
    }

    for r, (metric, formula, pct, status) in enumerate(kpi_data, 6):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate([metric, formula, pct, status], 1):
            cell = ws.cell(r, c, val)
            style_row(cell, alt=(r % 2 == 0))
            if c == 4 and val in status_map:
                cell.fill = status_map[val]
                cell.font = Font(name="Arial", size=10, bold=True)

    # Issues Summary Header
    row_start = len(kpi_data) + 8
    ws.merge_cells(f"A{row_start}:H{row_start}")
    ws[f"A{row_start}"].value = "OPEN COMPLIANCE ISSUES"
    ws[f"A{row_start}"].font  = WHITE_BOLD
    ws[f"A{row_start}"].fill  = HEADER_FILL
    ws[f"A{row_start}"].alignment = CENTER
    ws.row_dimensions[row_start].height = 26

    iss_headers = ["Severity", "Count", "Due Soon (<14 days)", "Overdue"]
    for i, h in enumerate(iss_headers, 1):
        c = ws.cell(row_start + 1, i, h)
        style_subheader(c)

    iss_data = [
        ["High",   '=COUNTIFS(Compliance_Issues!D2:D1000,"High",Compliance_Issues!H2:H1000,"<>Closed")', "", ""],
        ["Medium", '=COUNTIFS(Compliance_Issues!D2:D1000,"Medium",Compliance_Issues!H2:H1000,"<>Closed")', "", ""],
        ["Low",    '=COUNTIFS(Compliance_Issues!D2:D1000,"Low",Compliance_Issues!H2:H1000,"<>Closed")', "", ""],
    ]
    sev_f_map = {"High": RED_FILL, "Medium": YELLOW_FILL, "Low": GREEN_FILL}

    for i, row in enumerate(iss_data, row_start + 2):
        ws.row_dimensions[i].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(i, c, val)
            style_row(cell, alt=(i % 2 == 0))
            if c == 1 and val in sev_f_map:
                cell.fill = sev_f_map[val]
                cell.font = Font(name="Arial", size=10, bold=True)

    set_col_widths(ws, {"A": 28, "B": 12, "C": 22, "D": 16})

# ─── Main Builder ─────────────────────────────────────────────────────────────

def create_repository():
    wb = Workbook()
    wb.remove(wb.active)

    sheet_builders = [
        (SHEETS["dashboard"], build_dashboard_sheet),
        (SHEETS["assets"],    build_assets_sheet),
        (SHEETS["owners"],    build_owners_sheet),
        (SHEETS["issues"],    build_issues_sheet),
        (SHEETS["audit_log"], build_audit_log_sheet),
    ]

    for name, builder in sheet_builders:
        ws = wb.create_sheet(name)
        builder(ws)

    wb.save(REPO_FILE)
    print(f"Repository created: {REPO_FILE}")


if __name__ == "__main__":
    create_repository()
