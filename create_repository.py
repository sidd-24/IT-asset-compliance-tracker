import os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from compliance_rules import compute_lifecycle_status
from config import DUE_SOON_DAYS, REPO_FILE, SHEETS

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

# ─── Demo data, relative to today so statuses and deadlines stay meaningful ───

def d(days):
    return date.today() + timedelta(days=days)


# ─── Sheet: Asset Inventory ───────────────────────────────────────────────────

ASSET_HEADERS = [
    "Asset ID", "Asset Name", "Asset Type", "Category",
    "Vendor / Manufacturer", "Model / Version",
    "Serial / License Number", "Purchase Date",
    "Expiry / Warranty Date", "Lifecycle Status",
    "Department", "Location", "Owner ID",
    "Compliance Status", "Last Audited", "Notes",
    "Platform (endoflife.date)", "Platform Version",
]


def demo_assets():
    """Rows without Lifecycle Status; it is computed from the expiry date when the sheet is built."""
    return [
        ["HW-001", "Dell Latitude 5540",       "Hardware", "Laptop",         "Dell",        "Latitude 5540",   "SN-DL001", d(-1000), d(60),   "IT",         "HQ-Floor1",   "OWN-001", "Compliant",      d(-40),  "Annual audit done",                     "windows",        "11-23h2-e"],
        ["HW-002", "HP ProLiant DL380",        "Hardware", "Server",         "HP",          "DL380 Gen10",     "SN-HP002", d(-1600), d(400),  "IT",         "Data Center", "OWN-006", "Compliant",      d(-40),  "Primary app server",                    "windows-server", "2016"],
        ["HW-003", "Cisco Catalyst 9300",      "Hardware", "Network Device", "Cisco",       "Cat 9300",        "SN-CS003", d(-1500), d(45),   "IT",         "Data Center", "OWN-006", "Pending Review", d(-90),  "Firmware update needed",                "cisco-ios-xe",   "17.3"],
        ["HW-004", "Dell OptiPlex 7090",       "Hardware", "Desktop",        "Dell",        "OptiPlex 7090",   "SN-DL004", d(-1800), d(-300), "Finance",    "HQ-Floor2",   "OWN-003", "Non-Compliant",  d(-60),  "Warranty expired, needs replacement",   "windows",        "10-22h2"],
        ["HW-005", "HP LaserJet Pro M404",     "Hardware", "Printer",        "HP",          "M404dn",          "SN-HP005", d(-500),  d(600),  "HR",         "HQ-Floor3",   "OWN-004", "Compliant",      d(-40),  "",                                      "",               ""],
        ["HW-006", "Lenovo ThinkPad T14",      "Hardware", "Laptop",         "Lenovo",      "ThinkPad T14",    "SN-LN006", d(-700),  d(300),  "Operations", "Remote",      "OWN-005", "Compliant",      d(-40),  "Remote worker device",                  "windows",        "11-24h2-e"],
        ["HW-007", "Dell PowerEdge R740",      "Hardware", "Server",         "Dell",        "PowerEdge R740",  "SN-DL007", d(-1100), d(900),  "IT",         "Data Center", "OWN-006", "Compliant",      d(-40),  "Database server",                       "ubuntu",         "22.04"],
        ["HW-008", "Dell PowerEdge R630",      "Hardware", "Server",         "Dell",        "PowerEdge R630",  "SN-DL008", d(-2600), d(-400), "IT",         "Storage",     "OWN-006", "Compliant",      d(-100), "Retired. Disk wipe certificate WC-2291 on file", "",       ""],
        ["HW-009", "Lenovo ThinkPad X1 (2019)", "Hardware", "Laptop",        "Lenovo",      "X1 Carbon G7",    "SN-LN009", d(-2200), d(-500), "Operations", "Storage",     "OWN-005", "Pending Review", d(-120), "Retired. Returned by leaver, awaiting disposal", "",       ""],
        ["HW-010", "Apple MacBook Pro 14",     "Hardware", "Laptop",         "Apple",       "MacBook Pro M3",  "SN-AP010", d(-400),  d(330),  "Marketing",  "HQ-Floor2",   "OWN-004", "Compliant",      d(-40),  "Design team",                           "macos",          "14"],
        ["HW-011", "Legacy ERP App Server",    "Hardware", "Server",         "Dell",        "PowerEdge R720",  "SN-DL011", d(-3000), d(-30),  "Finance",    "Data Center", "OWN-006", "Non-Compliant",  d(-200), "Runs legacy ERP; migration planned",    "rhel",           "7"],
        ["HW-012", "Jump Host VM",             "Hardware", "Server",         "VMware",      "VM",              "VM-0012",  d(-900),  None,    "IT",         "Data Center", "",        "Pending Review", d(-420), "Owner left the company",                "ubuntu",         "20.04"],
        ["SW-001", "Microsoft 365 E3",         "Software", "App License",    "Microsoft",   "365 E3",          "LK-MS001", d(-300),  d(150),  "All",        "N/A",         "OWN-001", "Compliant",      d(-40),  "100 seats",                             "",               ""],
        ["SW-002", "Windows Server 2022 Datacenter", "Software", "OS License", "Microsoft", "Server 2022",     "LK-MS002", d(-1200), d(500),  "IT",         "Data Center", "OWN-006", "Compliant",      d(-40),  "96 cores",                              "windows-server", "2022"],
        ["SW-003", "CrowdStrike Falcon",       "Software", "Security Tool",  "CrowdStrike", "Falcon Pro",      "LK-CS003", d(-385),  d(-20),  "IT",         "N/A",         "OWN-007", "Non-Compliant",  d(-15),  "License renewal overdue",               "",               ""],
        ["SW-004", "Jira Software",            "Software", "SaaS Subscription", "Atlassian", "Jira Cloud",     "LK-AT004", d(-550),  d(180),  "IT",         "N/A",         "OWN-001", "Compliant",      d(-40),  "50 users",                              "",               ""],
        ["SW-005", "AutoCAD 2024",             "Software", "App License",    "Autodesk",    "AutoCAD 2024",    "LK-AD005", d(-425),  d(-60),  "Operations", "N/A",         "OWN-005", "Non-Compliant",  d(-50),  "Not renewed, usage flagged",            "",               ""],
        ["SW-006", "Adobe Acrobat Pro",        "Software", "App License",    "Adobe",       "Acrobat Pro DC",  "LK-AB006", d(-460),  d(270),  "Finance",    "N/A",         "OWN-003", "Compliant",      d(-40),  "5 seats",                               "",               ""],
        ["SW-007", "Zoom Business",            "Software", "SaaS Subscription", "Zoom",     "Business Plan",   "LK-ZM007", d(-320),  d(45),   "All",        "N/A",         "OWN-001", "Pending Review", d(-30),  "Renewal review in progress",            "",               ""],
        ["SW-008", "Salesforce Sales Cloud",   "Software", "SaaS Subscription", "Salesforce", "Enterprise",    "LK-SF008", d(-45),   d(320),  "Sales",      "N/A",         "OWN-003", "Compliant",      d(-40),  "25 users",                              "",               ""],
    ]


def build_assets_sheet(ws):
    ws.sheet_view.showGridLines = False

    headers = ASSET_HEADERS
    ws.row_dimensions[1].height = 35
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        style_header(c)

    today = date.today()
    retired = {"HW-008", "HW-009"}
    assets = []
    for row in demo_assets():
        status = "Retired" if row[0] in retired else compute_lifecycle_status(row[8], today)
        assets.append(row[:9] + [status] + row[9:])

    status_fills = {
        "Active":        GREEN_FILL,
        "Near Expiry":   YELLOW_FILL,
        "Expired":       RED_FILL,
        "Retired":       PatternFill("solid", fgColor="D9D9D9"),
        "Unknown Expiry": YELLOW_FILL,
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
                                   else ("9C6500" if val in ("Near Expiry", "Pending Review", "Unknown Expiry") else "9C0006"))

    set_col_widths(ws, {
        "A":10,"B":28,"C":12,"D":18,"E":22,"F":20,
        "G":16,"H":14,"I":18,"J":15,"K":14,"L":14,
        "M":10,"N":16,"O":14,"P":34,"Q":22,"R":14
    })
    freeze(ws)

# ─── Sheet: Compliance Issues ─────────────────────────────────────────────────

def build_issues_sheet(ws):
    ws.sheet_view.showGridLines = False

    headers = [
        "Issue ID", "Asset ID", "Issue Description", "Severity",
        "Identified Date", "Remediation Owner", "Due Date",
        "Status", "Resolution Notes", "Last Updated", "ISO 27001 Control"
    ]

    ws.row_dimensions[1].height = 35
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        style_header(c)

    # Due date = identified date + remediation SLA for the severity (High 30, Medium 60, Low 90 days)
    issues = [
        ["ISS-001","HW-004","Warranty expired. Device running without support coverage.",   "High",   d(-70), "OWN-003", d(-40), "Open",        "",                                d(-20), "A.8.1"],
        ["ISS-002","SW-003","CrowdStrike license not renewed. Endpoints unprotected.",       "High",   d(-18), "OWN-007", d(12),  "In Progress", "Renewal PO raised.",              d(-3),  "A.8.7"],
        ["ISS-003","SW-005","AutoCAD used without valid license. Compliance risk.",          "High",   d(-55), "OWN-005", d(-25), "In Progress", "Usage suspended pending review.", d(-10), "A.5.32"],
        ["ISS-004","HW-003","Cisco switch firmware several versions behind.",                "Medium", d(-50), "OWN-006", d(10),  "Open",        "",                                d(-5),  "A.8.9"],
        ["ISS-005","SW-007","Zoom renewal not actioned. Risk of service disruption.",       "Medium", d(-20), "OWN-001", d(40),  "Open",        "",                                d(-20), "A.5.23"],
        ["ISS-006","HW-001","Laptop warranty expires in <90 days. Plan replacement.",        "Low",    d(-10), "OWN-001", d(80),  "Open",        "",                                d(-10), "A.8.1"],
        ["ISS-007","HW-009","Retired laptop has no disk wipe certificate.",                  "Medium", d(-35), "OWN-005", d(25),  "Open",        "",                                d(-35), "A.7.14"],
        ["ISS-008","SW-001","MFA not enforced for Microsoft 365 admin accounts.",           "High",   d(-90), "OWN-001", d(-60), "Closed",      "Conditional access policy enabled.", d(-65), "A.5.23"],
        ["ISS-009","HW-006","BitLocker disk encryption not enabled.",                        "Medium", d(-40), "OWN-005", d(20),  "Closed",      "Encryption enabled via Intune.",  d(-2),  "A.8.1"],
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
        "A":10,"B":10,"C":46,"D":10,
        "E":15,"F":20,"G":13,"H":13,
        "I":36,"J":13,"K":14
    })
    freeze(ws)

# ─── Sheet: License Entitlements ──────────────────────────────────────────────

def build_licenses_sheet(ws):
    ws.sheet_view.showGridLines = False

    headers = [
        "License ID", "Asset ID", "Product", "Vendor", "License Metric",
        "Seats Purchased", "Seats Assigned", "Annual Cost per Seat (EUR)",
        "Assignment Source", "Last Reconciled"
    ]
    ws.row_dimensions[1].height = 35
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(1, i, h))

    licenses = [
        ["LIC-M365-E3",     "SW-001", "Microsoft 365 E3",                "Microsoft",   "Per user",   100, 87,  432,  "Entra ID license group",  d(-7)],
        ["LIC-WINSRV-DC",   "SW-002", "Windows Server 2022 Datacenter",  "Microsoft",   "Per 16-core pack", 6, 6, 1500, "Hyper-V host inventory",  d(-30)],
        ["LIC-FALCON",      "SW-003", "CrowdStrike Falcon Pro",          "CrowdStrike", "Per device", 120, 131, 60,   "Falcon console sensors", d(-7)],
        ["LIC-JIRA",        "SW-004", "Jira Software Cloud",             "Atlassian",   "Per user",   50,  34,  92,   "Atlassian admin",        d(-14)],
        ["LIC-AUTOCAD",     "SW-005", "AutoCAD 2024",                    "Autodesk",    "Per user",   3,   5,   2100, "Software inventory (Intune)", d(-14)],
        ["LIC-ACROBAT",     "SW-006", "Adobe Acrobat Pro",               "Adobe",       "Per user",   5,   5,   240,  "Adobe Admin Console",    d(-30)],
        ["LIC-ZOOM",        "SW-007", "Zoom Business",                   "Zoom",        "Per host",   40,  22,  200,  "Zoom admin",             d(-14)],
        ["LIC-SALESFORCE",  "SW-008", "Salesforce Sales Cloud",          "Salesforce",  "Per user",   25,  24,  1980, "Salesforce setup",       d(-7)],
    ]
    for r, row in enumerate(licenses, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            style_row(cell, alt=(r % 2 == 0))
            if isinstance(val, date):
                cell.number_format = "DD-MMM-YYYY"

    set_col_widths(ws, {"A": 18, "B": 10, "C": 32, "D": 14, "E": 18, "F": 12, "G": 12, "H": 16, "I": 26, "J": 14})
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
        ["AUD-001", d(-63), "Weekly",  17, 12, 3, 2, 5, 3, "generate_reports.py"],
        ["AUD-002", d(-56), "Weekly",  17, 12, 3, 2, 6, 3, "generate_reports.py"],
        ["AUD-003", d(-49), "Weekly",  18, 12, 3, 3, 6, 3, "generate_reports.py"],
        ["AUD-004", d(-42), "Weekly",  18, 13, 3, 2, 7, 4, "generate_reports.py"],
        ["AUD-005", d(-35), "Monthly", 18, 13, 3, 2, 7, 4, "generate_reports.py"],
        ["AUD-006", d(-28), "Weekly",  19, 13, 3, 3, 7, 3, "generate_reports.py"],
        ["AUD-007", d(-21), "Weekly",  19, 13, 4, 2, 7, 3, "generate_reports.py"],
        ["AUD-008", d(-14), "Weekly",  20, 13, 4, 3, 7, 3, "generate_reports.py"],
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
    ws["A2"].value = "Live formulas: counts recalculate from the sheets and today's date whenever the workbook opens"
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

    # Lifecycle counts are computed from the expiry dates (column I) and TODAY(),
    # so they stay correct without re-running any script. Retired assets (column J) are excluded.
    inv = "Asset_Inventory!"
    not_retired = f'{inv}J2:J1000,"<>Retired"'
    pct = lambda r: f'=IF($B$6=0,"",TEXT(B{r}/$B$6,"0%"))'   # always relative to Total Assets (B6)

    kpi_headers = ["Metric", "Count", "% of Total", "Status"]
    kpi_data = [
        ["Total Assets",           f"=COUNTA({inv}A2:A1000)", "100%", ""],
        ["Hardware Assets",        f'=COUNTIF({inv}C2:C1000,"Hardware")', pct(7), ""],
        ["Software Assets",        f'=COUNTIF({inv}C2:C1000,"Software")', pct(8), ""],
        ["Compliant",              f'=COUNTIF({inv}N2:N1000,"Compliant")', pct(9), "✅ Good"],
        ["Non-Compliant",          f'=COUNTIF({inv}N2:N1000,"Non-Compliant")', pct(10), "🔴 Action Needed"],
        ["Pending Review",         f'=COUNTIF({inv}N2:N1000,"Pending Review")', pct(11), "🟡 Review Due"],
        ["Active (> 90 days)",     f'=COUNTIFS({inv}I2:I1000,">"&(TODAY()+90),{not_retired})', pct(12), ""],
        ["Near Expiry (<= 90 days)", f'=COUNTIFS({inv}I2:I1000,">="&TODAY(),{inv}I2:I1000,"<="&(TODAY()+90),{not_retired})', pct(13), "🟡 Plan Renewal"],
        ["Expired",                f'=COUNTIFS({inv}I2:I1000,"<"&TODAY(),{not_retired})', pct(14), "🔴 Immediate Action"],
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

    for r, (metric, formula, pct_formula, status) in enumerate(kpi_data, 6):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate([metric, formula, pct_formula, status], 1):
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

    iss_headers = ["Severity", "Open", f"Due Soon (<= {DUE_SOON_DAYS} days)", "Overdue"]
    for i, h in enumerate(iss_headers, 1):
        c = ws.cell(row_start + 1, i, h)
        style_subheader(c)

    iss = "Compliance_Issues!"
    def counts(sev):
        open_ = f'{iss}D2:D1000,"{sev}",{iss}H2:H1000,"<>Closed"'
        return [
            sev,
            f"=COUNTIFS({open_})",
            f'=COUNTIFS({open_},{iss}G2:G1000,">="&TODAY(),{iss}G2:G1000,"<="&(TODAY()+{DUE_SOON_DAYS}))',
            f'=COUNTIFS({open_},{iss}G2:G1000,"<"&TODAY())',
        ]
    iss_data = [counts("High"), counts("Medium"), counts("Low")]
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
        (SHEETS["licenses"],  build_licenses_sheet),
        (SHEETS["audit_log"], build_audit_log_sheet),
    ]

    for name, builder in sheet_builders:
        ws = wb.create_sheet(name)
        builder(ws)

    wb.save(REPO_FILE)
    print(f"Repository created: {REPO_FILE}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Create the demo asset repository workbook")
    parser.add_argument("--force", action="store_true", help="Overwrite an existing repository")
    args = parser.parse_args()
    if os.path.exists(REPO_FILE) and not args.force:
        raise SystemExit(f"{REPO_FILE} already exists. Use --force to overwrite it with demo data.")
    create_repository()
