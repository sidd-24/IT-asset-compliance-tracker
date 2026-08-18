# IT Asset Repository & Compliance Tracker

This project helps organizations maintain a central IT asset repository, track ownership, monitor lifecycle status, and flag compliance issues across hardware and software assets.

It combines a structured Excel-based asset register with automated checks for expiry risks, compliance gaps, and overdue remediation actions.

## Overview

The repository stores and manages critical IT asset information including:

- asset inventory
- ownership and department mapping
- hardware and software classifications
- lifecycle state and renewal dates
- compliance status and remediation issues
- audit history and reporting snapshots

The tool is designed to support weekly or monthly governance reviews, asset risk checks, and operational compliance tracking.

## Key Features

- Centralized asset repository in Excel
- Ownership register for asset stewards and departments
- Hardware and software classification support
- Lifecycle status tracking (Active, Near Expiry, Expired, Retired)
- Compliance checks for non-compliant and pending-review items
- Issue tracking for remediation and overdue actions
- Audit log and summary reporting
- Weekly and monthly report generation

## Project Structure

- `config.py` — repository settings, sheet names, asset categories, and lifecycle rules
- `create_repository.py` — creates the main Excel repository structure and seeded data
- `compliance_checker.py` — scans assets and flags lifecycle/compliance drift
- `generate_reports.py` — builds weekly or monthly governance reports
- `IT_Asset_Repository.xlsx` — generated repository workbook
- `reports/` — output report folder

## Repository Model

The workbook includes these sheets:

- Asset Inventory
- Ownership Register
- Compliance Issues
- Audit Log
- Dashboard

These sheets are used to track operational ownership, risk, compliance, and reporting outcomes.

## Prerequisites

Install the required Python packages:

```bash
pip install pandas openpyxl
```

## Configuration

The main configuration lives in `config.py` and includes:

- repository file name
- reports folder
- sheet names
- lifecycle status categories
- compliance categories
- asset categories
- near-expiry threshold

Example settings:

```python
REPO_FILE = "IT_Asset_Repository.xlsx"
REPORTS_FOLDER = "reports"

SHEETS = {
    "assets": "Asset_Inventory",
    "owners": "Ownership_Register",
    "issues": "Compliance_Issues",
    "audit_log": "Audit_Log",
    "dashboard": "Dashboard",
}

NEAR_EXPIRY_DAYS = 90
```

## Create the Repository

To initialize the asset repository workbook:

```bash
python create_repository.py
```

This creates the Excel file defined in `REPO_FILE` and populates the core sheets.

## Run a Compliance Check

To review current asset health and identify lifecycle/compliance issues:

```bash
python compliance_checker.py
```

To apply detected lifecycle updates back to the workbook:

```bash
python compliance_checker.py --update
```

## Generate Reports

To generate a weekly or monthly governance report:

```bash
python generate_reports.py --type weekly
python generate_reports.py --type monthly
```

Or run the default combined report generation:

```bash
python generate_reports.py
```

Reports are saved under the `reports/` directory.

## Typical Use Cases

This project is useful for:

- IT asset inventory management
- software license tracking
- hardware lifecycle monitoring
- compliance auditing
- governance and audit reporting
- risk identification for near-expiry or expired assets

## Example Outcome

The tool can help identify:

- assets nearing end-of-support or warranty expiry
- non-compliant software or licenses
- overdue remediation tasks
- high-severity issues that require escalation
- monthly compliance trends across departments


## License

This project is intended for internal operational use and can be adapted to your organization’s governance and asset management process.
