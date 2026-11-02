# IT Asset Repository & ISO 27001 Compliance Toolkit

![tests](https://github.com/sidd-24/IT-asset-compliance-tracker/actions/workflows/tests.yml/badge.svg)

An Excel-based IT asset register with Python automation for the reviews an IT or ISMS team runs every week, and for the evidence an auditor asks for:

- **Asset register:** owners, lifecycle (Active, Near Expiry, Expired) and compliance status, with a dashboard that recalculates itself from today's date.
- **Compliance issues:** remediation deadlines by severity, with overdue and due-soon tracking.
- **End-of-life check:** flags operating systems and firmware that no longer get security updates, using the public [endoflife.date](https://endoflife.date) API.
- **License reconciliation:** seats purchased vs seats in use, showing audit risk (under-licensed) and wasted spend (over-licensed).
- **ISO/IEC 27001:2022 evidence pack:** a Statement of Applicability for asset-management controls, with automated evidence for each control.
- **Weekly and monthly audit reports**, plus an audit log of every run.

> Many organisations, especially small and mid-sized companies, still run asset and license governance in Excel. This toolkit keeps Excel as the interface everyone can open, and automates the checks behind it.

Looking for a conversational version? [it-ops-compliance-assistant](https://github.com/sidd-24/it-ops-compliance-assistant) answers questions about the same domain through an AI assistant with role-based tools.

## Workbook

`IT_Asset_Repository.xlsx` contains:

| Sheet | Contents |
|---|---|
| Dashboard | Live formulas: asset counts by compliance and lifecycle state, open / due-soon / overdue issues by severity |
| Asset_Inventory | Assets with owner, dates, compliance status and platform (for the EOL check) |
| Ownership_Register | Asset owners and departments |
| Compliance_Issues | Findings with severity, owner, due date and the ISO 27001 control they belong to |
| License_Entitlements | Seats purchased, seats assigned, cost per seat, source of the assignment count |
| Audit_Log | One line per report run |

The demo data is generated relative to today, so expiry dates and deadlines always produce a realistic mix of states.

## Scripts

| Script | What it does | Output |
|---|---|---|
| `create_repository.py` | Creates the demo workbook (`--force` to overwrite an existing one) | `IT_Asset_Repository.xlsx` |
| `compliance_checker.py` | Console summary: lifecycle as of today, non-compliant assets, overdue remediation; `--update` writes refreshed statuses back | console |
| `generate_reports.py` | Weekly and/or monthly audit report (`--type weekly\|monthly\|both`), appends to the audit log | `reports/Weekly_Audit_<date>.xlsx`, `reports/Monthly_Audit_<yyyy-mm>.xlsx` |
| `eol_check.py` | Checks each asset's platform version on endoflife.date; `--raise-issues` adds A.8.8 findings for unsupported versions; `--offline` uses the local cache | `reports/EOL_Check_<date>.xlsx` |
| `license_reconciliation.py` | Seats purchased vs assigned per license | `reports/License_Reconciliation_<date>.xlsx` |
| `evidence_pack.py` | ISO 27001 Statement of Applicability, control evidence, findings by control, EOL and license results in one workbook | `reports/ISO27001_Evidence_Pack_<date>.xlsx` |

### Rules (in `config.py`)

| Rule | Value |
|---|---|
| Near expiry | warranty/license ends within 90 days |
| Remediation deadline | High 30 days · Medium 60 · Low 90 |
| Due soon | deadline within 14 days |
| EOL warning | vendor support ends within 180 days |
| Over-licensed | fewer than 80% of purchased seats in use |

Lifecycle status and deadline state are **computed from the dates every time a script runs**. Reports never rely on a status saved in the workbook that may be out of date.

## ISO 27001 evidence pack

The Statement of Applicability covers the Annex A controls most relevant to IT asset management. Where possible, each control is assessed automatically:

| Control | Automated evidence |
|---|---|
| A.5.9 Inventory of assets | Assets without owner, without expiry date, or not audited in 12 months |
| A.5.23 Cloud services | SaaS subscriptions that are non-compliant, pending review or expired |
| A.5.30 ICT readiness for business continuity | Servers and network devices out of warranty or expiring |
| A.5.32 Intellectual property rights | License reconciliation: under-licensed products and true-up exposure |
| A.7.14 Secure disposal or re-use | Retired hardware without a recorded wipe/disposal certificate |
| A.8.1 User endpoint devices | Laptops/desktops non-compliant or out of warranty |
| A.8.8 Technical vulnerabilities | Platforms out of vendor support (endoflife.date) |
| A.5.10, A.5.11, A.8.7, A.8.9 | Manual review, plus any open findings mapped to the control |

Each control is rated **Effective**, **Needs Improvement**, **Not Effective** or **Manual Review**. An overdue High-severity finding makes a control Not Effective. Any open finding makes it at least Needs Improvement.

## Example run (demo data)

```
python license_reconciliation.py
  Under-licensed: 2   True-up exposure: EUR 4,860
  Over-licensed:  2   Annual spend on unused seats: EUR 12,668

python eol_check.py
  End of Life            5      e.g. Windows 10 22H2, Cisco IOS XE 17.3, RHEL 7, Ubuntu 20.04
  EOL within 180 days    2
  Supported              3

python evidence_pack.py
  A.5.32  Not Effective      Intellectual property rights
  A.8.8   Not Effective      Management of technical vulnerabilities
  A.5.9   Needs Improvement  Inventory of information and other associated assets
  ...
```

The EOL results depend on the date you run the check, because vendor support dates are fixed.

## Getting started

```bash
pip install -r requirements.txt
python create_repository.py          # demo workbook (add --force to recreate)
python compliance_checker.py
python generate_reports.py --type both
python eol_check.py --raise-issues
python license_reconciliation.py
python evidence_pack.py
pytest -v                            # 31 tests, no network needed
```

To use your own data, replace the rows in the workbook sheets and keep the column headers. To check an asset's platform, fill in `Platform (endoflife.date)` with the product id from endoflife.date (e.g. `windows-server`, `ubuntu`, `cisco-ios-xe`, `rhel`, `macos`) and `Platform Version` with the release (e.g. `2019`, `22.04`, `17.3`).

## Fixes in this version

- **"Overdue remediation"** was described but never computed. It is now calculated from the due dates and shown in the checker, the reports and the dashboard.
- **Reports used stored lifecycle statuses**, which went stale unless `--update` was run first. Statuses are now recomputed on every run.
- **Assets without an expiry date were reported as "Active".** They are now flagged as "Unknown Expiry".
- **Monthly statistics compared only the month number**, so June 2025 counted as June 2026. They now compare the year too.
- **`--type both` never logged the monthly report** in the audit log. Both are logged now.
- **Two dashboard percentages divided by the header cell**, and the "Due soon" and "Overdue" columns were empty. Both are fixed, and the dashboard counts now use `TODAY()`.
- **`create_repository.py` silently overwrote an existing workbook.** It now requires `--force`.
- **The checker's console output crashed** on Windows terminals using the cp1252 encoding.

## License

MIT, see [LICENSE](LICENSE).
