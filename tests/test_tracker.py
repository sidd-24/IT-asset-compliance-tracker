"""Tests on the demo repository and synthetic data. No network access needed."""
from datetime import date, timedelta

import pandas as pd
import pytest
from openpyxl import load_workbook

import create_repository
import eol_check
import generate_reports
from compliance_rules import compute_lifecycle_status, deadline_state, refresh_lifecycle, same_month
from config import SHEETS
from iso27001 import assess_controls
from license_reconciliation import reconcile, summarise

TODAY = date(2026, 6, 15)


def d(days):
    return TODAY + timedelta(days=days)


@pytest.fixture
def repo(tmp_path, monkeypatch):
    """Fresh demo repository in a temp folder (scripts use relative paths)."""
    monkeypatch.chdir(tmp_path)
    create_repository.create_repository()
    return tmp_path / "IT_Asset_Repository.xlsx"


# ─── Lifecycle and deadlines ──────────────────────────────────────────────────

@pytest.mark.parametrize("expiry, current, expected", [
    (d(-1), "Active", "Expired"),            # stored "Active" is stale: recomputed
    (d(0), None, "Near Expiry"),
    (d(90), None, "Near Expiry"),
    (d(91), "Expired", "Active"),
    (d(-30), "Retired", "Retired"),
    (None, "Active", "Unknown Expiry"),      # used to be reported as Active
    (pd.NaT, None, "Unknown Expiry"),
])
def test_lifecycle_status(expiry, current, expected):
    assert compute_lifecycle_status(expiry, TODAY, current) == expected


@pytest.mark.parametrize("status, due, expected", [
    ("Open", d(-1), "Overdue"),
    ("In Progress", d(14), "Due Soon"),
    ("Open", d(15), "On Track"),
    ("Closed", d(-100), "Closed"),
    ("Open", None, "No Due Date"),
])
def test_deadline_state(status, due, expected):
    assert deadline_state(status, due, TODAY) == expected


def test_same_month_checks_the_year():
    assert same_month(date(2026, 6, 1), TODAY)
    assert not same_month(date(2025, 6, 1), TODAY)   # the old check only compared the month


def test_refresh_lifecycle_does_not_modify_input():
    assets = pd.DataFrame({"Expiry / Warranty Date": [d(-5)], "Lifecycle Status": ["Active"]})
    assert refresh_lifecycle(assets, TODAY)["Lifecycle Status"].iloc[0] == "Expired"
    assert assets["Lifecycle Status"].iloc[0] == "Active"


# ─── End-of-life check ────────────────────────────────────────────────────────

CYCLES = [
    {"cycle": "2022", "eol": "2031-10-14", "latest": "10.0.20348"},
    {"cycle": "2016", "eol": "2026-09-01"},
    {"cycle": "2012-r2", "eol": "2023-10-10"},
    {"cycle": "17.3", "eol": True},
    {"cycle": "26.1", "eol": False},
]


@pytest.mark.parametrize("version, state", [
    ("2022", "Supported"),
    ("2016", "EOL within 180 days"),
    ("2012-R2", "End of Life"),
    ("17.3.4", "End of Life"),      # patch level maps to its release cycle
    ("26.1", "Supported"),
    ("1999", "Unknown"),
])
def test_eol_classification(version, state):
    assert eol_check.classify_eol(eol_check.find_cycle(CYCLES, version), TODAY)[0] == state


def test_eol_check_skips_retired_and_assets_without_platform():
    assets = pd.DataFrame({
        "Asset ID": ["A", "B", "C"], "Asset Name": ["a", "b", "c"], "Owner ID": ["O1", "O2", "O3"],
        "Lifecycle Status": ["Active", "Retired", "Active"],
        eol_check.PLATFORM_COL: ["windows-server", "windows-server", None],
        eol_check.VERSION_COL: ["2012-r2", "2012-r2", None],
    })
    result = eol_check.check_assets(assets, TODAY, cycles_for=lambda product: CYCLES)
    assert result["Asset ID"].tolist() == ["A"] and result["EOL State"].iloc[0] == "End of Life"


def test_unknown_product_is_reported_not_crashed():
    assets = pd.DataFrame({"Asset ID": ["A"], "Asset Name": ["a"], "Owner ID": ["O"], "Lifecycle Status": ["Active"],
                           eol_check.PLATFORM_COL: ["no-such-product"], eol_check.VERSION_COL: ["1"]})
    result = eol_check.check_assets(assets, TODAY, cycles_for=lambda product: None)
    assert result["EOL State"].iloc[0] == "Unknown" and "not found" in result["Note"].iloc[0]


def test_raise_issues_adds_a88_findings_once(repo):
    results = pd.DataFrame([{"Asset ID": "HW-007", "Asset Name": "x", "Owner ID": "OWN-006", "Platform": "ubuntu",
                             "Version": "20.04", "EOL State": "End of Life", "End of Support": date(2025, 5, 31),
                             "Days Left": None, "Latest Release": None, "Note": ""}])
    assert eol_check.raise_issues(results, TODAY, repo_file=str(repo)) == 1
    assert eol_check.raise_issues(results, TODAY, repo_file=str(repo)) == 0   # no duplicates
    issues = pd.read_excel(repo, sheet_name=SHEETS["issues"])
    new = issues.iloc[-1]
    assert new["ISO 27001 Control"] == "A.8.8" and new["Severity"] == "High" and new["Status"] == "Open"


# ─── License reconciliation ───────────────────────────────────────────────────

def test_license_reconciliation():
    licenses = pd.DataFrame({
        "License ID": ["OK", "UNDER", "OVER"],
        "Seats Purchased": [10, 3, 50], "Seats Assigned": [9, 5, 20],
        "Annual Cost per Seat (EUR)": [100, 2000, 50],
    })
    result = reconcile(licenses).set_index("License ID")
    assert result.loc["OK", "Reconciliation Status"] == "Balanced"
    assert result.loc["UNDER", "Reconciliation Status"] == "Under-licensed"
    assert result.loc["UNDER", "True-up Exposure (EUR)"] == 4000
    assert result.loc["OVER", "Reconciliation Status"] == "Over-licensed"
    assert result.loc["OVER", "Annual Waste (EUR)"] == 1500
    s = summarise(result.reset_index())
    assert s["true_up_exposure_eur"] == 4000 and s["annual_waste_eur"] == 1600


# ─── ISO 27001 assessment ─────────────────────────────────────────────────────

def minimal_assets():
    return pd.DataFrame({
        "Asset ID": ["L1"], "Asset Name": ["Laptop"], "Asset Type": ["Hardware"], "Category": ["Laptop"],
        "Lifecycle Status": ["Active"], "Compliance Status": ["Compliant"], "Owner ID": ["O1"],
        "Last Audited": [d(-10)], "Notes": [""],
    })


def issues_frame(rows):
    cols = ["Issue ID", "Severity", "Status", "Deadline State", "ISO 27001 Control"]
    return pd.DataFrame(rows, columns=cols)


def test_clean_control_is_effective():
    soa = assess_controls(minimal_assets(), issues_frame([]), TODAY).set_index("Control")
    assert soa.loc["A.8.1", "Implementation Status"] == "Effective"
    assert soa.loc["A.5.9", "Implementation Status"] == "Effective"
    assert soa.loc["A.5.10", "Implementation Status"] == "Manual Review"


def test_overdue_high_issue_makes_control_not_effective():
    issues = issues_frame([["ISS-1", "High", "Open", "Overdue", "A.8.1"], ["ISS-2", "Low", "Open", "On Track", "A.5.10"]])
    soa = assess_controls(minimal_assets(), issues, TODAY).set_index("Control")
    assert soa.loc["A.8.1", "Implementation Status"] == "Not Effective"
    assert soa.loc["A.5.10", "Implementation Status"] == "Needs Improvement"   # open issue beats manual review
    assert soa.loc["A.8.1", "Issue IDs"] == "ISS-1"


def test_under_licensing_fails_a532():
    licenses = reconcile(pd.DataFrame({"License ID": ["X"], "Seats Purchased": [1], "Seats Assigned": [2],
                                       "Annual Cost per Seat (EUR)": [10]}))
    soa = assess_controls(minimal_assets(), issues_frame([]), TODAY, licenses=licenses).set_index("Control")
    assert soa.loc["A.5.32", "Implementation Status"] == "Not Effective"


# ─── Workbook and reports ─────────────────────────────────────────────────────

def test_dashboard_formulas_use_total_and_today(repo):
    ws = load_workbook(repo)[SHEETS["dashboard"]]
    percent_formulas = [ws.cell(r, 3).value for r in range(7, 15)]
    assert all("$B$6" in f for f in percent_formulas)        # was B5 (the header) in two rows
    formulas = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    assert "TODAY()" in formulas and "Overdue" in formulas


def test_demo_dates_are_relative_to_today(repo):
    issues = pd.read_excel(repo, sheet_name=SHEETS["issues"])
    states = [deadline_state(s, due.date(), date.today()) for s, due in zip(issues["Status"], issues["Due Date"])]
    assert "Overdue" in states and "On Track" in states


def test_both_reports_log_weekly_and_monthly(repo):
    generate_reports.generate_reports("both")
    log = pd.read_excel(repo, sheet_name=SHEETS["audit_log"])
    assert log["Audit Type"].tail(2).tolist() == ["Weekly", "Monthly"]


def test_create_repository_refuses_to_overwrite(repo, monkeypatch):
    import subprocess
    import sys
    from pathlib import Path

    script = Path(create_repository.__file__)
    result = subprocess.run([sys.executable, str(script)], capture_output=True, text=True, cwd=repo.parent)
    assert result.returncode != 0 and "--force" in (result.stderr + result.stdout)
