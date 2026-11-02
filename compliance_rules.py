"""
compliance_rules.py
────────────────────
Business rules shared by every script: lifecycle status from the expiry
date, and the deadline state of a compliance issue. Both are computed from
dates when a script runs, so reports never rely on a stored status that
may be out of date.
"""

from datetime import date, timedelta

import pandas as pd

from config import DUE_SOON_DAYS, NEAR_EXPIRY_DAYS


def _as_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)) or value is pd.NaT:
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    return value if isinstance(value, date) else None


def compute_lifecycle_status(expiry_date, today, current_status=None):
    """Active, Near Expiry, Expired, or Unknown Expiry. A Retired asset stays Retired."""
    if str(current_status).strip() == "Retired":
        return "Retired"
    expiry = _as_date(expiry_date)
    if expiry is None:
        return "Unknown Expiry"   # previously treated as "Active", which hid missing data
    if expiry < today:
        return "Expired"
    if expiry <= today + timedelta(days=NEAR_EXPIRY_DAYS):
        return "Near Expiry"
    return "Active"


def refresh_lifecycle(assets: pd.DataFrame, today) -> pd.DataFrame:
    """Returns a copy of the asset table with Lifecycle Status recomputed for today."""
    assets = assets.copy()
    assets["Lifecycle Status"] = [
        compute_lifecycle_status(expiry, today, status)
        for expiry, status in zip(assets["Expiry / Warranty Date"], assets.get("Lifecycle Status", [None] * len(assets)))
    ]
    return assets


def deadline_state(status, due_date, today):
    """Closed, Overdue, Due Soon (within DUE_SOON_DAYS) or On Track."""
    if str(status).strip() == "Closed":
        return "Closed"
    due = _as_date(due_date)
    if due is None:
        return "No Due Date"
    if due < today:
        return "Overdue"
    if due <= today + timedelta(days=DUE_SOON_DAYS):
        return "Due Soon"
    return "On Track"


def add_deadline_state(issues: pd.DataFrame, today) -> pd.DataFrame:
    """Adds 'Deadline State' and 'Days Overdue' columns to the issue table."""
    issues = issues.copy()
    issues["Deadline State"] = [deadline_state(s, d, today) for s, d in zip(issues["Status"], issues["Due Date"])]
    issues["Days Overdue"] = [
        (today - _as_date(d)).days if state == "Overdue" else 0
        for state, d in zip(issues["Deadline State"], issues["Due Date"])
    ]
    return issues


def same_month(value, today):
    """True if value falls in today's calendar month of the same year."""
    d = _as_date(value)
    return d is not None and d.year == today.year and d.month == today.month
