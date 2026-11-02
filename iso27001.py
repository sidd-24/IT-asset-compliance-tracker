"""
iso27001.py
────────────
ISO/IEC 27001:2022 Annex A controls relevant to IT asset management, each
with an automated check against the repository data where one is possible.

Every assessment returns:
  status    Effective | Needs Improvement | Not Effective | Manual Review
  evidence  one-line summary of what was checked and found
  items     the asset/license/issue IDs behind a negative result
"""

from datetime import timedelta

import pandas as pd

CONTROLS = [
    # (control, title, theme, justification for applicability)
    ("A.5.9", "Inventory of information and other associated assets", "Organizational",
     "The organization maintains an IT asset register with owners and lifecycle dates."),
    ("A.5.10", "Acceptable use of information and other associated assets", "Organizational",
     "Employees use company hardware and software; usage rules must be defined."),
    ("A.5.11", "Return of assets", "Organizational",
     "Devices and licenses are issued to staff and must be returned on exit."),
    ("A.5.23", "Information security for use of cloud services", "Organizational",
     "SaaS subscriptions process company data."),
    ("A.5.30", "ICT readiness for business continuity", "Organizational",
     "Servers and network devices support critical services and need vendor support."),
    ("A.5.32", "Intellectual property rights", "Organizational",
     "Commercial software is used under license terms that must be complied with."),
    ("A.7.14", "Secure disposal or re-use of equipment", "Physical",
     "Retired hardware may hold data and must be wiped before disposal or re-use."),
    ("A.8.1", "User endpoint devices", "Technological",
     "Laptops and desktops are issued to staff and must be supported and compliant."),
    ("A.8.7", "Protection against malware", "Technological",
     "Endpoints and servers run an anti-malware / EDR agent under a valid license."),
    ("A.8.8", "Management of technical vulnerabilities", "Technological",
     "Platforms must receive security updates; unsupported versions are a vulnerability."),
    ("A.8.9", "Configuration management", "Technological",
     "Firmware and configuration baselines are maintained for infrastructure."),
]

ENDPOINT_CATEGORIES = {"Laptop", "Desktop"}
INFRA_CATEGORIES = {"Server", "Network Device"}


def _result(status, evidence, items=()):
    return {"status": status, "evidence": evidence, "items": list(items)}


def _ids(df, column="Asset ID"):
    return df[column].astype(str).tolist()


# ─── Automated checks (assets have lifecycle status refreshed for today) ──────

def check_inventory(assets, today, **_):
    active = assets[assets["Lifecycle Status"] != "Retired"]
    no_owner = active[active["Owner ID"].isna() | (active["Owner ID"].astype(str).str.strip() == "")]
    no_expiry = active[active["Lifecycle Status"] == "Unknown Expiry"]
    audited = pd.to_datetime(active["Last Audited"], errors="coerce").dt.date
    stale = active[audited.isna() | (audited < today - timedelta(days=365))]
    problems = pd.concat([no_owner, no_expiry, stale]).drop_duplicates("Asset ID")
    evidence = (f"{len(active)} active assets: {len(no_owner)} without owner, {len(no_expiry)} without expiry date, "
                f"{len(stale)} not audited in the last 12 months.")
    return _result("Effective" if problems.empty else "Needs Improvement", evidence, _ids(problems))


def check_cloud_services(assets, today, **_):
    saas = assets[(assets["Category"] == "SaaS Subscription") & (assets["Lifecycle Status"] != "Retired")]
    bad = saas[(saas["Compliance Status"] != "Compliant") | (saas["Lifecycle Status"] == "Expired")]
    evidence = f"{len(saas)} SaaS subscriptions: {len(bad)} non-compliant, pending review or expired."
    return _result("Effective" if bad.empty else "Needs Improvement", evidence, _ids(bad))


def check_continuity(assets, today, **_):
    infra = assets[assets["Category"].isin(INFRA_CATEGORIES) & (assets["Lifecycle Status"] != "Retired")]
    expired = infra[infra["Lifecycle Status"] == "Expired"]
    near = infra[infra["Lifecycle Status"] == "Near Expiry"]
    evidence = (f"{len(infra)} servers and network devices: {len(expired)} out of warranty/support, "
                f"{len(near)} expiring within 90 days.")
    status = "Not Effective" if len(expired) else ("Needs Improvement" if len(near) else "Effective")
    return _result(status, evidence, _ids(pd.concat([expired, near])))


def check_licensing(assets, today, licenses=None, **_):
    if licenses is None or licenses.empty:
        return _result("Manual Review", "No license entitlement data available.")
    under = licenses[licenses["Reconciliation Status"] == "Under-licensed"]
    over = licenses[licenses["Reconciliation Status"] == "Over-licensed"]
    evidence = (f"{len(licenses)} licenses reconciled: {len(under)} under-licensed "
                f"(true-up exposure EUR {under['True-up Exposure (EUR)'].sum():,.0f}), {len(over)} over-licensed.")
    return _result("Not Effective" if len(under) else "Effective", evidence, _ids(under, "License ID"))


def check_disposal(assets, today, **_):
    retired_hw = assets[(assets["Lifecycle Status"] == "Retired") & (assets["Asset Type"] == "Hardware")]
    notes = retired_hw["Notes"].fillna("").astype(str).str.lower()
    missing = retired_hw[~notes.str.contains("wipe|certificate|destroyed|sanitis|sanitiz")]
    if retired_hw.empty:
        return _result("Manual Review", "No retired hardware in the register.")
    evidence = f"{len(retired_hw)} retired hardware assets: {len(missing)} without a recorded wipe/disposal certificate."
    return _result("Effective" if missing.empty else "Needs Improvement", evidence, _ids(missing))


def check_endpoints(assets, today, **_):
    endpoints = assets[assets["Category"].isin(ENDPOINT_CATEGORIES) & (assets["Lifecycle Status"] != "Retired")]
    bad = endpoints[(endpoints["Compliance Status"] == "Non-Compliant") | (endpoints["Lifecycle Status"] == "Expired")]
    evidence = f"{len(endpoints)} laptops/desktops: {len(bad)} non-compliant or out of warranty."
    return _result("Effective" if bad.empty else "Needs Improvement", evidence, _ids(bad))


def check_vulnerabilities(assets, today, eol=None, **_):
    if eol is None or eol.empty:
        return _result("Manual Review", "No end-of-life check results available. Run eol_check.py.")
    end_of_life = eol[eol["EOL State"] == "End of Life"]
    soon = eol[eol["EOL State"].str.startswith("EOL within")]
    evidence = (f"{len(eol)} platforms checked against endoflife.date: {len(end_of_life)} out of support, "
                f"{len(soon)} losing support soon.")
    status = "Not Effective" if len(end_of_life) else ("Needs Improvement" if len(soon) else "Effective")
    return _result(status, evidence, _ids(pd.concat([end_of_life, soon])))


AUTOMATED_CHECKS = {
    "A.5.9": check_inventory,
    "A.5.23": check_cloud_services,
    "A.5.30": check_continuity,
    "A.5.32": check_licensing,
    "A.7.14": check_disposal,
    "A.8.1": check_endpoints,
    "A.8.8": check_vulnerabilities,
}

STATUS_ORDER = ["Effective", "Manual Review", "Needs Improvement", "Not Effective"]


def assess_controls(assets, issues, today, licenses=None, eol=None):
    """
    One row per control for the Statement of Applicability. Open and overdue
    compliance issues mapped to a control can only make its status worse:
    an overdue High issue means Not Effective, any open issue at least Needs Improvement.
    """
    rows = []
    control_col = issues["ISO 27001 Control"].astype(str) if "ISO 27001 Control" in issues else pd.Series("", index=issues.index)
    for control, title, theme, justification in CONTROLS:
        check = AUTOMATED_CHECKS.get(control)
        result = check(assets, today, licenses=licenses, eol=eol) if check else \
            _result("Manual Review", "No automated check; review policy and records manually.")

        mapped = issues[control_col == control]
        open_ = mapped[mapped["Status"] != "Closed"]
        overdue = open_[open_["Deadline State"] == "Overdue"]
        status = result["status"]
        if len(overdue[overdue["Severity"] == "High"]):
            status = "Not Effective"
        elif len(open_) and STATUS_ORDER.index(status) < STATUS_ORDER.index("Needs Improvement"):
            status = "Needs Improvement"

        rows.append({
            "Control": control, "Title": title, "Theme": theme, "Applicable": "Yes",
            "Justification": justification, "Implementation Status": status,
            "Automated Evidence": result["evidence"],
            "Affected Items": ", ".join(result["items"][:15]) + (" …" if len(result["items"]) > 15 else ""),
            "Open Issues": len(open_), "Overdue Issues": len(overdue),
            "Issue IDs": ", ".join(open_["Issue ID"].astype(str)),
        })
    return pd.DataFrame(rows)
