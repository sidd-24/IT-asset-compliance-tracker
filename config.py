REPO_FILE = "IT_Asset_Repository.xlsx"
REPORTS_FOLDER = "reports"
CACHE_FOLDER = ".cache"

SHEETS = {
    "assets": "Asset_Inventory",
    "owners": "Ownership_Register",
    "issues": "Compliance_Issues",
    "licenses": "License_Entitlements",
    "audit_log": "Audit_Log",
    "dashboard": "Dashboard",
}

LIFECYCLE_STATUS = ["Active", "Near Expiry", "Expired", "Retired", "Unknown Expiry"]
COMPLIANCE_STATUS = ["Compliant", "Non-Compliant", "Pending Review"]
ISSUE_SEVERITY = ["High", "Medium", "Low"]
ISSUE_STATUS = ["Open", "In Progress", "Closed"]
ASSET_TYPES = ["Hardware", "Software"]

HARDWARE_CATEGORIES = ["Laptop", "Desktop", "Server", "Monitor", "Network Device", "Printer", "Other"]
SOFTWARE_CATEGORIES = ["OS License", "App License", "SaaS Subscription", "Security Tool", "Other"]

NEAR_EXPIRY_DAYS = 90

# Remediation deadline by severity, and how early a deadline counts as "due soon"
REMEDIATION_SLA_DAYS = {"High": 30, "Medium": 60, "Low": 90}
DUE_SOON_DAYS = 14

# End-of-life check (https://endoflife.date): flag platforms whose support ends within this window
EOL_WARNING_DAYS = 180
EOL_API_URL = "https://endoflife.date/api/{product}.json"

# License reconciliation: utilisation below this share means paying for unused seats
LICENSE_UNDERUSE_THRESHOLD = 0.8
