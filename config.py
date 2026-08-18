REPO_FILE = "IT_Asset_Repository.xlsx"
REPORTS_FOLDER = "reports"

SHEETS = {
    "assets": "Asset_Inventory",
    "owners": "Ownership_Register",
    "issues": "Compliance_Issues",
    "audit_log": "Audit_Log",
    "dashboard": "Dashboard",
}

LIFECYCLE_STATUS = ["Active", "Near Expiry", "Expired", "Retired"]
COMPLIANCE_STATUS = ["Compliant", "Non-Compliant", "Pending Review"]
ISSUE_SEVERITY = ["High", "Medium", "Low"]
ISSUE_STATUS = ["Open", "In Progress", "Closed"]
ASSET_TYPES = ["Hardware", "Software"]

HARDWARE_CATEGORIES = ["Laptop", "Desktop", "Server", "Monitor", "Network Device", "Printer", "Other"]
SOFTWARE_CATEGORIES = ["OS License", "App License", "SaaS Subscription", "Security Tool", "Other"]

NEAR_EXPIRY_DAYS = 90
