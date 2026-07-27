"""Path and column configuration for the VDI report pipeline."""
from __future__ import annotations

# --- Network paths (Windows UNC) ---
VDI_REPORT_ROOT = (
    r"\\china.bmw.corp\winfs\Beijing-Data\IT China\FG-CN-6_Infra\Infrastructure"
    r"\09_Support_Center\01_IT_Support\01_Applications\06_Scripts"
    r"\VDI auto report\VDI VMware report"
)

AD_REPORT_ROOT = (
    r"\\china.bmw.corp\winfs\Beijing-Data\IT China\FG-CN-6_Infra\Infrastructure"
    r"\09_Support_Center\01_IT_Support\01_Applications\06_Scripts"
    r"\Active Directory\NameanddepartmentDailyReportExport"
)

# --- VDI VMware report columns ---
# Full header list as produced by the VDI auto report (used for validation).
VDI_EXPECTED_COLUMNS = [
    "Id",
    "Power Status",
    "IPv4 Address",
    "vNet/Subnet",
    "VM Model/Hardware",
    "Life Cycle Status",
    "Agent Version",
    "Agent Status",
    "Session Allocation Status",
    "Assigned Users",
    "User Count",
    "Hibernate State",
]

# Columns retained from the merged VDI report.
VDI_KEEP_COLUMNS = ["Id", "IPv4 Address", "Assigned Users"]

# --- AD Nameanddepartment report columns ---
AD_EXPECTED_COLUMNS = [
    "Account",
    "DepartmentCode",
    "Name",
    "Title",
    "Office",
    "EmailAddress",
    "whenCreated",
]

# Columns retained from the AD report (Account is the join key).
AD_KEEP_COLUMNS = ["Account", "DepartmentCode", "Name", "EmailAddress"]

# --- File naming ---
Z3_PREFIX = "Z3"
Z4_PREFIX = "Z4"

# Excel extensions we recognise when scanning a folder.
EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xlsb", ".xls")

# How many days back to look for an AD report if the latest day is empty.
AD_MAX_LOOKBACK_DAYS = 30

# ===========================================================================
# Asset report (Checkout PC List) + wass.bmwgroup.net automation
# ===========================================================================

# Asset report root: contains YYYY-MM-DD subfolders, each holding a
# "Checkout PC List MMDD.xlsx" file.
ASSET_REPORT_ROOT = (
    r"\\china.bmw.corp\winfs\Beijing-Data\IT China\FG-CN-6_Infra\Infrastructure"
    r"\09_Support_Center\01_IT_Support\01_Applications\06_Scripts"
    r"\VDI auto report\Asset report"
)

# Full header of the Checkout PC List asset report (used for validation).
ASSET_EXPECTED_COLUMNS = [
    "Name",
    "Item",
    "Status",
    "StatusReason",
    "Manufacturer",
    "Model",
    "S/N",
    "Responsible Q Number",
    "Responsible Name",
    "PurchaseDate",
    "ReadyReturnTime",
    "WarrantyEndDate",
    "Comments",
    "Responsible Company",
    "Responsible Department",
    "Responsible Email",
]

# Only the computer "Name" column is needed for the wass import.
ASSET_KEEP_COLUMNS = ["Name"]

# Filename prefix of the asset report inside each YYYY-MM-DD folder.
ASSET_FILE_PREFIX = "Checkout PC List"

# How many days back to look for an asset report folder.
ASSET_MAX_LOOKBACK_DAYS = 30

# --- wass.bmwgroup.net automation ---
WASS_URL = "https://wass.bmwgroup.net"

# wass "Import computer list" accepts at most 5000 entries. Use 4999 to be safe.
WASS_IMPORT_CHUNK_SIZE = 4999

# Polling interval (seconds) when waiting for report completion.
WASS_POLL_INTERVAL_SECONDS = 10

# Hard timeout (seconds) for waiting for a single report to finish.
WASS_REPORT_TIMEOUT_SECONDS = 1800

# Result elements to select in the wass wizard (under the "login" group).
WASS_RESULT_ELEMENTS = [
    "Last Logon",
    "Last User",
    "Last User(Account)",
    "Last User(Email)",
]

# Browser to use for automation: "edge" (default on Windows) or "chrome".
WASS_BROWSER = "edge"
