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

# Spreadsheet/report file extensions we recognise when scanning a folder.
# Includes .csv because VDI VMware exports (Z3/Z4) are produced as .csv.
EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xlsb", ".xls", ".csv")

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
# These must EXACTLY match the <div class="EmText EmTextClick"> text in the
# wass output-element picker, including the space before the parenthesis
# (verified from the inspected HTML: "Last User (Account)", "Last User (Email)").
WASS_RESULT_ELEMENTS = [
    "Last Logon",
    "Last User",
    "Last User (Account)",
    "Last User (Email)",
]

# Browser to use for automation: "edge" (default on Windows) or "chrome".
WASS_BROWSER = "edge"

# ===========================================================================
# Final report: VDI + AD + lastlogon + Checkout PC List integration
# ===========================================================================

# Filename of the VDI+AD merged report produced by run.py (input to this step).
VDI_AD_FINAL_FILENAME = "VDI_AD_final.xlsx"

# Base filename of the final integrated report. The actual file is written as
# "<base>_<YYYY-MM-DD>_<HHMMSS>.xlsx" so each run produces a unique snapshot
# that can be archived alongside historical reports.
FINAL_REPORT_BASENAME = "VDI_Laptop_Asset_final"

# Default output directory for the final report: BMW shared "Status gengerate
# report" folder. Intermediate files (VDI_AD_final.xlsx) stay in the local
# output/ folder; only the consolidated report goes to the share.
FINAL_REPORT_ROOT = (
    r"\\china.bmw.corp\winfs\Beijing-Data\IT China\FG-CN-6_Infra\Infrastructure"
    r"\09_Support_Center\01_IT_Support\01_Applications\06_Scripts"
    r"\VDI auto report\Status gengerate report"
)

# --- lastlogon report (downloaded by run_asset.py from wass) ---
# The file lands in the asset YYYY-MM-DD folder alongside the Checkout PC List.
# Name pattern: "VDI-laptop-lastlogin-YYYY-MM-DD*.xlsx" (chunk 0),
#               "VDI-laptop-lastlogin-YYYY-MM-DD-02.xlsx" (chunk 1), ...
#               "VDI-laptop-lastlogin-YYYY-MM-DD_merged.xlsx" (merged multi-chunk)
LASTLOGON_FILE_PREFIX = "VDI-laptop-lastlogin-"
LASTLOGON_MERGED_SUFFIX = "_merged"

# Full header of the wass "Last Logon" result report.
LASTLOGON_EXPECTED_COLUMNS = [
    "HOSTNAME",
    "MACHINEID",
    "LOGIN - LAST LOGON",
    "LOGIN - LAST USER",
    "LOGIN - LAST USER (EMAIL)",
    "LOGIN - LAST USER (ACCOUNT)",
]

# Columns kept from the lastlogon report.
# "LOGIN - LAST USER (ACCOUNT)" is the join key to VDI's "Assigned Users" and
# is dropped after the join (it duplicates Assigned Users).
LASTLOGON_KEEP_COLUMNS = [
    "HOSTNAME",
    "MACHINEID",
    "LOGIN - LAST LOGON",
    "LOGIN - LAST USER",
    "LOGIN - LAST USER (EMAIL)",
    "LOGIN - LAST USER (ACCOUNT)",
]

# --- Checkout PC List columns pulled into the final report ---
# "Name" is the join key to lastlogon's "HOSTNAME" and is dropped after the
# join (it duplicates HOSTNAME).
ASSET_RESPONSIBLE_COLUMNS = [
    "Name",
    "Responsible Q Number",
    "Responsible Name",
    "Responsible Company",
    "Responsible Department",
    "Responsible Email",
]

# --- Final report column order ---
# The report is laid out as:
#   [VDI/AD fixed columns] + [laptop-1 columns] + [laptop-2 columns] + ...
# Each laptop "slot" repeats the same set of columns with a "_N" suffix
# (e.g. "HOSTNAME_1", "Responsible Name_1", "HOSTNAME_2", ...). The number
# of slots equals the maximum number of laptops any single VDI user has.

# Columns that appear once per VDI row (from VDI + AD sources).
VDI_FIXED_COLUMNS = [
    "Id",
    "IPv4 Address",
    "Assigned Users",
    "DepartmentCode",
    "Name",
    "EmailAddress",
]

# Columns that repeat once per matched laptop (from lastlogon + Checkout).
# These are laid out in this exact order within each slot.
LAPTOP_COLUMN_TEMPLATE = [
    "HOSTNAME",
    "MACHINEID",
    "LOGIN - LAST LOGON",
    "LOGIN - LAST USER",
    "LOGIN - LAST USER (EMAIL)",
    "Responsible Q Number",
    "Responsible Name",
    "Responsible Company",
    "Responsible Department",
    "Responsible Email",
]

# Set of all column names that are laptop-derived (used internally to know
# which columns aggregate vs. which take "first" when grouping by VDI row).
LAPTOP_DERIVED_COLUMNS = set(LAPTOP_COLUMN_TEMPLATE)


def generate_final_columns(num_laptops: int) -> list:
    """Build the full column list for the final report.

    ``num_laptops`` is the maximum number of laptops matched by any single
    VDI user. Each laptop slot appends the LAPTOP_COLUMN_TEMPLATE columns
    with a "_N" suffix (1-indexed).
    """
    cols = list(VDI_FIXED_COLUMNS)
    for i in range(1, num_laptops + 1):
        for c in LAPTOP_COLUMN_TEMPLATE:
            cols.append(f"{c}_{i}")
    return cols
