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
