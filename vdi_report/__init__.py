"""VDI Report Analyzer.

Merges VDI VMware reports (Z3 + Z4) with AD Nameanddepartment info to
produce a combined final report keyed by VDI Assigned Users.
"""
from .analyzer import VDIReportAnalyzer

__all__ = ["VDIReportAnalyzer"]
__version__ = "0.1.0"
