"""VDI Report Analyzer.

Merges VDI VMware reports (Z3 + Z4) with AD Nameanddepartment info to
produce a combined final report keyed by VDI Assigned Users. The
``FinalReportBuilder`` further joins that with the wass last-login report
and the Checkout PC List to identify people who use both a VDI and a laptop.
"""
from .analyzer import VDIReportAnalyzer
from .final_report import FinalReportBuilder

__all__ = ["VDIReportAnalyzer", "FinalReportBuilder"]
__version__ = "0.2.0"
