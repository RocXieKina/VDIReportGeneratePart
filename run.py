"""Command-line entry point for the VDI report analyzer.

Examples
--------
    python run.py                       # use default network paths
    python run.py -v                    # verbose logging
    python run.py --no-explode          # keep multi-user cells as-is
    python run.py --vdi-root <dir> --ad-root <dir>
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from vdi_report import VDIReportAnalyzer
from vdi_report import config


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Merge VDI VMware reports (Z3+Z4) with AD Nameanddepartment info."
    )
    parser.add_argument(
        "--vdi-root",
        default=config.VDI_REPORT_ROOT,
        help="VDI VMware report root (default: the BMW share path).",
    )
    parser.add_argument(
        "--ad-root",
        default=config.AD_REPORT_ROOT,
        help="AD NameanddepartmentDailyReportExport root.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(Path.cwd() / "output"),
        help="Directory to write VDI_final.xlsx and VDI_AD_final.xlsx.",
    )
    parser.add_argument(
        "--no-explode",
        action="store_true",
        help="Do not split multi-user 'Assigned Users' cells into multiple rows.",
    )
    parser.add_argument(
        "--keep-source",
        action="store_true",
        help="Keep the internal __source__ (Z3/Z4) column in VDI_final.xlsx.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Increase logging verbosity (-v info, -vv debug).",
    )
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)

    level = logging.WARNING - 10 * args.verbose
    logging.basicConfig(
        level=max(level, logging.DEBUG),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
    )

    analyzer = VDIReportAnalyzer(
        vdi_root=args.vdi_root,
        ad_root=args.ad_root,
        output_dir=args.output_dir,
        explode_users=not args.no_explode,
        keep_vdi_source_col=args.keep_source,
    )
    df = analyzer.run()
    if df is None:
        print("No report generated. Check the logs above.", file=sys.stderr)
        return 1
    print(f"Done. {len(df)} rows written to {args.output_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
