"""Command-line entry point for the final integrated report.

Joins the VDI+AD master report (``VDI_AD_final.xlsx`` from ``run.py``) with
the wass last-login report and the Checkout PC List, producing
``VDI_Laptop_Asset_final.xlsx`` -- the list of people who use BOTH a VDI
and a laptop, enriched with AD info and the laptop's Responsible person.

Examples
--------
    python run_match.py                           # auto-discover everything
    python run_match.py -v                        # verbose logging
    python run_match.py --keep-all-vdi            # keep VDI rows w/o laptop
    python run_match.py --lastlogon-file X.xlsx   # explicit lastlogon path
    python run_match.py --vdi-ad-file Y.xlsx      # explicit VDI+AD path
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from vdi_report import config
from vdi_report.final_report import FinalReportBuilder


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=(
            "Build the final VDI + laptop + asset report. Joins "
            "VDI_AD_final.xlsx with the wass last-login report (by account) "
            "and the Checkout PC List (by hostname)."
        )
    )
    p.add_argument(
        "--vdi-ad-file",
        default=None,
        help=(
            "Path to VDI_AD_final.xlsx. Default: <output-dir>/"
            f"{config.VDI_AD_FINAL_FILENAME}."
        ),
    )
    p.add_argument(
        "--asset-root",
        default=config.ASSET_REPORT_ROOT,
        help="Asset report root containing YYYY-MM-DD subfolders.",
    )
    p.add_argument(
        "--lastlogon-file",
        default=None,
        help=(
            "Explicit path to the wass last-login .xlsx. Default: auto-find "
            "the latest 'VDI-laptop-lastlogin-*.xlsx' in the latest asset "
            "YYYY-MM-DD folder (prefers *_merged.xlsx)."
        ),
    )
    p.add_argument(
        "--checkout-file",
        default=None,
        help=(
            "Explicit path to the Checkout PC List .xlsx. Default: auto-find "
            "'Checkout PC List *.xlsx' in the same folder as the lastlogon "
            "file."
        ),
    )
    p.add_argument(
        "--output-dir",
        default=None,
        help=(
            "Directory to write the final report. Defaults to "
            "FINAL_REPORT_ROOT (BMW 'Status gengerate report' share)."
        ),
    )
    p.add_argument(
        "--keep-all-vdi",
        action="store_true",
        help="Deprecated: LEFT join is now always on (all VDI rows kept).",
    )
    p.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Increase logging verbosity (-v info, -vv debug).",
    )
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)

    level = logging.WARNING - 10 * args.verbose
    logging.basicConfig(
        level=max(level, logging.DEBUG),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
    )

    builder = FinalReportBuilder(
        vdi_ad_path=Path(args.vdi_ad_file) if args.vdi_ad_file else None,
        asset_root=Path(args.asset_root),
        lastlogon_path=Path(args.lastlogon_file) if args.lastlogon_file else None,
        checkout_path=Path(args.checkout_file) if args.checkout_file else None,
        output_dir=Path(args.output_dir) if args.output_dir else None,
        vdi_ad_dir=Path(args.output_dir) if args.output_dir else None,
    )
    df = builder.run()
    if df is None:
        print("No report generated. Check the logs above.", file=sys.stderr)
        return 1
    print(f"Done. {len(df)} rows written to {builder.output_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
