"""CLI entry point for the asset -> wass report pipeline.

Examples
--------
    python run_asset.py                       # default network path + Edge
    python run_asset.py -v                    # verbose logging
    python run_asset.py --browser chrome      # use Chrome instead of Edge
    python run_asset.py --headless            # headless (often breaks SSO)
    python run_asset.py --asset-root <dir>    # override asset root (testing)
    python run_asset.py --dry-run             # read+chunk only, skip browser
"""
from __future__ import annotations

import argparse
import datetime as dt
import logging
import sys
from pathlib import Path

from vdi_report import config
from vdi_report.asset_pipeline import AssetWassPipeline
from vdi_report.asset_reader import (
    chunk_names,
    chunk_report_name,
    find_checkout_pc_list_file,
    find_latest_asset_folder,
    load_asset_names,
)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Read latest Checkout PC List, drive wass Inventory report "
        "wizard and download the generated Excel."
    )
    p.add_argument(
        "--asset-root",
        default=config.ASSET_REPORT_ROOT,
        help="Asset report root containing YYYY-MM-DD subfolders.",
    )
    p.add_argument(
        "--download-dir",
        default=None,
        help="Where the generated .xlsx lands. Default: the asset YYYY-MM-DD folder.",
    )
    p.add_argument(
        "--browser",
        default=config.WASS_BROWSER,
        choices=["edge", "chrome"],
        help="Browser for wass automation (default: edge).",
    )
    p.add_argument(
        "--headless",
        action="store_true",
        help="Run the browser headless (may break corporate SSO).",
    )
    p.add_argument(
        "--chunk-size",
        type=int,
        default=config.WASS_IMPORT_CHUNK_SIZE,
        help="Max computer names per wass report (default: 4999).",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Only find+read+chunk the asset report; do not launch the browser.",
    )
    p.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Increase logging verbosity (-v info, -vv debug).",
    )
    return p


def _dry_run(asset_root: Path, chunk_size: int, today: dt.date) -> int:
    folder = find_latest_asset_folder(asset_root, today=today)
    if folder is None:
        print("No asset folder found.", file=sys.stderr)
        return 1
    checkout = find_checkout_pc_list_file(folder)
    if checkout is None:
        print(f"No Checkout PC List file in {folder}", file=sys.stderr)
        return 1
    names = load_asset_names(checkout)
    if not names:
        print("No computer names extracted.", file=sys.stderr)
        return 1
    base = f"VDI-laptop-lastlogin-{today.isoformat()}"
    chunks = chunk_names(names, chunk_size)
    print(f"Asset folder : {folder}")
    print(f"Source file  : {checkout}")
    print(f"Computer names: {len(names)} (deduped)")
    print(f"Report base  : {base}")
    print(f"Chunks       : {len(chunks)} (<= {chunk_size} each)")
    for i, c in enumerate(chunks):
        print(f"  - {chunk_report_name(base, i)}: {len(c)} names")
    return 0


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    level = logging.WARNING - 10 * args.verbose
    logging.basicConfig(
        level=max(level, logging.DEBUG),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
    )

    today = dt.date.today()

    if args.dry_run:
        return _dry_run(Path(args.asset_root), args.chunk_size, today)

    pipeline = AssetWassPipeline(
        asset_root=args.asset_root,
        download_dir=Path(args.download_dir) if args.download_dir else None,
        chunk_size=args.chunk_size,
        browser=args.browser,
        headless=args.headless,
    )
    files = pipeline.run(today=today)
    if not files:
        print("No files generated. Check the logs above.", file=sys.stderr)
        return 1
    print(f"Done. {len(files)} file(s):")
    for f in files:
        print(f"  {f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
