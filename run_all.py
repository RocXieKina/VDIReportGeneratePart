"""Unified entry point: run all three VDI report pipelines in sequence.

Pipeline A (run.py)        : merge Z3+Z4 VDI VMware reports with AD info
                             -> output/VDI_AD_final.xlsx
Pipeline B (run_asset.py)  : drive wass.bmwgroup.net wizard to download the
                             laptop last-login report
                             -> asset YYYY-MM-DD folder / VDI-laptop-lastlogin-*.xlsx
Pipeline C (run_match.py)  : three-way join VDI+AD + lastlogon + Checkout PC List
                             -> output/VDI_Laptop_Asset_final.xlsx

Default behaviour (no flags) is the "daily refresh" path the user asked for:

    1. Always run A  (re-merge today's Z3/Z4 + AD).
    2. Run B ONLY if today's lastlogon file does not exist yet. wass needs
       interactive SSO and ~30 minutes, so we skip it when the file is
       already on disk ("目前 lastlogon 的 文件我已经有了").
    3. Always run C  (final integration).

Flags:

    --with-asset         Force-run B even if lastlogon already exists.
    --skip-asset         Never run B (use whatever lastlogon file exists).
    --keep-all-vdi       Pass through to C (LEFT join, keep VDI-only users).
    --skip-c / --skip-a  Skip a stage (mainly for debugging).

This module only orchestrates; the heavy lifting lives in
``VDIReportAnalyzer``, ``AssetWassPipeline`` and ``FinalReportBuilder``.
"""
from __future__ import annotations

import argparse
import datetime as dt
import logging
import sys
from pathlib import Path
from typing import Optional

from vdi_report import config
from vdi_report.analyzer import VDIReportAnalyzer
from vdi_report.final_report import FinalReportBuilder, find_lastlogon_file
from vdi_report.asset_reader import find_latest_asset_folder

# NOTE: ``AssetWassPipeline`` (and thus selenium) is imported lazily inside
# ``run_pipeline_b``. This keeps ``run_all.py`` usable for the daily-refresh
# path (A -> C with B skipped) on machines that don't have selenium / the
# wass automation stack installed.

logger = logging.getLogger("run_all")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Run all three VDI report pipelines (A: VDI+AD, "
        "B: wass last-login, C: final integration) in one go."
    )
    # --- shared paths ---
    p.add_argument(
        "--output-dir",
        default=str(Path.cwd() / "output"),
        help="Where intermediate VDI_AD_final.xlsx lands (pipeline A).",
    )
    p.add_argument(
        "--final-output-dir",
        default=None,
        help="Where the final report lands (pipeline C). Defaults to "
             "FINAL_REPORT_ROOT (BMW 'Status gengerate report' share).",
    )
    p.add_argument(
        "--vdi-root",
        default=config.VDI_REPORT_ROOT,
        help="VDI VMware report root (pipeline A).",
    )
    p.add_argument(
        "--ad-root",
        default=config.AD_REPORT_ROOT,
        help="AD NameanddepartmentDailyReportExport root (pipeline A).",
    )
    p.add_argument(
        "--asset-root",
        default=config.ASSET_REPORT_ROOT,
        help="Asset report root containing YYYY-MM-DD subfolders (B & C).",
    )
    # --- pipeline B (wass) ---
    p.add_argument(
        "--with-asset",
        action="store_true",
        help="Force-run pipeline B (wass) even if today's lastlogon file exists.",
    )
    p.add_argument(
        "--skip-asset",
        action="store_true",
        help="Never run pipeline B (wass). Use whatever lastlogon file exists.",
    )
    p.add_argument(
        "--browser",
        default=config.WASS_BROWSER,
        choices=["edge", "chrome"],
        help="Browser for wass automation (pipeline B).",
    )
    p.add_argument(
        "--headless",
        action="store_true",
        help="Run wass browser headless (may break corporate SSO).",
    )
    p.add_argument(
        "--chunk-size",
        type=int,
        default=config.WASS_IMPORT_CHUNK_SIZE,
        help="Max computer names per wass report (default: 4999).",
    )
    # --- pipeline C ---
    p.add_argument(
        "--keep-all-vdi",
        action="store_true",
        help="Deprecated: LEFT join is now always on (all VDI rows kept). "
             "Kept for backwards compatibility; has no effect.",
    )
    # --- stage skip (debugging) ---
    p.add_argument("--skip-a", action="store_true", help="Skip pipeline A.")
    p.add_argument("--skip-c", action="store_true", help="Skip pipeline C.")
    p.add_argument(
        "-v", "--verbose", action="count", default=0,
        help="Increase logging verbosity (-v info, -vv debug).",
    )
    return p


# --------------------------------------------------------------------------- #
def _today_lastlogon_exists(asset_root: Path, today: dt.date) -> Optional[Path]:
    """Return the path of today's lastlogon file if it exists, else None."""
    folder = find_latest_asset_folder(asset_root, today=today)
    if folder is None:
        return None
    return find_lastlogon_file(folder)


def run_pipeline_a(args: argparse.Namespace) -> bool:
    logger.info("=== Pipeline A: VDI + AD merge ===")
    try:
        analyzer = VDIReportAnalyzer(
            vdi_root=args.vdi_root,
            ad_root=args.ad_root,
            output_dir=args.output_dir,
            write_vdi_final=False,  # skip VDI_final.xlsx; only VDI_AD_final.xlsx is needed
        )
        df = analyzer.run()
    except Exception as e:
        logger.exception("Pipeline A failed: %s", e)
        return False
    if df is None:
        logger.error("Pipeline A produced no data")
        return False
    logger.info("Pipeline A OK: %d rows", len(df))
    return True


def run_pipeline_b(args: argparse.Namespace, today: dt.date) -> bool:
    logger.info("=== Pipeline B: asset -> wass last-login ===")
    try:
        # Lazy import so selenium (and the whole wass stack) is only required
        # when pipeline B actually runs. Skipping B never needs these deps.
        from vdi_report.asset_pipeline import AssetWassPipeline

        pipeline = AssetWassPipeline(
            asset_root=args.asset_root,
            chunk_size=args.chunk_size,
            browser=args.browser,
            headless=args.headless,
        )
        files = pipeline.run(today=today)
    except Exception as e:
        logger.exception("Pipeline B failed: %s", e)
        return False
    if not files:
        logger.error("Pipeline B produced no files")
        return False
    logger.info("Pipeline B OK: %d file(s)", len(files))
    return True


def run_pipeline_c(args: argparse.Namespace) -> bool:
    logger.info("=== Pipeline C: final integration ===")
    try:
        # Don't pass output_dir -- let FinalReportBuilder default to
        # FINAL_REPORT_ROOT (BMW "Status gengerate report" share).
        # --final-output-dir overrides if the caller wants a different path.
        builder = FinalReportBuilder(
            asset_root=args.asset_root,
            output_dir=args.final_output_dir,
            vdi_ad_dir=args.output_dir,
        )
        df = builder.run()
    except Exception as e:
        logger.exception("Pipeline C failed: %s", e)
        return False
    if df is None:
        logger.error("Pipeline C produced no data")
        return False
    logger.info("Pipeline C OK: %d rows", len(df))
    return True


# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    level = logging.WARNING - 10 * args.verbose
    logging.basicConfig(
        level=max(level, logging.DEBUG),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
    )

    today = dt.date.today()
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    results = {"A": None, "B": None, "C": None}

    # --- A ---
    if args.skip_a:
        logger.info("Pipeline A skipped (--skip-a)")
        results["A"] = "skipped"
    else:
        results["A"] = "ok" if run_pipeline_a(args) else "fail"

    # --- B (decide whether to run) ---
    if args.skip_asset and args.with_asset:
        logger.warning("--skip-asset and --with-asset are mutually exclusive; "
                       "honouring --skip-asset")

    run_b = False
    reason = ""
    if args.skip_asset:
        run_b = False
        reason = "skipped by --skip-asset"
    elif args.with_asset:
        run_b = True
        reason = "forced by --with-asset"
    else:
        existing = _today_lastlogon_exists(Path(args.asset_root), today)
        if existing is None:
            run_b = True
            reason = "no lastlogon file found for today"
        else:
            run_b = False
            reason = f"lastlogon already exists ({existing.name})"

    if run_b:
        logger.info("Pipeline B will run: %s", reason)
        results["B"] = "ok" if run_pipeline_b(args, today) else "fail"
    else:
        logger.info("Pipeline B skipped: %s", reason)
        results["B"] = "skipped"

    # --- C ---
    if args.skip_c:
        logger.info("Pipeline C skipped (--skip-c)")
        results["C"] = "skipped"
    else:
        results["C"] = "ok" if run_pipeline_c(args) else "fail"

    # --- summary ---
    print()
    print("=" * 60)
    print("Pipeline summary")
    print("=" * 60)
    for stage, status in results.items():
        label = {"A": "A: VDI+AD merge", "B": "B: wass last-login",
                 "C": "C: final integration"}[stage]
        print(f"  {label:<25} {status}")
    if results["C"] == "ok":
        # Find the most recent timestamped file in the final output dir.
        final_dir = Path(args.final_output_dir or config.FINAL_REPORT_ROOT)
        pattern = f"{config.FINAL_REPORT_BASENAME}_*.xlsx"
        candidates = sorted(final_dir.glob(pattern))
        if candidates:
            print(f"\nFinal report: {candidates[-1]}")
    print("=" * 60)

    # exit non-zero if A or C failed (B is optional)
    failed = [s for s in ("A", "C") if results[s] == "fail"]
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
