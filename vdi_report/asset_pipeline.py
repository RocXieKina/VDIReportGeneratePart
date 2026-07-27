"""Orchestrates the asset -> wass -> download -> merge pipeline.

High-level flow:

  1. Find the latest ``YYYY-MM-DD`` asset folder and read the Checkout PC List
     ``Name`` column.
  2. Split the name list into chunks of <=4999 entries (wass limit). Chunk 0
     keeps the base report name; subsequent chunks get ``-02``, ``-03`` ...
  3. For each chunk, drive the wass Inventory wizard via Selenium, wait for the
     report to finish and download the generated .xlsx into the asset folder.
  4. If more than one chunk was needed, merge all downloaded reports into a
     single ``VDI-laptop-lastlogin-<date>_merged.xlsx``.
"""
from __future__ import annotations

import datetime as dt
import logging
from pathlib import Path
from typing import List, Optional

import pandas as pd

from . import config
from .asset_reader import (
    chunk_names,
    chunk_report_name,
    find_checkout_pc_list_file,
    find_latest_asset_folder,
    load_asset_names,
)
from .wass_automation import WassReportAutomator

logger = logging.getLogger(__name__)


class AssetWassPipeline:
    """Run the asset -> wass -> download/merge pipeline end to end."""

    def __init__(
        self,
        asset_root: Optional[Path] = None,
        download_dir: Optional[Path] = None,
        chunk_size: int = config.WASS_IMPORT_CHUNK_SIZE,
        browser: str = config.WASS_BROWSER,
        headless: bool = False,
        automator: Optional[WassReportAutomator] = None,
    ) -> None:
        self.asset_root = Path(asset_root) if asset_root else Path(config.ASSET_REPORT_ROOT)
        self.chunk_size = chunk_size
        # If a download dir is not given we resolve it after locating the
        # latest asset folder, so downloads land next to the source report.
        self._explicit_download_dir = download_dir
        self.browser = browser
        self.headless = headless
        self._automator = automator

    # ------------------------------------------------------------------ #
    def run(self, today: Optional[dt.date] = None) -> List[Path]:
        """Execute the full pipeline. Returns the list of downloaded files."""
        logger.info("=== Asset -> wass pipeline started ===")
        today = today or dt.date.today()

        asset_folder = find_latest_asset_folder(self.asset_root, today=today)
        if asset_folder is None:
            logger.error("No asset folder found; aborting")
            return []

        checkout_file = find_checkout_pc_list_file(asset_folder)
        if checkout_file is None:
            logger.error("No Checkout PC List file in %s; aborting", asset_folder)
            return []

        names = load_asset_names(checkout_file)
        if not names:
            logger.error("No computer names extracted from %s; aborting", checkout_file)
            return []

        download_dir = self._explicit_download_dir or asset_folder
        download_dir.mkdir(parents=True, exist_ok=True)

        base_name = f"VDI-laptop-lastlogin-{today.isoformat()}"
        chunks = chunk_names(names, self.chunk_size)
        logger.info(
            "Report base name: %s | %d names -> %d chunk(s) of <=%d",
            base_name,
            len(names),
            len(chunks),
            self.chunk_size,
        )

        downloaded: List[Path] = []
        own_automator = self._automator is None
        automator = self._automator or WassReportAutomator(
            download_dir=download_dir,
            browser=self.browser,
            headless=self.headless,
        )
        try:
            if own_automator:
                automator.start()
                automator.ensure_logged_in()
            for i, chunk in enumerate(chunks):
                report_name = chunk_report_name(base_name, i)
                path = automator.create_report(report_name, chunk)
                if path:
                    downloaded.append(path)
                else:
                    logger.error("chunk %d ('%s') did not produce a download", i, report_name)
        finally:
            if own_automator:
                automator.close()

        if len(downloaded) > 1:
            merged = self._merge_reports(downloaded, download_dir, base_name)
            if merged:
                logger.info("merged report: %s", merged)
                downloaded.append(merged)

        logger.info("=== Asset -> wass pipeline finished: %d file(s) ===", len(downloaded))
        return downloaded

    # ------------------------------------------------------------------ #
    def _merge_reports(
        self,
        paths: List[Path],
        out_dir: Path,
        base_name: str,
    ) -> Optional[Path]:
        """Concatenate downloaded wass reports into one Excel file.

        Keeps the header of the first file and appends data rows from the rest.
        """
        frames: List[pd.DataFrame] = []
        for p in paths:
            try:
                # Use shared read_report so .csv downloads are also merged.
                from .paths import read_report
                df = read_report(p)
                frames.append(df)
            except Exception as e:
                logger.warning("could not read %s for merge: %s", p, e)
        if not frames:
            return None
        merged = pd.concat(frames, ignore_index=True)
        out_path = out_dir / f"{base_name}_merged.xlsx"
        merged.to_excel(out_path, index=False)
        return out_path
