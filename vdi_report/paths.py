"""Utilities for locating the latest report folders/files on the network share."""
from __future__ import annotations

import datetime as dt
import logging
from pathlib import Path
from typing import List, Optional

from .config import AD_MAX_LOOKBACK_DAYS, EXCEL_EXTENSIONS

logger = logging.getLogger(__name__)


def _is_excel_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS


def list_excel_files(folder: Path) -> List[Path]:
    """Return sorted Excel files directly inside *folder* (non-recursive)."""
    if not folder.is_dir():
        return []
    return sorted(p for p in folder.iterdir() if _is_excel_file(p))


def find_latest_vdi_folder(
    root: Path,
    require_prefixes: Optional[List[str]] = None,
    max_lookback: int = 30,
) -> Optional[Path]:
    """Return the latest time-based subfolder under the VDI report root.

    Subfolder names are assumed to be date/time stamps. We try to parse them
    as dates (so lexicographic and chronological order agree); if parsing fails
    we fall back to filesystem modification time.

    If *require_prefixes* is given (e.g. ``["Z3", "Z4"]``), only folders that
    contain at least one Excel file whose name starts with one of the prefixes
    (case-insensitive) are considered. This makes the function skip empty
    folders and fall back to older ones -- mirroring
    :func:`find_latest_ad_report`'s "skip empty day, try previous" behaviour.
    Up to *max_lookback* folders (newest first) are inspected.
    """
    if not root.is_dir():
        logger.error("VDI report root not found or not accessible: %s", root)
        return None

    subdirs = [p for p in root.iterdir() if p.is_dir()]
    if not subdirs:
        logger.error("No time-based subfolders under VDI report root: %s", root)
        return None

    # Sort newest-first. Try date-name sorting first; fall back to mtime.
    try:
        from dateutil import parser as dateparser

        def parse_date(p: Path) -> Optional[dt.datetime]:
            try:
                return dateparser.parse(p.name, default=dt.datetime(1900, 1, 1))
            except (ValueError, OverflowError):
                return None

        dated = [(p, parse_date(p)) for p in subdirs]
        parseable = [(p, d) for p, d in dated if d is not None]
        if parseable:
            parseable.sort(key=lambda x: x[1], reverse=True)  # newest first
            ordered = [p for p, _ in parseable]
        else:
            ordered = sorted(subdirs, key=lambda p: p.stat().st_mtime, reverse=True)
    except ImportError:
        logger.debug("dateutil not available; falling back to mtime sorting")
        ordered = sorted(subdirs, key=lambda p: p.stat().st_mtime, reverse=True)

    # No prefix filter -> just return the newest folder (original behaviour).
    if not require_prefixes:
        latest = ordered[0]
        logger.info("Latest VDI folder (by date in name): %s", latest)
        return latest

    # With prefix filter: walk newest -> oldest, return first folder that has
    # at least one matching Excel file. Skip empty folders (the daily VDI
    # auto-report may have failed or not finished writing yet).
    prefixes_upper = [pfx.upper() for pfx in require_prefixes]
    inspected = 0
    for folder in ordered:
        if inspected >= max_lookback:
            break
        inspected += 1
        matching = [
            f for f in list_excel_files(folder)
            if any(f.name.upper().startswith(pfx) for pfx in prefixes_upper)
        ]
        if matching:
            if inspected > 1:
                logger.info(
                    "Latest VDI folder with %s files: %s "
                    "(skipped %d empty newer folder(s))",
                    require_prefixes, folder, inspected - 1,
                )
            else:
                logger.info("Latest VDI folder (by date in name): %s", folder)
            return folder
        logger.info(
            "VDI folder %s has no %s file; trying older folder",
            folder.name, require_prefixes,
        )

    logger.error(
        "No VDI folder with %s files found in the newest %d folders under %s",
        require_prefixes, inspected, root,
    )
    return None


def find_latest_ad_report(
    root: Path,
    today: Optional[dt.date] = None,
    max_lookback: int = AD_MAX_LOOKBACK_DAYS,
) -> Optional[Path]:
    """Find the most recent AD report Excel file under ``root/YYYY/MM/DD``.

    Walks backwards day by day from *today*. A day folder that exists but is
    empty is skipped (the daily export script may still be running), matching
    the documented behaviour of falling back to the previous day.
    """
    if not root.is_dir():
        logger.error("AD report root not found or not accessible: %s", root)
        return None

    today = today or dt.date.today()
    for offset in range(0, max_lookback + 1):
        day = today - dt.timedelta(days=offset)
        day_folder = root / f"{day.year:04d}" / f"{day.month:02d}" / f"{day.day:02d}"
        if not day_folder.is_dir():
            continue
        files = list_excel_files(day_folder)
        if not files:
            logger.info(
                "AD folder empty (export script may still be running); "
                "trying previous day: %s",
                day_folder,
            )
            continue
        logger.info("Latest AD report found for %s: %s", day.isoformat(), files[0])
        return files[0]

    logger.error(
        "No AD report found in the last %d days under %s", max_lookback, root
    )
    return None
