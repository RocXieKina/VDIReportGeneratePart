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


def find_latest_vdi_folder(root: Path) -> Optional[Path]:
    """Return the latest time-based subfolder under the VDI report root.

    Subfolder names are assumed to be date/time stamps. We try to parse them
    as dates (so lexicographic and chronological order agree); if parsing fails
    we fall back to filesystem modification time.
    """
    if not root.is_dir():
        logger.error("VDI report root not found or not accessible: %s", root)
        return None

    subdirs = [p for p in root.iterdir() if p.is_dir()]
    if not subdirs:
        logger.error("No time-based subfolders under VDI report root: %s", root)
        return None

    # Try date-name sorting first.
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
            parseable.sort(key=lambda x: x[1])
            latest = parseable[-1][0]
            logger.info("Latest VDI folder (by date in name): %s", latest)
            return latest
    except ImportError:
        logger.debug("dateutil not available; falling back to mtime sorting")

    # Fallback: newest modification time.
    subdirs.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    latest = subdirs[0]
    logger.info("Latest VDI folder (by mtime): %s", latest)
    return latest


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
