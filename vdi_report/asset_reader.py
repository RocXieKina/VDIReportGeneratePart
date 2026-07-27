"""Read the latest Checkout PC List asset report and extract computer Names.

The Asset report root contains one subfolder per day, named ``YYYY-MM-DD``.
Each subfolder holds a single ``Checkout PC List MMDD.xlsx`` file. We pick the
most recent subfolder (walking backwards from today) that actually contains a
report, then return the ``Name`` column as a list of computer names.
"""
from __future__ import annotations

import datetime as dt
import logging
from pathlib import Path
from typing import List, Optional

import pandas as pd

from .config import (
    ASSET_EXPECTED_COLUMNS,
    ASSET_FILE_PREFIX,
    ASSET_MAX_LOOKBACK_DAYS,
)
from .paths import read_report

logger = logging.getLogger(__name__)


def _is_excel_file(path: Path) -> bool:
    from .config import EXCEL_EXTENSIONS
    return path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS


def find_latest_asset_folder(
    root: Path,
    today: Optional[dt.date] = None,
    max_lookback: int = ASSET_MAX_LOOKBACK_DAYS,
) -> Optional[Path]:
    """Return the most recent ``YYYY-MM-DD`` asset subfolder that has a report.

    Walks backwards day by day from *today*. A day folder that exists but
    contains no Checkout PC List file is skipped (the export may still be
    running), matching the documented "fall back to the previous day" rule.
    """
    if not root.is_dir():
        logger.error("Asset report root not found or not accessible: %s", root)
        return None

    today = today or dt.date.today()
    for offset in range(0, max_lookback + 1):
        day = today - dt.timedelta(days=offset)
        day_folder = root / f"{day.year:04d}-{day.month:02d}-{day.day:02d}"
        if not day_folder.is_dir():
            continue
        if not _has_checkout_file(day_folder):
            logger.info(
                "Asset folder %s has no Checkout PC List file; trying previous day",
                day_folder.name,
            )
            continue
        logger.info("Latest asset folder: %s", day_folder)
        return day_folder

    logger.error(
        "No asset report found in the last %d days under %s", max_lookback, root
    )
    return None


def _has_checkout_file(folder: Path) -> bool:
    return any(
        _is_excel_file(p) and p.name.upper().startswith(ASSET_FILE_PREFIX.upper())
        for p in folder.iterdir()
    )


def find_checkout_pc_list_file(folder: Path) -> Optional[Path]:
    """Locate the ``Checkout PC List *.xlsx`` file inside *folder*."""
    if not folder.is_dir():
        return None
    candidates = sorted(
        p
        for p in folder.iterdir()
        if _is_excel_file(p) and p.name.upper().startswith(ASSET_FILE_PREFIX.upper())
    )
    if not candidates:
        logger.warning("No '%s' file in %s", ASSET_FILE_PREFIX, folder)
        return None
    if len(candidates) > 1:
        logger.warning(
            "Multiple Checkout PC List files in %s; using %s", folder, candidates[-1]
        )
    return candidates[-1]


def load_asset_names(
    path: Path,
    column: str = "Name",
    dedupe: bool = True,
) -> List[str]:
    """Read the asset report and return the computer ``Name`` column.

    Empty / NaN values are dropped. When *dedupe* is True (default) duplicates
    are removed while preserving order, so the wass 5000-entry quota is not
    wasted on repeated computer names.
    """
    logger.info("Reading asset report: %s", path)
    df = read_report(path)
    df.columns = [str(c).strip() for c in df.columns]

    have_lower = {str(c).lower(): c for c in df.columns}
    actual = have_lower.get(column.lower())
    if actual is None:
        logger.error(
            "Column '%s' not found; available columns: %s",
            column,
            list(df.columns),
        )
        return []

    missing = [
        c for c in ASSET_EXPECTED_COLUMNS if c.lower() not in have_lower
    ]
    if missing:
        logger.warning("Asset report is missing expected columns: %s", missing)

    names = (
        df[actual]
        .astype(str)
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        .dropna()
        .tolist()
    )
    if dedupe:
        seen = set()
        unique: List[str] = []
        for n in names:
            key = n.lower()
            if key not in seen:
                seen.add(key)
                unique.append(n)
        logger.info(
            "Asset names loaded: %d (deduped from %d)", len(unique), len(names)
        )
        return unique

    logger.info("Asset names loaded: %d", len(names))
    return names


def chunk_names(
    names: List[str], chunk_size: int
) -> List[List[str]]:
    """Split *names* into chunks of at most *chunk_size* entries."""
    if chunk_size <= 0:
        raise ValueError("chunk_size must be positive")
    return [names[i : i + chunk_size] for i in range(0, len(names), chunk_size)]


def chunk_report_name(base_name: str, index: int) -> str:
    """Return the wass report name for chunk *index*.

    Index 0 keeps the base name (``VDI-laptop-lastlogin-2026-07-27``);
    subsequent chunks get a ``-02``, ``-03`` ... suffix as required.
    """
    if index <= 0:
        return base_name
    return f"{base_name}-{index + 1:02d}"
