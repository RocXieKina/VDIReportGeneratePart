"""Step 1: merge the Z3 and Z4 VDI VMware reports into a single final frame.

Both source files share the VDI VMware report schema. We concatenate them and
keep only the columns needed downstream: ``Id``, ``IPv4 Address`` and
``Assigned Users``.
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

from .config import VDI_EXPECTED_COLUMNS, VDI_KEEP_COLUMNS, Z3_PREFIX, Z4_PREFIX
from .paths import list_excel_files

logger = logging.getLogger(__name__)


def find_z3_z4_files(folder: Path) -> Tuple[Optional[Path], Optional[Path]]:
    """Locate the Z3- and Z4-prefixed Excel files inside *folder*."""
    files = list_excel_files(folder)
    z3 = next(
        (f for f in files if f.name.upper().startswith(Z3_PREFIX.upper())), None
    )
    z4 = next(
        (f for f in files if f.name.upper().startswith(Z4_PREFIX.upper())), None
    )
    if z3 is None:
        logger.warning("No '%s'-prefixed file found in %s", Z3_PREFIX, folder)
    if z4 is None:
        logger.warning("No '%s'-prefixed file found in %s", Z4_PREFIX, folder)
    return z3, z4


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _select_columns(df: pd.DataFrame, wanted: List[str]) -> pd.DataFrame:
    """Select *wanted* columns case-insensitively; warn on missing ones."""
    have_lower = {str(c).lower(): c for c in df.columns}
    chosen = []
    for w in wanted:
        actual = have_lower.get(w.lower())
        if actual is None:
            logger.warning(
                "Column '%s' not found; available columns: %s",
                w,
                list(df.columns),
            )
            continue
        chosen.append(actual)
    return df[chosen]


def merge_vdi_reports(
    z3_path: Optional[Path],
    z4_path: Optional[Path],
    keep_columns: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Read the Z3 and Z4 reports, concatenate them and keep needed columns.

    A ``__source__`` column records which file each row came from (useful for
    debugging; the orchestrator may drop it before writing the final output).
    """
    keep_columns = list(keep_columns or VDI_KEEP_COLUMNS)
    frames: List[pd.DataFrame] = []
    for label, path in (("Z3", z3_path), ("Z4", z4_path)):
        if path is None:
            continue
        logger.info("Reading VDI %s report: %s", label, path)
        df = pd.read_excel(path)
        df = _normalize_columns(df)

        missing = [
            c
            for c in VDI_EXPECTED_COLUMNS
            if c.lower() not in {str(x).lower() for x in df.columns}
        ]
        if missing:
            logger.warning(
                "VDI %s report is missing expected columns: %s", label, missing
            )

        df = _select_columns(df, keep_columns)
        df["__source__"] = label
        frames.append(df)

    if not frames:
        logger.error("No VDI frames to merge (both Z3 and Z4 missing)")
        return pd.DataFrame(columns=keep_columns + ["__source__"])

    merged = pd.concat(frames, ignore_index=True)
    logger.info(
        "Merged VDI report: %d rows (Z3=%d, Z4=%d)",
        len(merged),
        sum(f["__source__"].eq("Z3").sum() for f in frames),
        sum(f["__source__"].eq("Z4").sum() for f in frames),
    )
    return merged
