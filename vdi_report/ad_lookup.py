"""Step 2: load the AD Nameanddepartment report and join it with VDI data.

The AD report is keyed by ``Account``, which corresponds to the VDI report's
``Assigned Users`` field. We enrich each VDI row with ``DepartmentCode``,
``Name`` and ``EmailAddress``.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import List, Optional

import pandas as pd

from .config import AD_EXPECTED_COLUMNS, AD_KEEP_COLUMNS

logger = logging.getLogger(__name__)

# Characters that may separate multiple accounts in an "Assigned Users" cell.
_USER_SEPARATORS = ",;\n"


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_ad_report(
    path: Path, keep_columns: Optional[List[str]] = None
) -> pd.DataFrame:
    """Read the AD report and keep only the needed columns.

    Duplicate accounts are collapsed to the first occurrence so the subsequent
    join does not fan out VDI rows.
    """
    keep_columns = list(keep_columns or AD_KEEP_COLUMNS)
    logger.info("Reading AD report: %s", path)
    df = pd.read_excel(path)
    df = _normalize_columns(df)

    have_lower = {str(c).lower(): c for c in df.columns}
    missing = [c for c in AD_EXPECTED_COLUMNS if c.lower() not in have_lower]
    if missing:
        logger.warning("AD report is missing expected columns: %s", missing)

    chosen = [have_lower[c.lower()] for c in keep_columns if c.lower() in have_lower]
    df = df[chosen]

    account_col = have_lower.get("account")
    if account_col and account_col in df.columns:
        before = len(df)
        df = df.drop_duplicates(subset=[account_col], keep="first")
        logger.info("AD report loaded: %d rows (deduped from %d)", len(df), before)
    else:
        logger.warning("AD report has no 'Account' column; dedupe skipped")
        logger.info("AD report loaded: %d rows", len(df))
    return df


def explode_assigned_users(
    vdi_df: pd.DataFrame, column: str = "Assigned Users"
) -> pd.DataFrame:
    """Split multi-user ``Assigned Users`` cells into one row per user.

    VDI cells often contain several accounts separated by commas, semicolons or
    newlines. Exploding them lets us look up AD info for each individual user.
    """
    if column not in vdi_df.columns:
        logger.warning("Column '%s' not in VDI dataframe; skipping explode", column)
        return vdi_df

    pattern = "[" + re.escape(_USER_SEPARATORS) + "]"
    df = vdi_df.copy()
    df[column] = df[column].astype(str).str.strip()
    df[column] = df[column].str.split(pattern)
    df = df.explode(column, ignore_index=True)
    df[column] = df[column].str.strip()
    df = df[
        df[column].notna()
        & (df[column] != "")
        & (~df[column].str.lower().isin(["nan", "none"]))
    ]
    logger.info("After exploding '%s': %d rows", column, len(df))
    return df


def join_vdi_with_ad(
    vdi_df: pd.DataFrame,
    ad_df: pd.DataFrame,
    vdi_key: str = "Assigned Users",
    ad_key: str = "Account",
) -> pd.DataFrame:
    """Left-join VDI data with AD info on ``Assigned Users == Account``.

    Matching is case-insensitive and ignores surrounding whitespace. All VDI
    rows are preserved; rows without an AD match keep NaN for the AD columns.
    """
    if vdi_key not in vdi_df.columns:
        logger.error("VDI dataframe missing join key '%s'", vdi_key)
        return vdi_df
    if ad_key not in ad_df.columns:
        logger.error("AD dataframe missing join key '%s'", ad_key)
        return vdi_df

    vdi_df = vdi_df.copy()
    ad_df = ad_df.copy()
    vdi_df["__key__"] = vdi_df[vdi_key].astype(str).str.strip().str.upper()
    ad_df["__key__"] = ad_df[ad_key].astype(str).str.strip().str.upper()

    # Guard against duplicate keys in AD causing fan-out.
    ad_df = ad_df.drop_duplicates(subset=["__key__"], keep="first")

    merged = vdi_df.merge(ad_df, on="__key__", how="left", suffixes=("", "_ad"))
    merged = merged.drop(columns=["__key__"])

    matched = int(merged[ad_key].notna().sum()) if ad_key in merged.columns else 0
    total = len(merged)
    logger.info(
        "Joined VDI + AD: %d rows, %d/%d assigned users matched an AD account",
        total,
        matched,
        total,
    )
    return merged
