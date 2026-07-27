"""Step 3: build the final integrated report.

Joins three data sources on top of the VDI+AD master report produced by
``run.py``:

  VDI_AD_final.xlsx  (master, one row per VDI assigned user)
    LEFT JOIN lastlogon.xlsx
        on Assigned Users == LOGIN - LAST USER (ACCOUNT)
        -> adds HOSTNAME / MACHINEID / LOGIN - LAST LOGON / LOGIN - LAST USER
           / LOGIN - LAST USER (EMAIL)
    LEFT JOIN Checkout PC List
        on HOSTNAME == Name
        -> adds Responsible Q Number / Responsible Name / Responsible Company
           / Responsible Department / Responsible Email

UNION semantics (default): every VDI row is preserved as exactly one row in
the output. If a VDI user matches multiple laptops, the laptop-derived
columns are expanded horizontally into one slot per laptop (HOSTNAME_1,
Responsible Name_1, HOSTNAME_2, Responsible Name_2, ...). VDI users
without any laptop match are kept with empty laptop columns. The VDI data
is the unique primary key -- the report never fans out to multiple rows
per VDI entry.

lastlogon rows whose ``LOGIN - LAST USER (ACCOUNT)`` is empty are dropped
before the join (the user said empty ones don't matter).

Output goes to ``FINAL_REPORT_ROOT`` (BMW "Status gengerate report" share)
as ``<basename>_<YYYY-MM-DD>_<HHMMSS>.xlsx`` so each run is a unique
archived snapshot.
"""
from __future__ import annotations

import datetime as dt
import logging
from pathlib import Path
from typing import List, Optional

import pandas as pd

from . import config
from .asset_reader import find_checkout_pc_list_file, find_latest_asset_folder
from .paths import read_report

logger = logging.getLogger(__name__)


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _is_excel_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in config.EXCEL_EXTENSIONS


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _column_map(df: pd.DataFrame) -> dict:
    """Case-insensitive lookup: lower(column) -> actual column name."""
    return {str(c).lower(): c for c in df.columns}


def _clean_key(series: pd.Series) -> pd.Series:
    """Normalise a join-key Series: strip + upper-case, NaN/blank -> NaN."""
    s = series.astype(str).str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA})
    return s.str.upper()


# --------------------------------------------------------------------------- #
# locate input files
# --------------------------------------------------------------------------- #
def find_lastlogon_file(folder: Path) -> Optional[Path]:
    """Find the latest ``VDI-laptop-lastlogin-*.xlsx`` inside *folder*.

    If the multi-chunk ``_merged.xlsx`` is present it is preferred (it already
    contains every chunk concatenated); otherwise the lexicographically last
    matching file is returned.
    """
    if not folder.is_dir():
        return None
    candidates = sorted(
        p
        for p in folder.iterdir()
        if _is_excel_file(p)
        and p.name.upper().startswith(config.LASTLOGON_FILE_PREFIX.upper())
    )
    if not candidates:
        return None
    merged = [
        c for c in candidates
        if c.stem.upper().endswith(config.LASTLOGON_MERGED_SUFFIX.upper())
    ]
    if merged:
        return merged[-1]
    return candidates[-1]


# --------------------------------------------------------------------------- #
# loaders
# --------------------------------------------------------------------------- #
def load_vdi_ad_report(path: Path) -> pd.DataFrame:
    """Read ``VDI_AD_final.xlsx`` (output of run.py)."""
    logger.info("Reading VDI+AD report: %s", path)
    df = read_report(path)
    df = _normalize_columns(df)
    logger.info("VDI+AD report loaded: %d rows, columns=%s", len(df), list(df.columns))
    return df


def load_lastlogon(path: Path) -> pd.DataFrame:
    """Read the wass last-login report and drop rows with empty account.

    The user said "空的就忽略无所谓" -- rows whose
    ``LOGIN - LAST USER (ACCOUNT)`` is blank/NaN are useless for the VDI join
    and are dropped here.
    """
    logger.info("Reading lastlogon report: %s", path)
    df = read_report(path)
    df = _normalize_columns(df)

    cmap = _column_map(df)
    missing = [
        c for c in config.LASTLOGON_EXPECTED_COLUMNS
        if c.lower() not in cmap
    ]
    if missing:
        logger.warning("lastlogon report is missing expected columns: %s", missing)

    account_col = cmap.get("login - last user (account)")
    if account_col is None:
        raise ValueError(
            f"lastlogon file has no 'LOGIN - LAST USER (ACCOUNT)' column; "
            f"available: {list(df.columns)}"
        )

    # Drop rows with empty / NaN account before doing anything else.
    before = len(df)
    df[account_col] = df[account_col].astype(str).str.strip()
    df = df[
        df[account_col].notna()
        & (df[account_col] != "")
        & (~df[account_col].str.lower().isin(["nan", "none", "nat"]))
    ].copy()
    logger.info("lastlogon: dropped %d empty-account rows, %d remain", before - len(df), len(df))

    # Keep only the configured columns.
    chosen = [cmap[c.lower()] for c in config.LASTLOGON_KEEP_COLUMNS if c.lower() in cmap]
    df = df[chosen]

    # De-duplicate exact duplicate rows (a user can legitimately have several
    # laptops, so we DO NOT dedupe by account -- only by full row).
    before = len(df)
    df = df.drop_duplicates()
    if before != len(df):
        logger.info("lastlogon: removed %d exact-duplicate rows", before - len(df))

    logger.info("lastlogon loaded: %d rows", len(df))
    return df


def load_checkout_responsible(path: Path) -> pd.DataFrame:
    """Read Checkout PC List and keep Name + Responsible columns."""
    logger.info("Reading Checkout PC List (responsible columns): %s", path)
    df = read_report(path)
    df = _normalize_columns(df)

    cmap = _column_map(df)
    chosen = [
        cmap[c.lower()]
        for c in config.ASSET_RESPONSIBLE_COLUMNS
        if c.lower() in cmap
    ]
    missing = [
        c for c in config.ASSET_RESPONSIBLE_COLUMNS
        if c.lower() not in cmap
    ]
    if missing:
        logger.warning("Checkout PC List is missing columns: %s", missing)
    df = df[chosen]

    # Dedupe by computer Name (case-insensitive) so the join does not fan out.
    name_col = cmap.get("name")
    if name_col and name_col in df.columns:
        before = len(df)
        df["__dedup_key__"] = _clean_key(df[name_col])
        df = df.drop_duplicates(subset=["__dedup_key__"], keep="first")
        df = df.drop(columns=["__dedup_key__"])
        logger.info("Checkout PC List deduped by Name: %d -> %d", before, len(df))

    logger.info("Checkout PC List loaded: %d rows", len(df))
    return df


# --------------------------------------------------------------------------- #
# joins
# --------------------------------------------------------------------------- #
def join_vdi_with_lastlogon(
    vdi_df: pd.DataFrame,
    lastlogon_df: pd.DataFrame,
    vdi_key: str = "Assigned Users",
    lastlogon_key: str = "LOGIN - LAST USER (ACCOUNT)",
    how: str = "left",
) -> pd.DataFrame:
    """Join VDI+AD data with the lastlogon report.

    Matching is case-insensitive and ignores surrounding whitespace. The
    lastlogon account column is dropped after the join because it duplicates
    VDI's ``Assigned Users``.

    A ``__row_id__`` column is added to the VDI side before the merge so that
    each original VDI row can be uniquely identified later for aggregation
    (a VDI user may match multiple laptops -> multiple rows after merge ->
    must be collapsed back to one row).
    """
    vdi_cmap = _column_map(vdi_df)
    ll_cmap = _column_map(lastlogon_df)

    vdi_key_col = vdi_cmap.get(vdi_key.lower())
    ll_key_col = ll_cmap.get(lastlogon_key.lower())
    if vdi_key_col is None:
        logger.error("VDI dataframe missing join key '%s'", vdi_key)
        return vdi_df
    if ll_key_col is None:
        logger.error("lastlogon dataframe missing join key '%s'", lastlogon_key)
        return vdi_df

    left = vdi_df.copy()
    right = lastlogon_df.copy()

    # Tag each VDI row with a stable id so we can group back after fan-out.
    left["__row_id__"] = range(len(left))

    left["__key__"] = _clean_key(left[vdi_key_col])
    right["__key__"] = _clean_key(right[ll_key_col])

    # Drop the lastlogon account column from the right side -- it duplicates
    # VDI's Assigned Users and would arrive as LOGIN - LAST USER (ACCOUNT).
    right = right.drop(columns=[ll_key_col])

    # Dedupe identical lastlogon rows once more (defensive).
    right = right.drop_duplicates()

    merged = left.merge(right, on="__key__", how=how, suffixes=("", "_ll"))
    merged = merged.drop(columns=["__key__"])

    merged_cmap = _column_map(merged)
    hostname_col = merged_cmap.get("hostname")
    matched = int(merged[hostname_col].notna().sum()) if hostname_col else 0
    logger.info(
        "Joined VDI + lastlogon (%s): %d rows (pre-aggregation), %d VDI rows matched a laptop login",
        how,
        len(merged),
        matched,
    )
    return merged


def join_with_checkout(
    df: pd.DataFrame,
    asset_df: pd.DataFrame,
    df_key: str = "HOSTNAME",
    asset_key: str = "Name",
) -> pd.DataFrame:
    """Left-join with Checkout PC List on HOSTNAME == Name.

    A LEFT join keeps every matched VDI+laptop row even when the laptop is not
    found in the Checkout PC List (Responsible columns just stay NaN). The
    asset ``Name`` column is dropped after the join because it duplicates
    ``HOSTNAME``.
    """
    df_cmap = _column_map(df)
    asset_cmap = _column_map(asset_df)

    df_key_col = df_cmap.get(df_key.lower())
    asset_key_col = asset_cmap.get(asset_key.lower())
    if df_key_col is None:
        logger.warning("dataframe missing HOSTNAME column; asset lookup skipped")
        return df
    if asset_key_col is None:
        logger.warning("Checkout PC List missing 'Name' column; asset lookup skipped")
        return df

    left = df.copy()
    right = asset_df.copy()

    left["__host__"] = _clean_key(left[df_key_col])
    right["__host__"] = _clean_key(right[asset_key_col])

    # Drop the asset Name column -- it duplicates HOSTNAME.
    right = right.drop(columns=[asset_key_col])

    merged = left.merge(right, on="__host__", how="left", suffixes=("", "_asset"))
    merged = merged.drop(columns=["__host__"])

    merged_cmap = _column_map(merged)
    resp_col = merged_cmap.get("responsible name")
    matched = int(merged[resp_col].notna().sum()) if resp_col else 0
    logger.info(
        "Joined with Checkout PC List (left): %d rows, %d hostnames matched an asset",
        len(merged),
        matched,
    )
    return merged


# --------------------------------------------------------------------------- #
# pivot: expand multiple laptops per VDI user into horizontal columns
# --------------------------------------------------------------------------- #
def pivot_laptops_horizontal(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse multiple laptop rows per VDI entry into a single wide row.

    Instead of concatenating laptop fields with a separator, this expands
    them horizontally: each laptop gets its own set of columns suffixed
    with ``_1``, ``_2``, ... The number of slots equals the maximum number
    of laptops any single VDI user matched. VDI users with fewer laptops
    get empty cells in the higher-numbered slots.

    Example: a VDI user with 2 laptops (NB-1 responsible Alice, NB-2
    responsible Bob) becomes a single row with::

        Id | Assigned Users | HOSTNAME_1 | Responsible Name_1 | HOSTNAME_2 | Responsible Name_2
        V1 | alice         | NB-1       | Alice Wang         | NB-2       | Bob Li
    """
    if "__row_id__" not in df.columns:
        logger.warning("__row_id__ column missing; skipping pivot")
        return df

    cmap = _column_map(df)

    # Map each expected column name to its actual (case-insensitive) name.
    fixed_actual = {
        name: cmap.get(name.lower())
        for name in config.VDI_FIXED_COLUMNS
        if cmap.get(name.lower()) is not None
    }
    laptop_actual = {
        name: cmap.get(name.lower())
        for name in config.LAPTOP_COLUMN_TEMPLATE
        if cmap.get(name.lower()) is not None
    }

    # Rank laptops within each VDI row (1-indexed). Rows with NaN HOSTNAME
    # (VDI users without any laptop match) get slot 1 but their laptop
    # columns are empty, so they occupy slot 1 with empty values.
    hostname_col = cmap.get("hostname")
    if hostname_col is not None:
        has_laptop = df[hostname_col].notna() & (df[hostname_col].astype(str).str.strip() != "")
    else:
        has_laptop = pd.Series([False] * len(df))

    df = df.copy()
    # Compute slot only among rows that actually have a laptop. Use a
    # separate dataframe so NaN __row_id__ doesn't corrupt the groupby.
    laptop_mask = has_laptop.fillna(False)
    if laptop_mask.any():
        slot_series = df.loc[laptop_mask].groupby("__row_id__").cumcount() + 1
        df["__slot__"] = slot_series
        df["__slot__"] = df["__slot__"].fillna(1).astype(int)
    else:
        df["__slot__"] = 1

    max_slots = int(df["__slot__"].max()) if not df.empty else 0
    if max_slots < 1:
        max_slots = 1
    logger.info(
        "Pivoting %d rows -> %d VDI rows, max %d laptop(s) per user",
        len(df),
        df["__row_id__"].nunique(),
        max_slots,
    )

    # Build the wide dataframe by iterating slots. For each slot, filter
    # rows where __slot__ == i, take the laptop columns, rename them with
    # _i suffix, and merge back onto the VDI-fixed columns (taken from
    # the first row of each __row_id__).
    base = df.groupby("__row_id__", as_index=False, dropna=False).agg(
        {actual: "first" for actual in fixed_actual.values()}
    )

    slot_frames = []
    for i in range(1, max_slots + 1):
        sub = df[df["__slot__"] == i][["__row_id__"] + list(laptop_actual.values())].copy()
        rename_map = {actual: f"{name}_{i}" for name, actual in laptop_actual.items()}
        sub = sub.rename(columns=rename_map)
        slot_frames.append(sub)

    result = base
    for sub in slot_frames:
        result = result.merge(sub, on="__row_id__", how="left")

    result = result.drop(columns=["__row_id__"])
    return result


# --------------------------------------------------------------------------- #
# orchestrator
# --------------------------------------------------------------------------- #
class FinalReportBuilder:
    """Build ``VDI_Laptop_Asset_final_<date>_<time>.xlsx`` from three inputs.

    Parameters
    ----------
    vdi_ad_path:
        Path to ``VDI_AD_final.xlsx`` (output of ``run.py``). If None, the
        builder looks for it in ``output_dir``.
    asset_root:
        Asset report root containing ``YYYY-MM-DD`` subfolders. The lastlogon
        file and Checkout PC List are both expected in the latest subfolder.
    lastlogon_path:
        Explicit path to the lastlogon .xlsx. If None, auto-discovered in the
        latest asset folder.
    checkout_path:
        Explicit path to the Checkout PC List .xlsx. If None, auto-discovered
        in the same folder as the lastlogon file.
    output_dir:
        Where to write the final report. Defaults to ``FINAL_REPORT_ROOT``
        (BMW "Status gengerate report" share). Falls back to ``cwd/output``
        if the share is not writable.
    keep_all_vdi:
        Deprecated -- always True now. LEFT join is the default (union
        semantics: all VDI rows are preserved). Kept for backwards
        compatibility with existing CLI flags.
    """

    def __init__(
        self,
        vdi_ad_path: Optional[Path] = None,
        asset_root: Optional[Path] = None,
        lastlogon_path: Optional[Path] = None,
        checkout_path: Optional[Path] = None,
        output_dir: Optional[Path] = None,
        vdi_ad_dir: Optional[Path] = None,
        keep_all_vdi: bool = True,
    ) -> None:
        self.vdi_ad_path = Path(vdi_ad_path) if vdi_ad_path else None
        self.asset_root = Path(asset_root) if asset_root else Path(config.ASSET_REPORT_ROOT)
        self.lastlogon_path = Path(lastlogon_path) if lastlogon_path else None
        self.checkout_path = Path(checkout_path) if checkout_path else None
        # Where to look for VDI_AD_final.xlsx if vdi_ad_path is not given.
        # Defaults to cwd/output (pipeline A's output dir).
        self.vdi_ad_dir = Path(vdi_ad_dir) if vdi_ad_dir else Path.cwd() / "output"
        # Default to FINAL_REPORT_ROOT (BMW share); fall back to cwd/output
        # if the caller didn't specify and the share path is unreachable.
        if output_dir is not None:
            self.output_dir = Path(output_dir)
        else:
            share = Path(config.FINAL_REPORT_ROOT)
            try:
                share.mkdir(parents=True, exist_ok=True)
                self.output_dir = share
            except Exception:
                logger.warning(
                    "FINAL_REPORT_ROOT not writable (%s); falling back to cwd/output",
                    share,
                )
                self.output_dir = Path.cwd() / "output"
        # keep_all_vdi is now always True (union semantics). The attribute is
        # kept for backwards compatibility but has no effect.
        self.keep_all_vdi = True

    # ------------------------------------------------------------------ #
    def run(self, today: Optional[dt.date] = None) -> Optional[pd.DataFrame]:
        logger.info("=== Final report builder started ===")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        today = today or dt.date.today()

        # 1. VDI+AD master
        vdi_ad_path = self._resolve_vdi_ad_path()
        if vdi_ad_path is None:
            return None
        vdi_df = load_vdi_ad_report(vdi_ad_path)
        if vdi_df.empty:
            logger.error("VDI+AD report is empty; aborting")
            return None

        # 2. lastlogon
        lastlogon_path, asset_folder = self._resolve_lastlogon_path(today)
        if lastlogon_path is None:
            logger.error("No lastlogon file found; aborting")
            return None
        lastlogon_df = load_lastlogon(lastlogon_path)
        if lastlogon_df.empty:
            logger.error("lastlogon file has no usable rows; aborting")
            return None

        # 3. LEFT join VDI + lastlogon (union: keep all VDI rows)
        merged = join_vdi_with_lastlogon(vdi_df, lastlogon_df, how="left")
        if merged.empty:
            logger.warning("VDI + lastlogon join produced 0 rows")

        # 4. Checkout PC List (optional -- if missing, just skip the enrichment)
        checkout_path = self._resolve_checkout_path(asset_folder)
        if checkout_path is not None:
            asset_df = load_checkout_responsible(checkout_path)
            if not asset_df.empty:
                merged = join_with_checkout(merged, asset_df)
        else:
            logger.warning("Checkout PC List not found; Responsible columns will be empty")

        # 5. Pivot: expand multiple laptops per VDI row into horizontal
        #    columns (HOSTNAME_1, Responsible Name_1, HOSTNAME_2, ...).
        merged = pivot_laptops_horizontal(merged)

        # 6. order columns + write with timestamped filename
        final_df = self._order_columns(merged)
        timestamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        final_name = f"{config.FINAL_REPORT_BASENAME}_{timestamp}.xlsx"
        final_out = self.output_dir / final_name
        final_df.to_excel(final_out, index=False)
        logger.info("Final report written: %s (%d rows, %d cols)",
                    final_out, len(final_df), len(final_df.columns))
        logger.info("=== Final report builder finished ===")
        return final_df

    # ------------------------------------------------------------------ #
    def _resolve_vdi_ad_path(self) -> Optional[Path]:
        if self.vdi_ad_path is not None:
            if not self.vdi_ad_path.is_file():
                logger.error("VDI+AD report not found: %s", self.vdi_ad_path)
                return None
            return self.vdi_ad_path
        # Look in vdi_ad_dir (pipeline A's output), NOT in output_dir
        # (which is the final report dir -- a different location).
        guessed = self.vdi_ad_dir / config.VDI_AD_FINAL_FILENAME
        if guessed.is_file():
            logger.info("Auto-found VDI+AD report: %s", guessed)
            return guessed
        logger.error(
            "VDI+AD report not found at %s; pass --vdi-ad-file or run run.py first",
            guessed,
        )
        return None

    def _resolve_lastlogon_path(
        self, today: dt.date
    ) -> tuple[Optional[Path], Optional[Path]]:
        """Return (lastlogon_path, asset_folder)."""
        if self.lastlogon_path is not None:
            if not self.lastlogon_path.is_file():
                logger.error("lastlogon file not found: %s", self.lastlogon_path)
                return None, None
            asset_folder = self.lastlogon_path.parent
            logger.info("Using explicit lastlogon file: %s", self.lastlogon_path)
            return self.lastlogon_path, asset_folder

        folder = find_latest_asset_folder(self.asset_root, today=today)
        if folder is None:
            logger.error("No asset folder found under %s", self.asset_root)
            return None, None
        lastlogon = find_lastlogon_file(folder)
        if lastlogon is None:
            logger.error(
                "No 'VDI-laptop-lastlogin-*.xlsx' found in %s; "
                "run run_asset.py first or pass --lastlogon-file",
                folder,
            )
            return None, None
        logger.info("Auto-found lastlogon file: %s", lastlogon)
        return lastlogon, folder

    def _resolve_checkout_path(
        self, asset_folder: Optional[Path]
    ) -> Optional[Path]:
        if self.checkout_path is not None:
            if not self.checkout_path.is_file():
                logger.warning("Checkout PC List not found: %s", self.checkout_path)
                return None
            return self.checkout_path
        if asset_folder is None:
            return None
        checkout = find_checkout_pc_list_file(asset_folder)
        if checkout is None:
            logger.warning("No Checkout PC List file in %s", asset_folder)
            return None
        logger.info("Auto-found Checkout PC List: %s", checkout)
        return checkout

    # ------------------------------------------------------------------ #
    def _order_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Reorder to the dynamic column layout.

        The layout is ``VDI_FIXED_COLUMNS`` followed by one slot per laptop
        (``LAPTOP_COLUMN_TEMPLATE`` with ``_N`` suffix). The number of slots
        is inferred from the dataframe's existing columns (e.g. if
        ``HOSTNAME_2`` exists, there are at least 2 slots).

        Missing columns are added as empty columns so every VDI row has the
        same shape (e.g. a VDI user with 1 laptop gets empty ``_2``,
        ``_3``, ... columns if another user has 3 laptops).
        """
        cmap = _column_map(df)

        # Detect how many laptop slots exist in the dataframe. Columns
        # are named "<base>_<N>" (e.g. "HOSTNAME_1", "Responsible Name_2").
        # We look for any column ending in "_<digit>" and take the max N.
        max_slots = 0
        for col in df.columns:
            s = str(col)
            # Strip the trailing _<digits>
            if "_" in s:
                last_underscore = s.rfind("_")
                suffix = s[last_underscore + 1:]
                if suffix.isdigit():
                    max_slots = max(max_slots, int(suffix))
        if max_slots < 1:
            # No laptop columns at all (e.g. Checkout PC List missing).
            # Still emit one slot so the report shape is stable.
            max_slots = 1

        target_cols = config.generate_final_columns(max_slots)
        ordered: List[pd.Series] = []
        for name in target_cols:
            actual = cmap.get(name.lower())
            if actual is not None:
                ordered.append(df[actual])
            else:
                ordered.append(pd.Series([pd.NA] * len(df), name=name))
        result = pd.concat(ordered, axis=1)
        result.columns = target_cols
        logger.info("Final report column layout: %d slots (%d total columns)",
                    max_slots, len(target_cols))
        return result
