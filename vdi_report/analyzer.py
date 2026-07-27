"""Orchestrates the full VDI -> AD report pipeline."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

import pandas as pd

from . import config
from .ad_lookup import explode_assigned_users, join_vdi_with_ad, load_ad_report
from .paths import find_latest_ad_report, find_latest_vdi_folder
from .vdi_merger import find_z3_z4_files, merge_vdi_reports

logger = logging.getLogger(__name__)


class VDIReportAnalyzer:
    """Run the two-step pipeline and write intermediate + final reports.

    Step 1 merges the latest Z3/Z4 VDI VMware reports into ``VDI_final.xlsx``.
    Step 2 joins that with the latest AD Nameanddepartment report, producing
    ``VDI_AD_final.xlsx`` keyed by VDI Assigned Users -> AD Account.
    """

    def __init__(
        self,
        vdi_root: Optional[Path] = None,
        ad_root: Optional[Path] = None,
        output_dir: Optional[Path] = None,
        explode_users: bool = True,
        keep_vdi_source_col: bool = False,
        write_vdi_final: bool = True,
    ) -> None:
        self.vdi_root = Path(vdi_root) if vdi_root else Path(config.VDI_REPORT_ROOT)
        self.ad_root = Path(ad_root) if ad_root else Path(config.AD_REPORT_ROOT)
        self.output_dir = Path(output_dir) if output_dir else Path.cwd() / "output"
        self.explode_users = explode_users
        self.keep_vdi_source_col = keep_vdi_source_col
        # When False, the intermediate VDI_final.xlsx (Z3+Z4 merge only) is
        # skipped -- run_all.py sets this since the consolidated report is the
        # only output users care about.
        self.write_vdi_final = write_vdi_final

    # ------------------------------------------------------------------ #
    def run(self) -> Optional[pd.DataFrame]:
        logger.info("=== VDI Report Analyzer started ===")
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # --- Step 1: merge Z3 + Z4 ---
        vdi_df = self._step1_merge_vdi()
        if vdi_df is None or vdi_df.empty:
            logger.error("Step 1 produced no data; aborting pipeline")
            return None
        if self.write_vdi_final:
            vdi_out = self.output_dir / "VDI_final.xlsx"
            vdi_df.to_excel(vdi_out, index=False)
            logger.info("Step 1 written: %s (%d rows)", vdi_out, len(vdi_df))
        else:
            logger.info("Step 1 skipped VDI_final.xlsx write (write_vdi_final=False)")

        # --- Step 2: join with AD ---
        final_df = self._step2_join_ad(vdi_df)
        if final_df is None:
            logger.warning("Step 2 skipped; only VDI data is available")
            return vdi_df
        final_out = self.output_dir / "VDI_AD_final.xlsx"
        final_df.to_excel(final_out, index=False)
        logger.info("Step 2 written: %s (%d rows)", final_out, len(final_df))
        logger.info("=== VDI Report Analyzer finished ===")
        return final_df

    # ------------------------------------------------------------------ #
    def _step1_merge_vdi(self) -> Optional[pd.DataFrame]:
        logger.info("[Step 1] Merging Z3 + Z4 VDI VMware reports")
        folder = find_latest_vdi_folder(self.vdi_root)
        if folder is None:
            return None
        z3, z4 = find_z3_z4_files(folder)
        if z3 is None and z4 is None:
            logger.error("Neither Z3 nor Z4 file found in %s", folder)
            return None
        df = merge_vdi_reports(z3, z4)
        if not self.keep_vdi_source_col and "__source__" in df.columns:
            df = df.drop(columns=["__source__"])
        return df

    def _step2_join_ad(self, vdi_df: pd.DataFrame) -> Optional[pd.DataFrame]:
        logger.info("[Step 2] Joining with AD Nameanddepartment report")
        ad_path = find_latest_ad_report(self.ad_root)
        if ad_path is None:
            logger.warning("No AD report found; returning VDI-only result")
            return None
        ad_df = load_ad_report(ad_path)
        if self.explode_users:
            vdi_df = explode_assigned_users(vdi_df)
        return join_vdi_with_ad(vdi_df, ad_df)
