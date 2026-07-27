"""Selenium automation for the wass.bmwgroup.net Inventory report wizard.

Flow (per the documented steps):

  1. Open https://wass.bmwgroup.net, switch to the **Inventory** tab.
  2. Click **create report** -> **New Report**.
  3. Fill the report *name*.
  4. Open the **Import** picker (right-side first button) -> **Import computer list**.
  5. Paste the computer name list (<=4999 per report) into the text area,
     click the **Import computer list** link, wait for the right pane to list
     the computers, then **OK**.
  6. **Next** -> open **Result elements** -> expand **login** -> tick
     Last Logon / Last User / Last User(Account) / Last User(Email) -> **OK**.
  7. **Next** -> **OK** to start the report.
  8. Poll the small top-right refresh button until the entry's status changes
     from "Report is running" to "Report successfully created".
  9. Right-click the entry -> **Result** -> click **Excel Download** ->
     click the "Click here to download the file." hyperlink.
 10. The generated .xlsx lands in the configured download directory.

IMPORTANT
---------
The wass UI is an internal BMW tool whose exact HTML is not known up-front.
Locators are centralised in :data:`DEFAULT_LOCATORS` and are text-based
(XPath ``contains(text(), ...)``). When a step fails on the real site, adjust
the corresponding locator string -- no other code should need to change.
This module can only be exercised on a Windows machine that can reach
``wass.bmwgroup.net``; it cannot run in the dev sandbox.
"""
from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from . import config

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Locators. All are (By, value) tuples. Text-based XPath is used so the
# automation follows the visible labels described by the user. Adjust these
# strings if the real wass UI uses different wording.
# ---------------------------------------------------------------------------
DEFAULT_LOCATORS: Dict[str, Any] = {
    # Top-level navigation. All text-based XPaths use translate() so they are
    # case-insensitive -- wass uses "Create report" (mixed case), not the
    # lowercase "create report" the user described verbally.
    "inventory_tab": (
        By.XPATH,
        "//*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='inventory']",
    ),
    "create_report": (
        By.XPATH,
        "//*[contains(normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')),'create report')]",
    ),
    "new_report": (
        By.XPATH,
        "//*[contains(normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')),'new report')]",
    ),
    # Report name input -- any visible text input near the top of the wizard.
    "report_name_input": (By.CSS_SELECTOR, "input[type='text']"),
    # Import step
    "import_button": (
        By.XPATH,
        "//*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='import']",
    ),
    "import_computer_list_link": (
        By.XPATH,
        "//*[contains(normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')),'import computer list')]",
    ),
    "import_textarea": (By.CSS_SELECTOR, "textarea"),
    # Wizard navigation
    "ok_button": (
        By.XPATH,
        "//*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='ok']",
    ),
    "next_button": (
        By.XPATH,
        "//*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='next']",
    ),
    "result_elements_button": (
        By.XPATH,
        "//*[contains(normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')),'result element')]",
    ),
    "login_group": (
        By.XPATH,
        "//*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='login']",
    ),
    # Status / refresh
    "refresh_button": (
        By.XPATH,
        "//button[contains(@title,'refresh') or contains(@aria-label,'refresh')]"
        " | //*[contains(@class,'refresh')]",
    ),
    # Result dialog
    "result_menu_item": (
        By.XPATH,
        "//*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='result']",
    ),
    "excel_download_button": (
        By.XPATH,
        "//*[contains(normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')),'excel download')]",
    ),
    "download_link": (
        By.XPATH,
        "//a[contains(.,'Click here to download')]"
        " | //*[contains(.,'successfully created')]",
    ),
}

# Status text the wizard reports while a report is being generated / done.
STATUS_RUNNING = "Report is running"
STATUS_DONE = "Report successfully created"


class WassReportAutomator:
    """Drive the wass Inventory report wizard via Selenium.

    Parameters
    ----------
    download_dir:
        Directory where the generated Excel lands. Defaults to the asset
        report's YYYY-MM-DD folder (set by the pipeline).
    browser:
        ``"edge"`` (default, ships with Windows) or ``"chrome"``.
    headless:
        Run the browser headless. Off by default -- the wizard often needs a
        visible session and corporate SSO may not work headless.
    login_timeout:
        Seconds to wait for the user to complete manual SSO login on first
        launch before the automation takes over.
    locators:
        Override individual entries of :data:`DEFAULT_LOCATORS`.
    """

    def __init__(
        self,
        download_dir: Optional[Path] = None,
        browser: str = config.WASS_BROWSER,
        headless: bool = False,
        login_timeout: int = 120,
        locators: Optional[Dict[str, Any]] = None,
    ) -> None:
        self.download_dir = Path(download_dir) if download_dir else Path.cwd()
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.browser = (browser or config.WASS_BROWSER).lower()
        self.headless = headless
        self.login_timeout = login_timeout
        self.locators = {**DEFAULT_LOCATORS, **(locators or {})}
        self.driver: Optional[webdriver.Remote] = None

    # ------------------------------------------------------------------ #
    # Browser lifecycle
    # ------------------------------------------------------------------ #
    def start(self) -> None:
        """Launch the configured browser and navigate to wass."""
        logger.info("Launching %s browser (headless=%s)", self.browser, self.headless)
        if self.browser == "edge":
            opts = EdgeOptions()
            opts.add_experimental_option("prefs", self._download_prefs())
            if self.headless:
                opts.add_argument("--headless=new")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            try:
                from webdriver_manager.microsoft import EdgeChromiumDriverManager

                service = EdgeService(EdgeChromiumDriverManager().install())
            except Exception:  # fall back to system msedgedriver
                service = EdgeService()
            self.driver = webdriver.Edge(service=service, options=opts)
        elif self.browser == "chrome":
            opts = ChromeOptions()
            opts.add_experimental_option("prefs", self._download_prefs())
            if self.headless:
                opts.add_argument("--headless=new")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            try:
                from webdriver_manager.chrome import ChromeDriverManager

                service = ChromeService(ChromeDriverManager().install())
            except Exception:
                service = ChromeService()
            self.driver = webdriver.Chrome(service=service, options=opts)
        else:
            raise ValueError(f"Unsupported browser: {self.browser}")

        self.driver.maximize_window()
        self.driver.get(config.WASS_URL)
        logger.info("Navigated to %s", config.WASS_URL)

    def _download_prefs(self) -> Dict[str, Any]:
        return {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "plugins.always_open_pdf_externally": True,
        }

    def close(self) -> None:
        if self.driver:
            try:
                self.driver.quit()
            finally:
                self.driver = None

    def __enter__(self) -> "WassReportAutomator":
        self.start()
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    # ------------------------------------------------------------------ #
    # Low-level helpers
    # ------------------------------------------------------------------ #
    @property
    def _d(self) -> webdriver.Remote:
        if self.driver is None:
            raise RuntimeError("Browser not started; call start() first")
        return self.driver

    def _wait(self, timeout: float = 20) -> WebDriverWait:
        return WebDriverWait(self._d, timeout)

    def _loc(self, name: str):
        return self.locators[name]

    def _click(self, name: str, timeout: float = 20) -> None:
        by, val = self._loc(name)
        el = self._wait(timeout).until(
            EC.element_to_be_clickable((by, val)),
            message=f"locator '{name}' not clickable",
        )
        self._scroll_into_view(el)
        try:
            el.click()
        except Exception as e:
            # "element click intercepted" is common on overlay/hover nav menus
            # (e.g. wass MainNaviBottom). Fall back to a JS click, which bypasses
            # the overlay-hit-test that the normal WebDriver click performs.
            logger.debug("normal click failed (%s); retrying via JS", type(e).__name__)
            self._d.execute_script("arguments[0].click();", el)
        logger.info("clicked: %s", name)

    def _scroll_into_view(self, el) -> None:
        try:
            self._d.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", el
            )
        except Exception:
            pass

    # ------------------------------------------------------------------ #
    # High-level wizard steps
    # ------------------------------------------------------------------ #
    def ensure_logged_in(self) -> None:
        """Wait until the Inventory tab is reachable.

        If corporate SSO requires manual interaction on first launch, the
        automation waits up to ``login_timeout`` seconds for the user.
        """
        logger.info("Waiting for Inventory tab (login if needed)...")
        try:
            self._wait(self.login_timeout).until(
                EC.presence_of_element_located(self._loc("inventory_tab")),
                message="Inventory tab not visible",
            )
            logger.info("Logged in / Inventory tab visible")
        except Exception:
            logger.warning(
                "Inventory tab not detected within %ds. "
                "If a login page is shown, complete SSO in the browser window, "
                "then re-run or wait.",
                self.login_timeout,
            )
            raise

    def open_inventory(self) -> None:
        self._click("inventory_tab")
        self._short_pause()

    def start_new_report(self, report_name: str) -> None:
        self._click("create_report")
        self._short_pause()
        self._click("new_report")
        self._short_pause()
        # Fill the report name (top-most text input).
        by, val = self._loc("report_name_input")
        name_el = self._wait().until(
            EC.presence_of_element_located((by, val)),
            message="report name input not found",
        )
        name_el.clear()
        name_el.send_keys(report_name)
        logger.info("report name set: %s", report_name)

    def import_computer_list(self, names: List[str]) -> None:
        """Open the Import picker, paste *names*, trigger the import, confirm."""
        self._click("import_button")
        self._short_pause()
        self._click("import_computer_list_link")
        self._short_pause()

        by, val = self._loc("import_textarea")
        textarea = self._wait().until(
            EC.presence_of_element_located((by, val)),
            message="import textarea not found",
        )
        textarea.clear()
        textarea.send_keys("\n".join(names))
        logger.info("pasted %d computer names into import textarea", len(names))

        # Click the link again to actually import what was pasted.
        self._click("import_computer_list_link")

        # Wait for the right pane to reflect the imported computers. There is
        # no stable selector, so we just give the UI a moment to render.
        time.sleep(3)
        logger.info("imported computer list submitted")

        self._click("ok_button")

    def select_result_elements(self, elements: List[str]) -> None:
        self._click("result_elements_button")
        self._short_pause()
        # Expand the "login" group (click its row to toggle).
        self._click("login_group")
        self._short_pause()
        for label in elements:
            # Tick the checkbox whose row text contains the label.
            xpath = (
                f"//tr[contains(.,'{label}')]"
                f"//input[@type='checkbox']"
                f" | //*[normalize-space(text())='{label}']"
                f"/preceding::input[@type='checkbox'][1]"
                f" | //*[normalize-space(text())='{label}']"
                f"/ancestor::tr[1]//input[@type='checkbox']"
            )
            try:
                cb = self._wait(10).until(
                    EC.presence_of_element_located((By.XPATH, xpath)),
                    message=f"checkbox '{label}' not found",
                )
                if not cb.is_selected():
                    cb.click()
                    logger.info("ticked: %s", label)
            except Exception:
                logger.warning("could not tick '%s' -- adjust locators if needed", label)
        self._click("ok_button")

    def finish_wizard(self) -> None:
        self._click("next_button")
        self._short_pause()
        self._click("ok_button")

    def wait_for_completion(
        self,
        report_name: str,
        timeout: int = config.WASS_REPORT_TIMEOUT_SECONDS,
        poll_interval: int = config.WASS_POLL_INTERVAL_SECONDS,
    ) -> None:
        """Click the top-right refresh button until the report is done."""
        logger.info(
            "waiting for report '%s' to finish (timeout=%ds)", report_name, timeout
        )
        deadline = time.time() + timeout
        last_status = None
        while time.time() < deadline:
            self._refresh()
            status = self._read_status(report_name)
            if status != last_status:
                logger.info("status: %s", status or "(unknown)")
                last_status = status
            if status and STATUS_DONE.lower() in status.lower():
                logger.info("report '%s' finished", report_name)
                return
            time.sleep(poll_interval)
        raise TimeoutError(
            f"Report '{report_name}' did not finish within {timeout}s "
            f"(last status: {last_status})"
        )

    def _refresh(self) -> None:
        """Click the small top-right refresh button; fall back to F5."""
        try:
            self._click("refresh_button", timeout=5)
        except Exception:
            try:
                self._d.refresh()
            except Exception:
                pass

    def _read_status(self, report_name: str) -> str:
        """Best-effort read of the status text next to *report_name*."""
        try:
            # Look for any element whose text mentions the report name and a
            # status keyword.
            xpath = (
                f"//*[contains(.,'{report_name}')]"
                f"//*[contains(.,'{STATUS_RUNNING}')"
                f" or contains(.,'{STATUS_DONE}')]"
                f" | //*[contains(.,'{report_name}')]"
                f"[contains(.,'{STATUS_RUNNING}')"
                f" or contains(.,'{STATUS_DONE}')]"
            )
            els = self._d.find_elements(By.XPATH, xpath)
            for el in els:
                txt = (el.text or "").strip()
                if STATUS_DONE.lower() in txt.lower():
                    return STATUS_DONE
                if STATUS_RUNNING.lower() in txt.lower():
                    return STATUS_RUNNING
            return ""
        except Exception:
            return ""

    def download_result(self, report_name: str) -> Optional[Path]:
        """Right-click the entry -> Result -> Excel Download -> download link.

        Returns the path of the newly downloaded .xlsx file.
        """
        before = set(self.download_dir.glob("*.xlsx"))
        logger.info("opening Result for '%s'", report_name)

        # Locate the report row and right-click it to open the context menu.
        row_xpath = f"//*[contains(.,'{report_name}')]"
        try:
            row = self._wait(15).until(
                EC.presence_of_element_located((By.XPATH, row_xpath)),
                message=f"report row '{report_name}' not found",
            )
            ActionChains(self._d).context_click(row).perform()
            self._short_pause()
            self._click("result_menu_item", timeout=10)
        except Exception:
            # Fallback: some UIs expose Result as a direct button/link.
            logger.warning("context-menu path failed; trying direct Result click")
            self._click("result_menu_item", timeout=10)

        self._short_pause()
        self._click("excel_download_button", timeout=15)
        self._short_pause()

        try:
            self._click("download_link", timeout=30)
        except Exception:
            logger.warning("download link not clicked via locator; "
                           "the file may still download via Excel Download")

        downloaded = self._wait_for_new_xlsx(before)
        if downloaded:
            logger.info("downloaded: %s", downloaded)
        else:
            logger.error("no new .xlsx detected in %s", self.download_dir)
        return downloaded

    def _wait_for_new_xlsx(
        self, before, timeout: int = 180
    ) -> Optional[Path]:
        """Wait until a new .xlsx appears (and no .crdownload is in progress)."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            # If a download is in progress, keep waiting.
            if any(self.download_dir.glob("*.crdownload")) or any(
                self.download_dir.glob("*.tmp")
            ):
                time.sleep(1)
                continue
            after = set(self.download_dir.glob("*.xlsx"))
            new_files = sorted(after - before, key=lambda p: p.stat().st_mtime)
            if new_files:
                return new_files[-1]
            time.sleep(2)
        return None

    def _short_pause(self, seconds: float = 1.0) -> None:
        time.sleep(seconds)

    # ------------------------------------------------------------------ #
    # End-to-end driver for a single chunk
    # ------------------------------------------------------------------ #
    def create_report(
        self,
        report_name: str,
        computer_names: List[str],
        result_elements: Optional[List[str]] = None,
    ) -> Optional[Path]:
        """Run the full wizard once for *computer_names* and download the result."""
        result_elements = result_elements or config.WASS_RESULT_ELEMENTS
        logger.info(
            "=== creating wass report '%s' (%d computers) ===",
            report_name,
            len(computer_names),
        )
        self.open_inventory()
        self.start_new_report(report_name)
        self.import_computer_list(computer_names)
        self.select_result_elements(result_elements)
        self.finish_wizard()
        self.wait_for_completion(report_name)
        return self.download_result(report_name)
