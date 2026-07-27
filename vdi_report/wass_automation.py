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
import re
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
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

from . import config

logger = logging.getLogger(__name__)


def _decode_js_unicode(s: str) -> str:
    """Decode ``\\uXXXX`` escape sequences in a JS source string.

    ``<script>.textContent`` returns the *raw* source text -- escape sequences
    like ``\\u0027`` (a single quote) are kept as the 6 literal characters
    ``\\ u 0 0 2 7`` and are only interpreted when the JS is executed. wass/
    Matrix42 registers button onclick handlers via
    ``EmControls.Events.SetOnClick("BtnId", {"OnClick": "ClosePopup(\\u0027Level3\\u0027)"})``
    so reading the script text yields ``ClosePopup(\\u0027Level3\\u0027)`` --
    which never substring-matches the literal ``ClosePopup('Level3')`` we look
    for. Decoding the escapes first makes the match work.
    """
    if not s:
        return s
    return re.sub(
        r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), s
    )


def _ci(label: str) -> str:
    """Build a robust, case-insensitive XPath that matches *label* as text.

    wass is a Matrix42 Enterprise Manager customization. Its buttons are
    rendered as ``<div id="Button<guid>">Label</div>`` widgets, and it also
    uses Kendo UI (``.k-button``). To be resilient we match a label against
    several element shapes:

      * any element whose own text equals *label* (case-insensitive)
      * any element whose own text contains *label* (case-insensitive)
      * a Matrix42 button div (``div[starts-with(@id,'Button')]``) whose text
        contains *label*
      * a Kendo button (``.k-button``) whose text contains *label*
      * an element with ``title``/``aria-label`` containing *label*
    """
    low = label.lower()
    upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    lower = "abcdefghijklmnopqrstuvwxyz"
    tl = f"translate(text(),'{upper}','{lower}')"
    tns = f"translate(normalize-space(.),'{upper}','{lower}')"
    return (
        f"//*[normalize-space(translate(text(),'{upper}','{lower}'))='{low}']"
        f" | //*[contains(normalize-space(translate(text(),'{upper}','{lower}')),'{low}')]"
        f" | //div[starts-with(@id,'Button')][contains({tns},'{low}')]"
        f" | //*[(contains(@class,'k-button') or contains(@class,'btn'))][contains({tns},'{low}')]"
        f" | //*[contains(@title,'{label}') or contains(@aria-label,'{label}')]"
    )


# ---------------------------------------------------------------------------
# Locators. All are (By, value) tuples. Text-based XPath is used so the
# automation follows the visible labels described by the user. Adjust these
# strings if the real wass UI uses different wording.
# ---------------------------------------------------------------------------
DEFAULT_LOCATORS: Dict[str, Any] = {
    # Top-level navigation: Matrix42 nav items have stable ids.
    # MainNaviBottom_2 = Inventory, MainNaviBottom_1 = Home, etc.
    "inventory_tab": (By.CSS_SELECTOR, "#MainNaviBottom_2, #MainNaviBottom_2 span"),
    "create_report": (By.XPATH, _ci("create report")),
    "new_report": (By.XPATH, _ci("new report")),
    # Report name input -- Matrix42 uses <input id="Name">.
    "report_name_input": (By.CSS_SELECTOR, "#Name, input#Name"),
    # Import step. #ImportList is the IconButton that opens the import picker.
    "import_button": (By.CSS_SELECTOR, "#ImportList"),
    "import_computer_list_link": (
        By.CSS_SELECTOR,
        "a.EmTextLink[href*='Host_Import']",
    ),
    "import_textarea": (By.CSS_SELECTOR, "#ListResult, textarea#ListResult"),
    # Result elements step. #Output is the IconButton (tooltip "Select output
    # elements") that opens the output-element picker popup (Level3).
    "result_elements_button": (By.CSS_SELECTOR, "#Output"),
    # NOTE: wizard Next/OK buttons are NOT matched by static locators -- wass
    # has multiple buttons with the same text ("Next", "OK") across wizard
    # steps. They are clicked via _click_button_by_onclick(), which finds all
    # Matrix42 Button widgets with the given label and filters by the onclick
    # handler registered in the sibling <script>. Mapping (used by the
    # high-level steps):
    #   - Step1 Next  -> label="Next",  onclick~"Inv_Rep_Step2.aspx"
    #   - Step2 Next  -> label="Next",  onclick~"Inv_Rep_Step3.aspx"
    #   - Picker OK   -> label="OK",    onclick~"ClosePopup('Level3')"
    #   - Generate OK -> label="OK",    onclick~"Save_Report.aspx"
    # Generic fallbacks kept for any other step:
    "ok_button": (By.XPATH, _ci("ok")),
    "next_button": (By.XPATH, _ci("next")),
    # Status / refresh. The refresh icon is a small <img alt="Update"
    # src="img/Table_Refresh.png"> in the top-right of the report list.
    "refresh_button": (
        By.CSS_SELECTOR,
        "img[alt='Update'], img[src*='Table_Refresh.png']",
    ),
    # Result dialog (right-click report row -> Result -> Excel Download ->
    # "Your file was successfully created. Click here to download the file.").
    # "Result" is a jQuery contextMenu item: the menu is a
    # <ul class="context-menu-list"> whose items are
    # <li class="context-menu-item"><span>Result</span></li>. Scope the XPath
    # to the menu so we don't accidentally match a "Result" column header or
    # other stray "Result" text elsewhere on the page (those appear earlier in
    # DOM order and would otherwise be clicked first).
    "result_menu_item": (
        By.XPATH,
        "//ul[contains(@class,'context-menu')]"
        "//li[contains(@class,'context-menu-item')]"
        "[normalize-space()='Result']"
        " | //li[contains(@class,'context-menu-item')][normalize-space()='Result']"
        " | //*[contains(@class,'context-menu-item')]//span[normalize-space()='Result']",
    ),
    # "Excel Download" is a <td> whose inner <div class='EmText'> text is
    # exactly "Excel Download".
    "excel_download_button": (
        By.XPATH,
        "//td[div[contains(@class,'EmText') and normalize-space()='Excel Download']]"
        " | //div[normalize-space()='Excel Download']"
        " | //*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='excel download']",
    ),
    # The download link is an <a class='EmTextLink'> pointing at a .xlsx file,
    # with the text "Your file was successfully created. Click here to
    # download the file."
    "download_link": (
        By.XPATH,
        "//a[contains(@href,'.xlsx')]"
        " | //a[contains(.,'Click here to download')]"
        " | //*[contains(.,'successfully created')]//a"
        " | //a[contains(.,'successfully created')]",
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
        el = self._find_clickable(name, by, val, timeout)
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

    def _find_clickable(self, name: str, by, val, timeout: float):
        """Find *name* in the top document; if not found, search every iframe.

        Matrix42 wizard content is sometimes rendered inside an <iframe>, in
        which case the top-document XPath won't see it. We try the top document
        first, then walk all iframes (one level deep) and retry there,
        restoring the top frame afterwards.
        """
        try:
            return self._wait(timeout).until(
                EC.element_to_be_clickable((by, val)),
                message=f"locator '{name}' not clickable",
            )
        except Exception:
            pass

        # Try each iframe.
        self._d.switch_to.default_content()
        iframes = self._d.find_elements(By.TAG_NAME, "iframe")
        logger.debug(
            "top document miss for '%s'; searching %d iframe(s)", name, len(iframes)
        )
        for idx, f in enumerate(iframes):
            try:
                self._d.switch_to.default_content()
                self._d.switch_to.frame(f)
                el = self._wait(3).until(
                    EC.element_to_be_clickable((by, val)),
                    message=f"locator '{name}' not in iframe {idx}",
                )
                logger.debug("'%s' found in iframe %d", name, idx)
                return el
            except Exception:
                continue
            finally:
                try:
                    self._d.switch_to.default_content()
                except Exception:
                    pass
        raise TimeoutException(f"locator '{name}' not clickable (top doc + iframes)")

    def _scroll_into_view(self, el) -> None:
        try:
            self._d.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", el
            )
        except Exception:
            pass

    def _click_button_by_onclick(
        self, name: str, label: str, onclick_contains: str, timeout: float = 20
    ) -> None:
        """Click a Matrix42 Button widget (``<div id="Button...">``) matched by
        both its visible label and a substring of its onclick handler.

        wass reuses the labels "Next" and "OK" across multiple wizard steps.
        The buttons share text but have different onclick handlers, so the
        only reliable way to pick the right one is to find all Button widgets
        with the given label and filter by the onclick content. The onclick
        is registered via ``EmControls.Events.SetOnClick`` in a ``<script>``
        tag immediately following the button ``<div>``.

        Parameters
        ----------
        name:
            Logical name used in log messages (e.g. "next_step1_button").
        label:
            Visible button text, e.g. "Next", "OK".
        onclick_contains:
            Substring that must appear in the button's registration script,
            e.g. "Inv_Rep_Step3.aspx" or "Save_Report.aspx".
        """
        deadline = time.time() + timeout
        last_err = None
        last_candidate_info = []
        while time.time() < deadline:
            # All Matrix42 Button widgets: <div id="Button..."> containing a
            # <span class="ui-button-text"> with the label.
            xpath = (
                f"//div[starts-with(@id,'Button')]"
                f"[.//span[@class='ui-button-text' and normalize-space()='{label}']]"
            )
            candidates = self._d.find_elements(By.XPATH, xpath)
            last_candidate_info = []
            for btn in candidates:
                try:
                    btn_id = btn.get_attribute("id") or ""
                    if not btn_id:
                        continue
                    # The onclick handler is registered via
                    # EmControls.Events.SetOnClick("btn_id", {...OnClick...})
                    # in a <script> tag. The script is usually a sibling of the
                    # button div, but the exact position varies (sometimes it's
                    # the parent's sibling, or further down). Search the whole
                    # document for any <script> mentioning this btn_id and read
                    # its text -- that's the most robust source.
                    script = self._d.execute_script(
                        "var id = arguments[0];"
                        "var scripts = document.getElementsByTagName('script');"
                        "for (var i = 0; i < scripts.length; i++) {"
                        "  var t = scripts[i].textContent || '';"
                        "  if (t.indexOf(id) !== -1) return t;"
                        "}"
                        "return '';",
                        btn_id,
                    ) or ""
                    # Decode \uXXXX escapes: Matrix42 stores the onclick as a
                    # JSON string, so quotes inside it appear as \u0027 in the
                    # raw script text and would never match a literal target
                    # like ClosePopup('Level3').
                    script_norm = _decode_js_unicode(script)
                    visible = btn.is_displayed()
                    last_candidate_info.append(
                        f"id={btn_id} visible={visible} script_has_target="
                        f"{onclick_contains in script_norm}"
                    )
                    if onclick_contains in script_norm and visible:
                        self._scroll_into_view(btn)
                        try:
                            btn.click()
                        except Exception:
                            self._d.execute_script("arguments[0].click();", btn)
                        logger.info(
                            "clicked: %s (label='%s', onclick~'%s')",
                            name, label, onclick_contains,
                        )
                        return
                except Exception as e:
                    last_err = e
                    continue
            time.sleep(0.5)
        # Diagnostic: log what we found so the user/AI can adjust the locator.
        logger.error(
            "button '%s' not found. Candidates seen (label='%s', target='%s'): %s",
            name, label, onclick_contains,
            last_candidate_info or "(none)",
        )
        raise TimeoutException(
            f"button '{name}' (label='{label}', onclick~'{onclick_contains}') "
            f"not found; candidates: {last_candidate_info or '(none)'}; "
            f"last error: {last_err}"
        )

    def _click_any_visible_button(
        self, name: str, label: str, timeout: float = 10
    ) -> None:
        """Click the first *visible* Matrix42 Button widget whose label matches.

        Fallback used when :meth:`_click_button_by_onclick` cannot match the
        onclick handler. At any given wizard step only the relevant button is
        on screen, so the first visible ``<div id="Button...">`` with the
        matching ``<span class="ui-button-text">`` text is the right one.
        """
        deadline = time.time() + timeout
        xpath = (
            f"//div[starts-with(@id,'Button')]"
            f"[.//span[@class='ui-button-text' and normalize-space()='{label}']]"
        )
        while time.time() < deadline:
            for btn in self._d.find_elements(By.XPATH, xpath):
                try:
                    if not btn.is_displayed():
                        continue
                    self._scroll_into_view(btn)
                    try:
                        btn.click()
                    except Exception:
                        self._d.execute_script("arguments[0].click();", btn)
                    logger.info(
                        "clicked: %s (any visible '%s' button)", name, label
                    )
                    return
                except Exception:
                    continue
            time.sleep(0.5)
        raise TimeoutException(
            f"no visible '{label}' button found for '{name}'"
        )

    def _click_context_menu_item(
        self, name: str, label: str, timeout: float = 10
    ) -> None:
        """Click a jQuery contextMenu item by its visible label.

        wass renders right-click menus with the jQuery contextMenu plugin as
        ``<ul class="context-menu-list"><li class="context-menu-item"><span>
        Label</span></li></ul>``. Two snags with the normal :meth:`_click`:

          1. ``EC.element_to_be_clickable`` often times out on these ``<li>``
             items (the plugin uses CSS that makes Selenium's visibility /
             enabled check fail even though the item is plainly on screen).
          2. A bare ``text()='label'`` XPath would also match a same-named
             column header or other stray text elsewhere on the page, which
             appears earlier in DOM order and gets clicked instead of the menu
             item.

        We therefore scope the search to ``.context-menu-item`` elements,
        wait for *presence* (not clickability), then JS-click the first
        visible one. JS ``.click()`` reliably fires the plugin's delegated
        handler on the ``<li>``.
        """
        xpath = (
            "//li[contains(@class,'context-menu-item')]"
            f"[normalize-space()='{label}']"
            " | //*[contains(@class,'context-menu-item')]//span"
            f"[normalize-space()='{label}']"
        )
        deadline = time.time() + timeout
        last_seen: List[str] = []
        while time.time() < deadline:
            els = self._d.find_elements(By.XPATH, xpath)
            last_seen = []
            for el in els:
                try:
                    vis = el.is_displayed()
                    last_seen.append(
                        f"<{el.tag_name}> visible={vis} text={el.text!r}"
                    )
                    if not vis:
                        continue
                    self._scroll_into_view(el)
                    try:
                        el.click()
                    except Exception:
                        self._d.execute_script("arguments[0].click();", el)
                    logger.info("clicked: %s (context menu '%s')", name, label)
                    return
                except Exception:
                    continue
            time.sleep(0.4)
        raise TimeoutException(
            f"context menu item '{label}' for '{name}' not found; "
            f"candidates: {last_seen or '(none)'}"
        )

    def _find_report_row(self, report_name: str, timeout: float = 20):
        """Find the *newest* report-list row whose name matches *report_name*.

        wass renders the report list as rows
        ``<tr cr="true" cm="Menu1" cp="#id#|<id>||">`` with cells:
        [icon, id, name, scope, owner, date]. The list accumulates **every**
        report the user ever created, so matching by name alone can return
        multiple rows -- re-runs of the same day, or chunk suffixes
        (``-02``/``-03``) that substring-match the base name. Report IDs are
        monotonic, so the newest report has the largest ID; we pick that one.

        The name match is **exact** (cell text == report_name), not a
        substring, so ``VDI-laptop-lastlogin-2026-07-27`` will NOT match
        ``VDI-laptop-lastlogin-2026-07-27-02``.
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            rows = self._d.find_elements(
                By.XPATH, "//tr[@cr='true' and @cm]"
            )
            candidates = []  # (id_int, row_el, info_str)
            for r in rows:
                try:
                    tds = r.find_elements(By.TAG_NAME, "td")
                    if len(tds) < 3:
                        continue
                    # Find the cell whose text exactly equals report_name
                    # (don't hardcode the index -- column order could shift).
                    matched = False
                    for td in tds:
                        if (td.text or "").strip() == report_name:
                            matched = True
                            break
                    if not matched:
                        continue
                    # Report ID from the cp attribute: "#id#|32806||" -> 32806.
                    cp = r.get_attribute("cp") or ""
                    rid = -1
                    parts = cp.split("|")
                    if len(parts) >= 2:
                        try:
                            rid = int(parts[1].strip())
                        except ValueError:
                            pass
                    # Fallback: 2nd <td> text is usually the ID.
                    if rid < 0:
                        try:
                            rid = int((tds[1].text or "").strip())
                        except ValueError:
                            rid = -1
                    date_text = (tds[-1].text or "").strip() if tds else ""
                    candidates.append(
                        (rid, r, f"id={rid} date={date_text!r}")
                    )
                except Exception:
                    continue
            if candidates:
                # Highest ID = newest report.
                candidates.sort(key=lambda c: c[0], reverse=True)
                picked_id, picked_row, picked_info = candidates[0]
                logger.info(
                    "report row: %d candidate(s) for '%s'; picked %s "
                    "(newest by ID)",
                    len(candidates), report_name, picked_info,
                )
                return picked_row
            time.sleep(1)
        raise TimeoutException(
            f"no report row matching '{report_name}' found within {timeout}s"
        )

    def _open_context_menu_for_report(
        self, report_name: str, max_attempts: int = 5
    ) -> None:
        """Right-click the newest matching report row, retrying on stale.

        wass rebuilds the report-list table DOM after every status change /
        refresh, so a ``<tr>`` element reference we just resolved can go stale
        before we finish dispatching the contextmenu event. The retry loop
        re-finds the row (via :meth:`_find_report_row`) on each attempt and
        fires the right-click in the **same** JS call that located the cell --
        so no stale Selenium element reference is ever passed across the
        boundary.

        It also waits for the jQuery contextMenu ``<ul class="context-menu">``
        to actually appear before returning, so the subsequent
        ``_click_context_menu_item('Result')`` doesn't race the menu open.
        """
        for attempt in range(1, max_attempts + 1):
            try:
                row = self._find_report_row(report_name, timeout=20)
                # Scroll into view first (real scroll, no element return).
                try:
                    self._scroll_into_view(row)
                except Exception:
                    pass
                time.sleep(0.3)
                # Try native Selenium context click first -- fastest path.
                try:
                    ActionChains(self._d).context_click(row).perform()
                    logger.info("context-clicked report row (attempt %d)", attempt)
                except Exception as e:
                    if "stale" in str(e).lower():
                        logger.debug(
                            "stale on attempt %d, re-finding row", attempt
                        )
                        time.sleep(0.5)
                        continue
                    raise
                # Confirm the menu actually opened. If not, retry.
                if self._wait_for_context_menu(timeout=3):
                    return
                logger.debug(
                    "context menu did not appear on attempt %d; retrying",
                    attempt,
                )
                time.sleep(0.5)
            except StaleElementReferenceException:
                logger.debug(
                    "stale element on attempt %d, re-finding row", attempt
                )
                time.sleep(0.5)
                continue
        raise TimeoutException(
            f"could not open context menu for '{report_name}' "
            f"after {max_attempts} attempts (list kept going stale)"
        )

    def _wait_for_context_menu(self, timeout: float = 3) -> bool:
        """Return True if a jQuery contextMenu ``<ul>`` is visible."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            els = self._d.find_elements(
                By.CSS_SELECTOR, "ul.context-menu-list"
            )
            for el in els:
                try:
                    if el.is_displayed():
                        return True
                except Exception:
                    continue
            time.sleep(0.2)
        return False

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
        """Open the Import picker, paste *names*, trigger the import, confirm.

        wass/Matrix42 flow (per inspected HTML):
          1. Click ``#ImportList`` icon button -> loads Host_Import.aspx into
             ``#Level3_Content`` (the right pane with the textarea + link).
          2. Wait for ``#ListResult`` textarea to appear, paste names into it.
          3. Click the ``a.EmTextLink[href*='Host_Import']`` link ("Import
             computer list") to actually submit the pasted list.
          4. Wait briefly for the right pane to list the imported computers,
             then click OK.
        """
        # Step 1: open the import picker.
        self._click("import_button")
        self._short_pause(2)

        # Step 2: wait for the textarea and paste the name list.
        # IMPORTANT: do NOT use textarea.send_keys() for thousands of lines --
        # Selenium sends one WebDriver command per character, which takes
        # minutes and often hangs the Edge driver mid-way. Instead set the
        # value via JavaScript (one round-trip) and dispatch 'input'/'change'
        # events so Matrix42's EmControls framework picks up the new value.
        # The text is passed as arguments[1] so WebDriver handles all escaping.
        by, val = self._loc("import_textarea")
        textarea = self._wait().until(
            EC.presence_of_element_located((by, val)),
            message="import textarea (#ListResult) not found",
        )
        text = "\n".join(names)
        self._d.execute_script(
            "var el = arguments[0], txt = arguments[1];"
            "el.value = txt;"
            "el.dispatchEvent(new Event('input', {bubbles:true}));"
            "el.dispatchEvent(new Event('change', {bubbles:true}));",
            textarea,
            text,
        )
        logger.info("pasted %d computer names into #ListResult (via JS)", len(names))

        # Step 3: click the "Import computer list" link to submit.
        self._click("import_computer_list_link")

        # Step 3b: submitting a large batch pops up a confirmation dialog
        # with an "OK" button. Click it to confirm. If no dialog appears
        # within the window, assume no confirmation was needed and continue.
        self._dismiss_import_confirmation(timeout=20)

        # Step 4: wait for the right pane to reflect the imported computers.
        # The server-side import of a large batch takes ~10-15s to process,
        # so give it room before reading the result list.
        time.sleep(15)
        logger.info("imported computer list submitted")

        # NOTE: do NOT click OK here. The import dialog has its own OK button
        # that *closes the whole New Report popup* (verified by the user --
        # clicking it dropped them back to the report-name page). The correct
        # flow is: after the import link submits, the wizard stays on the
        # report-name page and we proceed directly to clicking Next.

    def _dismiss_import_confirmation(self, timeout: float = 8) -> None:
        """Dismiss the "are you sure?" confirmation that pops up after
        submitting a large computer-name batch via Import computer list.

        wass/Matrix42 confirmation dialogs are rendered as a MsgBox overlay
        (see the host page's #MsgBox element) containing a single OK button
        implemented as a Matrix42 Button widget (``<div id="Button...">`` with
        text "OK" and an onclick of ``CloseMessageBox()``). Some confirmations
        may use Yes/Confirm instead. We blindly click whichever confirm button
        is visible. If nothing appears within *timeout* seconds, assume no
        confirmation was required and return silently.
        """
        # Candidate button texts (case-insensitive). "OK" is by far the most
        # common for Matrix42 message boxes.
        labels = ("OK", "Yes", "Confirm", "Continue", "Accept")
        deadline = time.time() + timeout
        while time.time() < deadline:
            for label in labels:
                xpath = (
                    f"//div[starts-with(@id,'Button')]"
                    f"[.//span[@class='ui-button-text' and normalize-space()='{label}']]"
                )
                try:
                    btns = self._d.find_elements(By.XPATH, xpath)
                except Exception:
                    btns = []
                for btn in btns:
                    try:
                        if not btn.is_displayed():
                            continue
                        # Prefer buttons whose onclick closes a message box
                        # (CloseMessageBox), but accept any visible one as a
                        # last resort -- the user said "blindly confirm".
                        btn_id = btn.get_attribute("id") or ""
                        script = ""
                        if btn_id:
                            script = self._d.execute_script(
                                "var b=document.getElementById(arguments[0]);"
                                "if(!b)return'';var s=b.nextElementSibling;"
                                "return s?s.textContent||'':'';",
                                btn_id,
                            ) or ""
                        self._scroll_into_view(btn)
                        try:
                            btn.click()
                        except Exception:
                            self._d.execute_script("arguments[0].click();", btn)
                        logger.info(
                            "dismissed import confirmation (clicked '%s'%s)",
                            label,
                            " ~ CloseMessageBox" if "CloseMessageBox" in script else "",
                        )
                        # Give the UI a beat to settle after dismissing.
                        self._short_pause(1)
                        return
                    except Exception:
                        continue
            time.sleep(0.5)
        logger.debug("no import confirmation dialog appeared within %.0fs", timeout)

    def select_result_elements(self, elements: List[str]) -> None:
        """Open the output-element picker, expand Login, tick the requested
        checkboxes, and close the picker with its OK button.

        From the inspected HTML:
          - #Output is the IconButton that opens the picker (Level3 popup).
          - The Login group is a row whose text is 'Login' with a Plus.png
            expand icon; clicking that row expands it to show the checkboxes.
          - Each option is a <tr> whose text label is in a <div class='EmTextClick'>
            and whose checkbox container is a <div class='EmCheckbox'> with an
            <input id='Choice_NNN'> inside. Clicking the EmCheckbox div (or the
            label) toggles it via EmCheckboxClicked().
          - The OK button that closes this picker has onclick containing
            "ClosePopup('Level3')".
        """
        # Open the picker.
        self._click("result_elements_button")
        self._short_pause(2)

        # Expand the Login group by clicking the expand icon in the Login row.
        # The row contains <img src='sym/Plus.png' alt='Open'>; clicking it (or
        # the Login text cell) toggles the subgroup open.
        try:
            login_row = self._wait(8).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//tr[td/div[normalize-space()='Login']]")
                ),
                message="login group row not found",
            )
            # Scroll the Login row into view inside the popup's scroll area
            # before clicking -- it may be below the fold.
            self._d.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", login_row
            )
            time.sleep(0.4)
            # Prefer the Plus.png icon; fall back to the Login text cell.
            try:
                plus = login_row.find_element(
                    By.XPATH, ".//img[contains(@src,'Plus.png') or @alt='Open']"
                )
                self._d.execute_script("arguments[0].click();", plus)
            except Exception:
                self._d.execute_script("arguments[0].click();", login_row)
            logger.info("clicked: login group expand")
        except Exception as e:
            logger.warning("could not expand Login group: %s", e)
        self._short_pause(2)

        # Tick each requested checkbox by matching the label text inside the
        # row's <div class='EmText EmTextClick'>, then toggling via Matrix42's
        # own CheckBoxToggle() function using the row's Choice id.
        #
        # Why CheckBoxToggle instead of clicking the EmCheckbox div:
        #   - The picker popup has its own scrollable content area. After
        #     expanding Login, the Last Logon / Last User / ... rows live
        #     below the fold. Even after scrollIntoView, the .click() on the
        #     EmCheckbox div sometimes does not register for the lower rows
        #     (the user observed the first two tick but the last two do not).
        #   - Matrix42 wires the toggle on EmCheckboxClicked(this) which reads
        #     the div's Choice_NNN id and flips currentstate. Calling
        #     CheckBoxToggle('Choice_NNN') directly does the same thing
        #     without depending on a DOM click landing on the right pixel.
        #   - We still scrollIntoView first so the row is visible (some
        #     Matrix42 builds refuse to toggle hidden inputs).
        for label in elements:
            try:
                # Find the label div whose text matches. Matching is
                # case-insensitive AND whitespace-insensitive: the wass HTML
                # uses "Last User (Account)" (space before "("), but callers
                # may pass "Last User(Account)". We normalize both sides by
                # lowercasing and stripping all spaces, then compare.
                key = label.lower().replace(" ", "")
                label_div = None
                deadline_lbl = time.time() + 10
                while time.time() < deadline_lbl:
                    candidates = self._d.find_elements(
                        By.XPATH, "//div[contains(@class,'EmTextClick')]"
                    )
                    for el in candidates:
                        try:
                            txt = (el.text or "").lower().replace(" ", "")
                        except Exception:
                            continue
                        if txt == key:
                            label_div = el
                            break
                    if label_div is not None:
                        break
                    time.sleep(0.3)
                if label_div is None:
                    raise TimeoutException(f"label '{label}' not found")
                # The EmCheckbox div is in the same <tr>, last <td>. It contains
                # <input id="Choice_NNN" ...> -- the id is the toggle key.
                row = label_div.find_element(By.XPATH, "./ancestor::tr[1]")
                cb_input = row.find_element(
                    By.XPATH, ".//div[contains(@class,'EmCheckbox')]//input[@type='checkbox']"
                )
                choice_id = cb_input.get_attribute("id") or ""
                # Scroll the row into view inside the popup's scroll area.
                self._d.execute_script(
                    "arguments[0].scrollIntoView({block:'center', inline:'nearest'});",
                    row,
                )
                time.sleep(0.4)
                if not cb_input.is_selected():
                    # Call Matrix42's own toggle function. This is the same
                    # call the row's onclick handlers make
                    # (CheckBoxToggle('Choice_NNN'); DisabledCheckAll(...)).
                    # DisabledCheckAll is cosmetic (greys out a "check all"
                    # box), so we skip it.
                    if choice_id:
                        self._d.execute_script(
                            "if (typeof CheckBoxToggle === 'function') {"
                            "  CheckBoxToggle(arguments[0]);"
                            "} else {"
                            "  var cb = document.getElementById(arguments[0]);"
                            "  if (cb) { cb.checked = !cb.checked; "
                            "    cb.dispatchEvent(new Event('change',{bubbles:true})); }"
                            "}",
                            choice_id,
                        )
                        logger.info("ticked: %s (via CheckBoxToggle('%s'))", label, choice_id)
                    else:
                        # Fallback: click the EmCheckbox div directly.
                        cb_div = row.find_element(
                            By.XPATH, ".//div[contains(@class,'EmCheckbox')]"
                        )
                        self._d.execute_script("arguments[0].click();", cb_div)
                        logger.info("ticked: %s (via div click fallback)", label)
                else:
                    logger.info("already ticked: %s", label)
            except Exception as e:
                logger.warning("could not tick '%s': %s", label, e)

        # Close the picker with the OK button whose onclick closes Level3.
        # If we can't find that exact button (e.g. the script registration
        # differs slightly), fall back to clicking any visible OK button --
        # at this point the only OK on screen should be the picker's.
        try:
            self._click_button_by_onclick(
                "result_elements_ok_button", "OK", "ClosePopup('Level3')"
            )
        except Exception as e:
            logger.warning(
                "could not find picker OK via ClosePopup('Level3') (%s); "
                "trying any visible OK button", e
            )
            self._click_any_visible_button(
                "result_elements_ok_button", "OK", timeout=10
            )

    def finish_wizard(self) -> None:
        """Click Step2 Next (-> Step3), then the generate-report OK.

        These two buttons have the same text as other Next/OK buttons in the
        wizard, so they are disambiguated by their onclick handlers:
          - Step2 Next: onclick contains 'Inv_Rep_Step3.aspx'
          - Generate-report OK: onclick contains 'Save_Report.aspx'

        Falls back to clicking any visible Next/OK if the onclick-specific
        match fails (at each step only the relevant button is on screen).
        """
        try:
            self._click_button_by_onclick(
                "next_step2_button", "Next", "Inv_Rep_Step3.aspx"
            )
        except Exception as e:
            logger.warning(
                "could not find Step2 Next via Inv_Rep_Step3.aspx (%s); "
                "trying any visible Next button", e
            )
            self._click_any_visible_button(
                "next_step2_button", "Next", timeout=10
            )
        self._short_pause(2)
        try:
            self._click_button_by_onclick(
                "generate_report_ok_button", "OK", "Save_Report.aspx"
            )
        except Exception as e:
            logger.warning(
                "could not find generate OK via Save_Report.aspx (%s); "
                "trying any visible OK button", e
            )
            self._click_any_visible_button(
                "generate_report_ok_button", "OK", timeout=10
            )

    def wait_for_completion(
        self,
        report_name: str,
        timeout: int = config.WASS_REPORT_TIMEOUT_SECONDS,
        poll_interval: int = config.WASS_POLL_INTERVAL_SECONDS,
    ) -> None:
        """Click the top-right refresh button until the report is done.

        We record the highest report ID present at submit time, then only
        trust status readings from rows whose ID is strictly greater than
        that -- this is the report we just launched. Older same-named reports
        (re-runs from earlier attempts) are ignored even if they share the
        exact name, so we never inherit a stale "successfully created"
        status from a previous run.
        """
        logger.info(
            "waiting for report '%s' to finish (timeout=%ds)", report_name, timeout
        )
        prev_max_id = self._max_report_id(report_name)
        logger.info(
            "report IDs present at submit: max=%s (new report must have ID > this)",
            prev_max_id if prev_max_id is not None else "(no matching rows yet)",
        )
        deadline = time.time() + timeout
        last_status = None
        while time.time() < deadline:
            self._refresh()
            status = self._read_status(report_name, min_id=prev_max_id)
            if status != last_status:
                logger.info("status: %s", status or "(unknown)")
                last_status = status
            if status and STATUS_DONE.lower() in status.lower():
                logger.info("report '%s' finished", report_name)
                # Give the report-list table a beat to settle after the last
                # status change -- Kendo/jQuery rebuilds the <tr> nodes, and
                # any element reference we hold goes stale during the rebuild.
                # download_result re-finds the row in a retry loop, but a short
                # pause here avoids the first attempt almost always failing.
                time.sleep(2)
                return
            time.sleep(poll_interval)
        raise TimeoutError(
            f"Report '{report_name}' did not finish within {timeout}s "
            f"(last status: {last_status})"
        )

    def _max_report_id(self, report_name: str):
        """Return the highest report ID among rows named *report_name*, or
        ``None`` if no matching row exists yet."""
        try:
            rows = self._d.find_elements(
                By.XPATH, "//tr[@cr='true' and @cm]"
            )
            best = None
            for r in rows:
                try:
                    tds = r.find_elements(By.TAG_NAME, "td")
                    if len(tds) < 3:
                        continue
                    if not any(
                        (td.text or "").strip() == report_name for td in tds
                    ):
                        continue
                    cp = r.get_attribute("cp") or ""
                    rid = -1
                    parts = cp.split("|")
                    if len(parts) >= 2:
                        try:
                            rid = int(parts[1].strip())
                        except ValueError:
                            pass
                    if rid < 0:
                        try:
                            rid = int((tds[1].text or "").strip())
                        except ValueError:
                            rid = -1
                    if rid >= 0 and (best is None or rid > best):
                        best = rid
                except Exception:
                    continue
            return best
        except Exception:
            return None

    def _refresh(self) -> None:
        """Click the small top-right refresh button; fall back to F5."""
        try:
            self._click("refresh_button", timeout=5)
        except Exception:
            try:
                self._d.refresh()
            except Exception:
                pass

    def _read_status(self, report_name: str, min_id: Optional[int] = None) -> str:
        """Read the status of the *newest* report named *report_name*.

        The report list accumulates every report the user ever created, so a
        name-substring match (the old approach) reads the status of an older
        same-named report that is already 'successfully created' -- making
        wait_for_completion return instantly while the report we just
        launched is still running. We instead locate the newest matching row
        by report ID (monotonic) and read *that* row's icon.

        If *min_id* is given, only rows with ID strictly greater than
        *min_id* are considered -- this lets wait_for_completion ignore every
        same-named report that already existed at submit time, so a freshly
        submitted run is never mistaken for an older finished one.

        Status signal: only the row's own ``<div class="EmIcon">``
        ``background-image`` URL containing ``Report_Success`` / ``success``
        counts as :data:`STATUS_DONE`. Any other icon (or no icon) is treated
        as :data:`STATUS_RUNNING` -- we never inherit a stale "done" status
        from a neighbouring element's text. See :meth:`_read_row_status`.

        Returns "" if no matching row is found (e.g. the new row has not been
        rendered yet right after the wizard submitted).
        """
        try:
            rows = self._d.find_elements(
                By.XPATH, "//tr[@cr='true' and @cm]"
            )
            best_id = -1
            best_row = None
            for r in rows:
                try:
                    tds = r.find_elements(By.TAG_NAME, "td")
                    if len(tds) < 3:
                        continue
                    if not any(
                        (td.text or "").strip() == report_name for td in tds
                    ):
                        continue
                    cp = r.get_attribute("cp") or ""
                    rid = -1
                    parts = cp.split("|")
                    if len(parts) >= 2:
                        try:
                            rid = int(parts[1].strip())
                        except ValueError:
                            pass
                    if rid < 0:
                        try:
                            rid = int((tds[1].text or "").strip())
                        except ValueError:
                            rid = -1
                    # Ignore rows that already existed at submit time.
                    if min_id is not None and rid >= 0 and rid <= min_id:
                        continue
                    if rid > best_id:
                        best_id = rid
                        best_row = r
                except Exception:
                    continue
            if best_row is None:
                return ""
            return self._read_row_status(best_row)
        except Exception:
            return ""

    def _read_row_status(self, row) -> str:
        """Read the status signal from a single report ``<tr>``.

        Only the row's own ``<div class="EmIcon">`` background-image URL is
        trusted as a DONE signal -- everything else (row text, parent text)
        is treated as still RUNNING. This is intentional: wass shows status
        purely via the icon, and the row's computed text can inadvertently
        include stray "successfully created" strings from neighbouring
        elements / toast notifications, which previously caused
        wait_for_completion to return ~1 second after submit while the
        report was still generating.
        """
        try:
            icons = row.find_elements(By.CSS_SELECTOR, "div.EmIcon")
            for icon in icons:
                try:
                    style = icon.get_attribute("style") or ""
                except Exception:
                    continue
                low = style.lower()
                # Only the explicit Success icon means done.
                if "report_success" in low or "success" in low:
                    return STATUS_DONE
                # Running / processing / waiting icons.
                if "running" in low or "process" in low or "wait" in low:
                    return STATUS_RUNNING
                # Any other icon (e.g. New, Queued) -- still in flight.
                return STATUS_RUNNING
            # No icon at all -- row may be mid-render. Treat as running so we
            # keep polling instead of falsely returning done.
            return STATUS_RUNNING
        except Exception:
            return ""

    def download_result(self, report_name: str) -> Optional[Path]:
        """Right-click the entry -> Result -> Excel Download -> download link.

        Returns the path of the newly downloaded .xlsx file.

        From the inspected wass HTML:
          - The report list shows one row per report (name + status). Right-
            clicking a row opens a context menu.
          - "Result" is a context-menu item with that exact text.
          - Clicking Result opens a popup whose body contains an "Excel
            Download" cell (<td><div class='EmText'>Excel Download</div></td>).
          - Clicking Excel Download triggers generation of the file and
            replaces the popup body with a success message containing an
            <a class='EmTextLink' href='.../*.xlsx'> link.
          - Clicking that link downloads the .xlsx.
        """
        before = set(self.download_dir.glob("*.xlsx"))
        logger.info("opening Result for '%s'", report_name)

        # The wass report list is rebuilt by jQuery/Kendo every time the
        # status changes (and every time we click the refresh button in
        # wait_for_completion). Any <tr> reference we hold goes stale within
        # milliseconds of the rebuild, so ActionChains.context_click(row)
        # throws StaleElementReferenceException -- and so does the JS fallback
        # because it still passes the stale element as arguments[0].
        #
        # Fix: re-find the row inside a retry loop, and immediately
        # right-click the FIRST <td> (the icon cell) of that freshly-found
        # row via a single JS call -- no element argument is passed across
        # the stale boundary, the JS re-resolves the row by name+max ID.
        self._open_context_menu_for_report(report_name)
        self._short_pause(1)

        # Click "Result" in the context menu. Use the dedicated helper instead
        # of _click: jQuery contextMenu <li> items often fail Selenium's
        # element_to_be_clickable check, and a bare text()='Result' XPath would
        # match a same-named column header elsewhere on the page first.
        try:
            self._click_context_menu_item("result_menu_item", "Result", timeout=10)
        except Exception as e:
            logger.warning(
                "could not click 'Result' via context-menu helper (%s); "
                "trying generic locator", e
            )
            self._click("result_menu_item", timeout=10)
        self._short_pause(2)

        # Click "Excel Download".
        self._click("excel_download_button", timeout=15)
        logger.info("clicked: Excel Download")
        self._short_pause(3)

        # Click the success-link to download the .xlsx.
        try:
            self._click("download_link", timeout=30)
            logger.info("clicked: download link")
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
        # After import submits, click Step1 Next (onclick -> Inv_Rep_Step2.aspx)
        # to advance to the Result-elements page.
        self._click_button_by_onclick(
            "next_step1_button", "Next", "Inv_Rep_Step2.aspx"
        )
        self._short_pause(2)
        self.select_result_elements(result_elements)
        self.finish_wizard()
        self.wait_for_completion(report_name)
        return self.download_result(report_name)
