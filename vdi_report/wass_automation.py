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
from selenium.common.exceptions import TimeoutException

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
    # "Result" is a context-menu item whose text is exactly "Result".
    "result_menu_item": (
        By.XPATH,
        "//*[normalize-space(translate(text(),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='result']"
        " | //td[div[normalize-space()='Result']]",
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

        # Locate the report row. Prefer a <tr> or <td> that contains the
        # report name as its own text (not a giant container like <body>).
        row_xpaths = [
            f"//tr[td[contains(.,'{report_name}')]]",
            f"//td[normalize-space()='{report_name}']",
            f"//*[contains(.,'{report_name}') and not(*)]",
            f"//*[contains(.,'{report_name}')]",
        ]
        row = None
        for xp in row_xpaths:
            try:
                row = self._wait(5).until(
                    EC.presence_of_element_located((By.XPATH, xp)),
                    message=f"report row via {xp[:40]}",
                )
                break
            except Exception:
                continue
        if row is None:
            raise TimeoutException(f"report row '{report_name}' not found")

        # Right-click the row to open the context menu.
        try:
            self._scroll_into_view(row)
            ActionChains(self._d).context_click(row).perform()
            logger.info("context-clicked report row")
        except Exception as e:
            logger.warning("context_click failed (%s); trying JS", e)
            self._d.execute_script(
                "var el = arguments[0];"
                "var ev = document.createEvent('MouseEvents');"
                "ev.initMouseEvent('contextmenu', true, true, window,"
                "1, 0, 0, 0, 0, false, false, false, false, 2, null);"
                "el.dispatchEvent(ev);",
                row,
            )
        self._short_pause(1)

        # Click "Result" in the context menu.
        self._click("result_menu_item", timeout=10)
        logger.info("clicked: Result (context menu)")
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
