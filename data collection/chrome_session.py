"""
chrome_session.py — A self-healing undetected-chromedriver wrapper.

The Chrome session can die mid-scrape (browser crash, "invalid session id").
BrowserSession is a transparent proxy over uc.Chrome: every attribute access
(title, page_source, get, execute_async_script, ...) forwards to the underlying
driver unchanged, so existing call sites work without modification.

Crash recovery is centralized: callers that navigate (see simbli._navigate)
call ensure_alive() before driver.get(), which restarts Chrome if the session
is dead. Every Simbli browser action is preceded by such a navigate, so a dead
session is always revived before the next real request.
"""
import undetected_chromedriver as uc


class BrowserSession:
    """Transparent proxy over uc.Chrome that relaunches on a dead session."""

    def __init__(self, version_main: int = 149, page_load_timeout: int = 30):
        self.version_main = version_main
        self.page_load_timeout = page_load_timeout
        self._driver = None
        self._create()

    def _create(self):
        """(Re)create the underlying driver. Quitting the old one first if any."""
        if self._driver is not None:
            try:
                self._driver.quit()
            except Exception:
                pass
        options = uc.ChromeOptions()
        self._driver = uc.Chrome(options=options, version_main=self.version_main)
        self._driver.set_page_load_timeout(self.page_load_timeout)
        return self._driver

    def is_alive(self) -> bool:
        """True if the session still responds to a lightweight command."""
        if self._driver is None:
            return False
        try:
            _ = self._driver.current_url  # raises if session is dead
            return True
        except Exception:
            return False

    def ensure_alive(self):
        """Return a live driver, relaunching Chrome if the session is dead."""
        if not self.is_alive():
            print("      [Browser] Session dead — restarting Chrome...")
            self._create()
            print("      [Browser] Chrome restarted successfully.")
        return self._driver

    def quit(self):
        """Quit the underlying driver (wrapper's own cleanup)."""
        if self._driver is not None:
            try:
                self._driver.quit()
            except Exception:
                pass
            self._driver = None

    def __getattr__(self, name):
        """Forward any other attribute access to the underlying driver.

        NOTE: this is a dumb passthrough and does NOT auto-restart. Restart
        happens at the navigate chokepoint (ensure_alive in simbli._navigate).
        """
        # __getattr__ is only invoked when normal lookup fails, so the explicit
        # methods/attributes above (_driver, ensure_alive, quit, etc.) are not
        # proxied here. Everything else (get, title, page_source, ...) is.
        driver = self.__dict__.get("_driver")
        if driver is None:
            raise AttributeError(name)
        return getattr(driver, name)
