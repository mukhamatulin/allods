import json
from contextlib import ExitStack
from time import monotonic
from typing import Any
from urllib.error import HTTPError

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, Response, sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError


class BrowserRatingsClient:
    """Read official ratings in an isolated Edge session with JavaScript enabled."""

    def __init__(self) -> None:
        self._stack: ExitStack | None = None
        self._page: Page | None = None
        self._response: Response | None = None

    def __enter__(self) -> "BrowserRatingsClient":
        stack = ExitStack()
        try:
            playwright = stack.enter_context(sync_playwright())
            browser = playwright.chromium.launch(channel="msedge", headless=True)
            stack.callback(browser.close)
            self._page = browser.new_page()
            self._page.on("response", self._record_response)
        except PlaywrightError as exc:
            stack.close()
            raise RuntimeError(
                "Could not start Microsoft Edge. Install Microsoft Edge and run "
                "python -m pip install -r requirements.txt using the export's Python."
            ) from exc
        self._stack = stack
        return self

    def __exit__(self, *args: Any) -> None:
        try:
            if self._stack is not None:
                self._stack.__exit__(*args)
        finally:
            self._stack = None
            self._page = None
            self._response = None

    def _record_response(self, response: Response) -> None:
        if (
            self._page is not None
            and response.request.is_navigation_request()
            and response.frame == self._page.main_frame
        ):
            self._response = response

    def fetch_json(self, url: str, timeout: int) -> Any:
        if timeout <= 0:
            raise ValueError("--timeout must be greater than zero")
        if self._page is None:
            raise RuntimeError("BrowserRatingsClient must be used as a context manager")

        deadline = monotonic() + timeout
        self._response = None
        try:
            self._page.goto(url, wait_until="domcontentloaded", timeout=timeout * 1000)
            # The site's script redirects after its check. Let Edge run it normally;
            # keep one session so subsequent ratings requests retain its cookies.
            self._page.wait_for_function(
                "() => typeof window.bp_chl_vars === 'undefined'",
                timeout=max(1, (deadline - monotonic()) * 1000),
            )
            self._page.wait_for_load_state(
                "domcontentloaded", timeout=max(1, (deadline - monotonic()) * 1000)
            )
            response = self._response
            if response is None:
                raise RuntimeError(f"No HTTP response received: {url}")
            if not response.ok:
                error = HTTPError(url, response.status, response.status_text, None, None)
                error.close()
                raise error
            body = response.text()
        except PlaywrightTimeoutError as exc:
            raise RuntimeError(
                f"Timed out after {timeout}s waiting for ratings or the security check: "
                f"{url}. Open allods.ru in a browser to check whether manual action is required."
            ) from exc
        except PlaywrightError as exc:
            reason = str(exc).splitlines()[0]
            raise RuntimeError(f"Browser request failed: {url}: {reason}") from exc

        # Read the raw response: HTML rendering changes '<' and '&' in JSON strings.
        try:
            return json.loads(body)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"Expected ratings JSON, received a non-JSON page: {url}") from exc
