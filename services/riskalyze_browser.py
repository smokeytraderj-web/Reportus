"""Interactive, local Riskalyze capture using a dedicated persistent browser profile."""

from __future__ import annotations

import csv
import os
import re
import tempfile
import time
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from urllib.parse import urljoin

from security.portal_capture import PortalCaptureError, crop_riskalyze_analytics


RISKALYZE_CLIENTS_URL = "https://pro.riskalyze.com/advisor/clients"
_SEARCH_SELECTORS = (
    'input[placeholder="Search"]',
    'main input[placeholder="Search"]',
    'input[placeholder*="search" i]',
    'input[aria-label*="search" i]',
)


class RiskalyzeCaptureError(RuntimeError):
    """Riskalyze could not produce a verified, privacy-safe portfolio capture."""


@dataclass(frozen=True, slots=True)
class RiskalyzeCaptureResult:
    source_path: Path
    preview_path: Path
    matched_name: str


@dataclass(frozen=True, slots=True)
class RiskalyzePortfolioData:
    portfolio_total: str
    risk_number: str
    allocation: tuple[tuple[str, Decimal], ...]
    historical_loss: str
    historical_loss_percent: str
    historical_gain: str
    historical_gain_percent: str
    annual_dividend: str
    max_drawdown: str
    annual_range_midpoint: str
    expense_ratio: str = ""
    portfolio_costs: str = ""


def _required_match(pattern: str, text: str, label: str) -> str:
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if match is None:
        raise RiskalyzeCaptureError(f"Riskalyze did not expose {label}. Refresh the page and try again.")
    return match.group(1).strip()


def parse_riskalyze_portfolio(text: str) -> RiskalyzePortfolioData:
    """Parse only the approved Current Portfolio summary fields from visible DOM text."""

    compact = text.replace("\xa0", " ")
    portfolio_total = _required_match(
        r"PORTFOLIO\s+TOTAL(?:\s+@)?\s*(\$[\d,]+(?:\.\d{2})?)",
        compact,
        "the portfolio total",
    )
    risk_number = _required_match(
        r"PORTFOLIO\s+TOTAL\s*\$[\d,]+(?:\.\d{2})?\s*(?:RISK\s*)?(\d{1,3})\s*ANALYTICS",
        compact,
        "the overall Risk Number",
    )
    allocation = []
    for label in ("Stocks", "Bonds", "Other", "Cash"):
        match = re.search(rf"\b{label}\b\s*(\d+(?:\.\d+)?)%", compact, re.IGNORECASE)
        if match is not None:
            allocation.append((label, Decimal(match.group(1))))
    if len(allocation) < 2:
        raise RiskalyzeCaptureError(
            "Riskalyze did not expose enough allocation values. Refresh the portfolio and try again."
        )

    range_block = _required_match(
        r"95%\s+Historical\s+Range(?:\s*\([^)]*\))?(.*?)(?:\bStocks\b|Riskalyze\s+GPA|Annual\s+Dividend)",
        compact,
        "the historical range",
    )
    historical_loss = _required_match(r"(-\$[\d,]+)", range_block, "historical loss")
    historical_gain = _required_match(r"(\+\$[\d,]+)", range_block, "historical gain")
    historical_loss_percent = _required_match(
        r"(-\d+(?:\.\d+)?%)", range_block, "historical loss percentage"
    )
    historical_gain_percent = _required_match(
        r"(\+\d+(?:\.\d+)?%)", range_block, "historical gain percentage"
    )
    annual_dividend = _required_match(
        r"Annual\s+Dividend\s*(\d+(?:\.\d+)?%)", compact, "annual dividend"
    )
    max_drawdown = _required_match(
        r"Max\s+Drawdown\s*(-\d+(?:\.\d+)?%)", compact, "maximum drawdown"
    )
    annual_range_midpoint = _required_match(
        r"Annual\s+Range\s+Midpoint\s*(-?\d+(?:\.\d+)?%)",
        compact,
        "annual range midpoint",
    )
    expense_match = re.search(
        r"(?:Expense\s+Ratio|Portfolio\s+Costs)\s*(\d+(?:\.\d+)?%)",
        compact,
        re.IGNORECASE,
    )
    costs_match = re.search(
        r"Portfolio\s+Costs\s*(?:0%\s*1%\s*)?(\d+\.\d+%)",
        compact,
        re.IGNORECASE,
    )
    return RiskalyzePortfolioData(
        portfolio_total=portfolio_total,
        risk_number=risk_number,
        allocation=tuple(allocation),
        historical_loss=historical_loss,
        historical_loss_percent=historical_loss_percent,
        historical_gain=historical_gain,
        historical_gain_percent=historical_gain_percent,
        annual_dividend=annual_dividend,
        max_drawdown=max_drawdown,
        annual_range_midpoint=annual_range_midpoint,
        expense_ratio=(
            expense_match.group(1)
            if expense_match is not None and "Expense" in expense_match.group(0)
            else ""
        ),
        portfolio_costs=costs_match.group(1) if costs_match is not None else "",
    )


def write_riskalyze_source(data: RiskalyzePortfolioData, destination: Path) -> Path:
    """Write a privacy-safe, numerically grounded source without account-level data."""

    total = Decimal(data.portfolio_total.replace("$", "").replace(",", ""))
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(("Section", "Metric", "Value", "Weight"))
        for label, percent in data.allocation:
            value = (total * percent / Decimal("100")).quantize(Decimal("0.01"))
            writer.writerow(("Allocation", label, f"{value:.2f}", f"{percent:.2f}%"))
        for label, value in (
            ("Portfolio total", data.portfolio_total),
            ("Risk", data.risk_number),
            ("Historical loss", data.historical_loss),
            ("Historical loss %", data.historical_loss_percent),
            ("Historical gain", data.historical_gain),
            ("Historical gain %", data.historical_gain_percent),
            ("Annual dividend", data.annual_dividend),
            ("Max drawdown", data.max_drawdown),
            ("Annual range midpoint", data.annual_range_midpoint),
            ("Expense ratio", data.expense_ratio),
            ("Portfolio costs", data.portfolio_costs),
        ):
            if value:
                writer.writerow(("Risk", label, value, ""))
    return destination


def default_riskalyze_profile_directory() -> Path:
    """Return a stable, app-owned profile directory without using the daily browser profile."""

    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if local_app_data:
        return Path(local_app_data) / "Reporticles" / "BrowserProfiles" / "Riskalyze"
    return Path.home() / ".reporticles" / "browser-profiles" / "riskalyze"


def _normalize_household(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", text.casefold())


def exact_client_href(client_name: str, candidates: Iterable[tuple[str, str]]) -> str:
    """Return one exact household link or fail rather than selecting a similar client."""

    target = _normalize_household(client_name)
    matches: dict[str, str] = {}
    for text, href in candidates:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if any(_normalize_household(line) == target for line in lines):
            matches[href] = text
    if not matches:
        raise RiskalyzeCaptureError(
            "No exact Riskalyze household match was found. Check the client name and try again."
        )
    if len(matches) > 1:
        raise RiskalyzeCaptureError(
            "More than one exact Riskalyze household match was found. Refine the client name."
        )
    return next(iter(matches))


class RiskalyzeBrowserCapture:
    """Search an exact household and capture its Current Portfolio analytics locally."""

    def __init__(
        self,
        *,
        profile_directory: Path | None = None,
        login_timeout_seconds: int = 5 * 60,
    ):
        self.profile_directory = profile_directory or default_riskalyze_profile_directory()
        self.login_timeout_seconds = login_timeout_seconds

    def capture(
        self,
        client_name: str,
        destination: Path,
        *,
        status: Callable[[str], None] | None = None,
    ) -> RiskalyzeCaptureResult:
        client_name = client_name.strip()
        if not client_name:
            raise RiskalyzeCaptureError("Enter the client or household name first.")
        notify = status or (lambda _message: None)
        try:
            from playwright.sync_api import Error as PlaywrightError
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise RiskalyzeCaptureError(
                "Riskalyze browser support is not installed. Run: pip install -r requirements.txt"
            ) from exc

        self.profile_directory.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="reporticles-riskalyze-raw-") as temporary:
            raw_capture = Path(temporary) / "current-portfolio.png"
            preview_path = destination.with_suffix(".png")
            try:
                with sync_playwright() as playwright:
                    context = self._launch_context(playwright, PlaywrightError)
                    try:
                        page = context.pages[0] if context.pages else context.new_page()
                        page.set_viewport_size({"width": 1920, "height": 1080})
                        notify("Riskalyze opened. Sign in and complete MFA if requested.")
                        page.goto(RISKALYZE_CLIENTS_URL, wait_until="domcontentloaded", timeout=60_000)
                        search = self._wait_for_client_search(page)
                        notify("Searching for the exact Riskalyze household…")
                        search.fill(client_name)
                        page.wait_for_timeout(1_500)
                        candidates = self._client_candidates(page)
                        client_href = exact_client_href(client_name, candidates)
                        page.goto(urljoin(RISKALYZE_CLIENTS_URL, client_href), wait_until="domcontentloaded")
                        portfolio_href = self._portfolio_href(page)
                        page.goto(urljoin(RISKALYZE_CLIENTS_URL, portfolio_href), wait_until="domcontentloaded")
                        page.get_by_text("PORTFOLIO TOTAL", exact=True).wait_for(
                            state="visible", timeout=60_000
                        )
                        body_text = page.locator("body").inner_text(timeout=30_000)
                        if "PORTFOLIO TOTAL" not in body_text.upper() or "RISK" not in body_text.upper():
                            raise RiskalyzeCaptureError(
                                "The Current Portfolio analytics panel did not finish loading."
                            )
                        notify("Portfolio found. Zooming out and capturing the analytics…")
                        page.evaluate("document.documentElement.style.zoom = '80%'")
                        page.wait_for_timeout(1_000)
                        page.screenshot(path=str(raw_capture), full_page=False, animations="disabled")
                    finally:
                        context.close()
            except RiskalyzeCaptureError:
                raise
            except PlaywrightTimeoutError as exc:
                raise RiskalyzeCaptureError(
                    "Riskalyze timed out. Finish sign-in or MFA in the opened browser, then try again."
                ) from exc
            except PlaywrightError as exc:
                raise RiskalyzeCaptureError(
                    "Riskalyze browser automation could not continue. Close other Reporticles browser windows and try again."
                ) from exc

            try:
                data = parse_riskalyze_portfolio(body_text)
                crop_riskalyze_analytics(raw_capture, preview_path)
                write_riskalyze_source(data, destination)
            except PortalCaptureError as exc:
                raise RiskalyzeCaptureError(str(exc)) from exc
        notify("Riskalyze capture verified and ready.")
        return RiskalyzeCaptureResult(destination, preview_path, client_name)

    def _launch_context(self, playwright, playwright_error):
        options = {
            "user_data_dir": str(self.profile_directory),
            "headless": False,
            "viewport": {"width": 1920, "height": 1080},
            "accept_downloads": False,
            "args": [
                "--disable-features=PasswordManagerOnboarding,PasswordLeakDetection",
            ],
        }
        errors = []
        for channel in (None, "chrome", "msedge"):
            try:
                return playwright.chromium.launch_persistent_context(
                    channel=channel, **options
                )
            except playwright_error as exc:
                errors.append(exc)
        raise RiskalyzeCaptureError(
            "No compatible browser could be opened. Install Chromium with: python -m playwright install chromium"
        ) from errors[-1]

    def _wait_for_client_search(self, page):
        deadline = time.monotonic() + self.login_timeout_seconds
        while time.monotonic() < deadline:
            if page.is_closed():
                raise RiskalyzeCaptureError("The Riskalyze browser was closed before capture finished.")
            for selector in _SEARCH_SELECTORS:
                locator = page.locator(selector).first
                if locator.count() and locator.is_visible():
                    return locator
            page.wait_for_timeout(750)
        raise RiskalyzeCaptureError(
            "Riskalyze sign-in was not completed within five minutes. Try Fetch from Riskalyze again."
        )

    @staticmethod
    def _client_candidates(page) -> tuple[tuple[str, str], ...]:
        links = page.locator('a[href*="/client-details/"]')
        candidates = []
        for index in range(links.count()):
            link = links.nth(index)
            href = link.get_attribute("href") or ""
            text = link.inner_text().strip()
            if href and text:
                candidates.append((text, href))
        return tuple(candidates)

    @staticmethod
    def _portfolio_href(page) -> str:
        links = page.locator('a[href*="/portfolio"]')
        options = []
        for index in range(links.count()):
            link = links.nth(index)
            href = link.get_attribute("href") or ""
            text = link.inner_text().strip()
            if href and "portfolio" in text.casefold():
                options.append(href)
        if not options:
            raise RiskalyzeCaptureError("The Riskalyze Portfolio tab could not be found.")
        return options[0]
