#!/usr/bin/env python3
"""
Web automation for downloading financial data from Investing.com Pro.

Uses Playwright for reliable browser automation with:
- Automatic waiting for elements
- Built-in download handling
- Headless mode support
"""

import asyncio
import re
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

try:
    from playwright.async_api import async_playwright, Browser, Page, Download
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False


@dataclass
class CompanyInfo:
    """Information about a company found on Investing.com."""
    name: str
    ticker: str
    exchange: str
    url: str
    country: Optional[str] = None


@dataclass
class DownloadResult:
    """Result of downloading financial statements."""
    income_statement_path: Optional[Path] = None
    balance_sheet_path: Optional[Path] = None
    company_info: Optional[CompanyInfo] = None
    error: Optional[str] = None


class InvestingComScraper:
    """
    Automates login and financial data download from Investing.com Pro.

    Usage:
        async with InvestingComScraper(download_dir="/path/to/downloads") as scraper:
            await scraper.login(email, password)
            result = await scraper.download_financials("7859", exchange="TYO")
    """

    BASE_URL = "https://www.investing.com"
    LOGIN_URL = "https://www.investing.com/members-admin/login"
    SEARCH_URL = "https://www.investing.com/search/?q={query}"

    def __init__(
        self,
        download_dir: str | Path = None,
        headless: bool = True,
        timeout: int = 30000,
    ):
        """
        Initialize the scraper.

        Args:
            download_dir: Directory to save downloaded files. Defaults to current directory.
            headless: Run browser in headless mode (no visible window).
            timeout: Default timeout for operations in milliseconds.
        """
        if not PLAYWRIGHT_AVAILABLE:
            raise ImportError(
                "Playwright not installed. Install with:\n"
                "  pip install playwright\n"
                "  playwright install chromium"
            )

        self.download_dir = Path(download_dir) if download_dir else Path.cwd()
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.headless = headless
        self.timeout = timeout

        self._playwright = None
        self._browser: Optional[Browser] = None
        self._page: Optional[Page] = None
        self._logged_in = False

    async def __aenter__(self):
        """Async context manager entry."""
        await self._start_browser()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """Async context manager exit."""
        await self.close()

    async def _start_browser(self):
        """Start the browser."""
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(
            headless=self.headless,
        )
        context = await self._browser.new_context(
            accept_downloads=True,
            viewport={"width": 1280, "height": 800},
        )
        self._page = await context.new_page()
        self._page.set_default_timeout(self.timeout)

    async def close(self):
        """Close the browser and cleanup."""
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        self._browser = None
        self._page = None
        self._playwright = None
        self._logged_in = False

    async def login(self, email: str, password: str) -> bool:
        """
        Login to Investing.com.

        Args:
            email: Account email address
            password: Account password

        Returns:
            True if login successful

        Raises:
            Exception: If login fails
        """
        if not self._page:
            raise RuntimeError("Browser not started. Use 'async with' context manager.")

        print("Navigating to Investing.com...")
        await self._page.goto(self.BASE_URL)

        # Wait for page to load and look for sign in button
        await self._page.wait_for_load_state("networkidle")

        # Try to find and click sign in button
        # The site may have different layouts, so try multiple selectors
        sign_in_selectors = [
            'a[data-test="login-btn"]',
            'a:has-text("Sign In")',
            'button:has-text("Sign In")',
            '.login-btn',
            '#loginBtn',
        ]

        clicked = False
        for selector in sign_in_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    await element.click()
                    clicked = True
                    break
            except Exception:
                continue

        if not clicked:
            # Try navigating directly to login page
            print("Sign in button not found, navigating to login page directly...")
            await self._page.goto(self.LOGIN_URL)

        # Wait for login form
        await self._page.wait_for_load_state("networkidle")

        # Fill in credentials - try multiple possible selectors
        email_selectors = [
            'input[name="email"]',
            'input[type="email"]',
            '#email',
            'input[placeholder*="email" i]',
            'input[data-test="email-input"]',
        ]

        password_selectors = [
            'input[name="password"]',
            'input[type="password"]',
            '#password',
            'input[placeholder*="password" i]',
            'input[data-test="password-input"]',
        ]

        # Fill email
        email_filled = False
        for selector in email_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    await element.fill(email)
                    email_filled = True
                    print("Email entered.")
                    break
            except Exception:
                continue

        if not email_filled:
            raise Exception("Could not find email input field")

        # Fill password
        password_filled = False
        for selector in password_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    await element.fill(password)
                    password_filled = True
                    print("Password entered.")
                    break
            except Exception:
                continue

        if not password_filled:
            raise Exception("Could not find password input field")

        # Click submit button
        submit_selectors = [
            'button[type="submit"]',
            'button:has-text("Sign In")',
            'button:has-text("Log In")',
            'input[type="submit"]',
            'button[data-test="login-submit"]',
        ]

        submitted = False
        for selector in submit_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    await element.click()
                    submitted = True
                    print("Submitting login...")
                    break
            except Exception:
                continue

        if not submitted:
            raise Exception("Could not find submit button")

        # Wait for navigation after login
        await self._page.wait_for_load_state("networkidle")
        await asyncio.sleep(2)  # Extra wait for any redirects

        # Check if login was successful by looking for user-specific elements
        # or checking if we're still on login page
        current_url = self._page.url

        if "login" in current_url.lower() or "signin" in current_url.lower():
            # Still on login page - check for error messages
            error_selectors = [
                '.error-message',
                '.alert-danger',
                '[data-test="login-error"]',
                '.login-error',
            ]

            for selector in error_selectors:
                try:
                    element = await self._page.query_selector(selector)
                    if element and await element.is_visible():
                        error_text = await element.text_content()
                        raise Exception(f"Login failed: {error_text}")
                except Exception as e:
                    if "Login failed" in str(e):
                        raise
                    continue

            raise Exception("Login appears to have failed - still on login page")

        self._logged_in = True
        print("Login successful!")
        return True

    async def search_company(self, ticker: str, exchange: str = None) -> Optional[CompanyInfo]:
        """
        Search for a company by ticker symbol.

        Args:
            ticker: Stock ticker symbol (e.g., "7859", "AAPL")
            exchange: Optional exchange code to filter results (e.g., "TYO", "NASDAQ")

        Returns:
            CompanyInfo if found, None otherwise
        """
        if not self._page:
            raise RuntimeError("Browser not started")

        search_query = f"{ticker} {exchange}" if exchange else ticker
        search_url = self.SEARCH_URL.format(query=search_query)

        print(f"Searching for: {search_query}")
        await self._page.goto(search_url)
        await self._page.wait_for_load_state("networkidle")

        # Look for search results - specifically equity results
        # The search page typically shows results in sections (Equities, ETFs, etc.)
        result_selectors = [
            '.search-results-quotes a',
            '.js-inner-all-results-quotes-wrapper a',
            '[data-test="search-result-quote"] a',
            '.searchSectionMain a.js-quote-link',
        ]

        for selector in result_selectors:
            try:
                results = await self._page.query_selector_all(selector)
                for result in results:
                    try:
                        href = await result.get_attribute("href")
                        text = await result.text_content()

                        if not href or "/equities/" not in href:
                            continue

                        # Check if exchange matches (if specified)
                        if exchange:
                            exchange_lower = exchange.lower()
                            # Check text content or URL for exchange
                            text_lower = text.lower() if text else ""
                            href_lower = href.lower()

                            # Map common exchange codes to what might appear in results
                            exchange_patterns = {
                                "tyo": ["tokyo", "tyo", "japan", "tse"],
                                "hkg": ["hong kong", "hkg", "hkex"],
                                "nyse": ["nyse", "new york"],
                                "nasdaq": ["nasdaq"],
                                "lon": ["london", "lse", "lon"],
                                "epa": ["paris", "epa", "euronext paris"],
                            }

                            patterns = exchange_patterns.get(exchange_lower, [exchange_lower])
                            matches_exchange = any(p in text_lower or p in href_lower for p in patterns)

                            if not matches_exchange:
                                continue

                        # Extract company info
                        name = text.strip() if text else ""
                        # Clean up the name (remove exchange info if present)
                        name = re.sub(r'\s*\([^)]+\)\s*$', '', name)

                        company_url = href if href.startswith("http") else self.BASE_URL + href

                        print(f"Found: {name}")
                        return CompanyInfo(
                            name=name,
                            ticker=ticker,
                            exchange=exchange or "Unknown",
                            url=company_url,
                        )

                    except Exception:
                        continue

            except Exception:
                continue

        print(f"Company not found for ticker: {ticker}")
        return None

    async def navigate_to_financials(self, company_url: str) -> bool:
        """
        Navigate to the company's financials page.

        Args:
            company_url: URL of the company's main page on Investing.com

        Returns:
            True if navigation successful
        """
        if not self._page:
            raise RuntimeError("Browser not started")

        # Navigate to company page
        await self._page.goto(company_url)
        await self._page.wait_for_load_state("networkidle")

        # Find and click on Financials tab/link
        financials_selectors = [
            'a:has-text("Financials")',
            'a[href*="financials"]',
            '[data-test="financials-tab"]',
            '.nav-item a:has-text("Financials")',
        ]

        for selector in financials_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    await element.click()
                    await self._page.wait_for_load_state("networkidle")
                    print("Navigated to Financials page.")
                    return True
            except Exception:
                continue

        # Try constructing URL directly
        if "/equities/" in company_url:
            financials_url = company_url.rstrip("/") + "-financial-summary"
            await self._page.goto(financials_url)
            await self._page.wait_for_load_state("networkidle")
            return True

        raise Exception("Could not navigate to financials page")

    async def _download_statement(self, statement_type: str) -> Optional[Path]:
        """
        Download a financial statement (Income Statement or Balance Sheet).

        Args:
            statement_type: "income-statement" or "balance-sheet"

        Returns:
            Path to downloaded file, or None if download failed
        """
        if not self._page:
            raise RuntimeError("Browser not started")

        # Click on the statement tab
        tab_text = "Income Statement" if statement_type == "income-statement" else "Balance Sheet"
        tab_selectors = [
            f'a:has-text("{tab_text}")',
            f'button:has-text("{tab_text}")',
            f'[data-test="{statement_type}-tab"]',
        ]

        clicked = False
        for selector in tab_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    await element.click()
                    await self._page.wait_for_load_state("networkidle")
                    await asyncio.sleep(1)
                    clicked = True
                    break
            except Exception:
                continue

        if not clicked:
            print(f"Could not find {tab_text} tab")
            return None

        # Switch to Quarterly view (if available)
        quarterly_selectors = [
            'button:has-text("Quarterly")',
            'a:has-text("Quarterly")',
            '[data-test="quarterly-btn"]',
            '.quarterly-btn',
        ]

        for selector in quarterly_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    await element.click()
                    await self._page.wait_for_load_state("networkidle")
                    await asyncio.sleep(1)
                    print("Switched to Quarterly view.")
                    break
            except Exception:
                continue

        # Find and click the Excel download button
        # This is typically a Pro feature
        download_selectors = [
            'button:has-text("Excel")',
            'a:has-text("Excel")',
            'button:has-text("Export")',
            '[data-test="export-excel"]',
            '.export-btn',
            'button[title*="Excel"]',
            'a[title*="Excel"]',
            # Pro export button
            '.pro-export-btn',
            '[data-test="pro-export"]',
        ]

        download_path = None

        for selector in download_selectors:
            try:
                element = await self._page.query_selector(selector)
                if element and await element.is_visible():
                    # Set up download handler
                    async with self._page.expect_download(timeout=60000) as download_info:
                        await element.click()

                    download: Download = await download_info.value
                    suggested_filename = download.suggested_filename

                    # Save to download directory
                    save_path = self.download_dir / suggested_filename
                    await download.save_as(str(save_path))

                    print(f"Downloaded: {suggested_filename}")
                    download_path = save_path
                    break

            except Exception as e:
                # If expect_download times out, the element might not trigger a download
                continue

        if not download_path:
            print(f"Could not download {tab_text}")

        return download_path

    async def download_financials(
        self,
        ticker: str,
        exchange: str = None,
    ) -> DownloadResult:
        """
        Download Income Statement and Balance Sheet for a company.

        Args:
            ticker: Stock ticker symbol
            exchange: Optional exchange code

        Returns:
            DownloadResult with paths to downloaded files
        """
        result = DownloadResult()

        # Search for company
        company_info = await self.search_company(ticker, exchange)
        if not company_info:
            result.error = f"Company not found: {ticker}"
            return result

        result.company_info = company_info

        # Navigate to financials
        try:
            await self.navigate_to_financials(company_info.url)
        except Exception as e:
            result.error = f"Failed to navigate to financials: {e}"
            return result

        # Download Income Statement
        print("\nDownloading Income Statement...")
        result.income_statement_path = await self._download_statement("income-statement")

        # Download Balance Sheet
        print("\nDownloading Balance Sheet...")
        result.balance_sheet_path = await self._download_statement("balance-sheet")

        if not result.income_statement_path and not result.balance_sheet_path:
            result.error = "Failed to download any financial statements"

        return result


async def test_scraper(email: str, password: str, ticker: str, exchange: str = None):
    """Test the scraper functionality."""
    download_dir = Path.cwd() / "test_downloads"
    download_dir.mkdir(exist_ok=True)

    async with InvestingComScraper(
        download_dir=download_dir,
        headless=False,  # Show browser for testing
    ) as scraper:
        # Login
        await scraper.login(email, password)

        # Download financials
        result = await scraper.download_financials(ticker, exchange)

        print("\n" + "=" * 50)
        print("DOWNLOAD RESULT")
        print("=" * 50)

        if result.company_info:
            print(f"Company: {result.company_info.name}")
            print(f"Ticker: {result.company_info.ticker}")
            print(f"Exchange: {result.company_info.exchange}")

        if result.income_statement_path:
            print(f"Income Statement: {result.income_statement_path}")
        if result.balance_sheet_path:
            print(f"Balance Sheet: {result.balance_sheet_path}")
        if result.error:
            print(f"Error: {result.error}")

    return result


def main():
    """CLI for testing the scraper."""
    import argparse
    from credentials import get_credentials

    parser = argparse.ArgumentParser(description="Test Investing.com scraper")
    parser.add_argument("ticker", help="Stock ticker to search for")
    parser.add_argument("--exchange", "-e", help="Exchange code (e.g., TYO, NYSE)")

    args = parser.parse_args()

    try:
        email, password = get_credentials()
    except ValueError as e:
        print(f"ERROR: {e}")
        return 1

    result = asyncio.run(test_scraper(email, password, args.ticker, args.exchange))
    return 0 if not result.error else 1


if __name__ == "__main__":
    exit(main())
