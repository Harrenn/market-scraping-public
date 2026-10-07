#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qs, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

try:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright
except ImportError:  # pragma: no cover
    PlaywrightTimeoutError = Exception
    sync_playwright = None


USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0 Safari/537.36"
)


class ScrapeError(Exception):
    pass


class ManualReviewNeeded(ScrapeError):
    pass


@dataclass
class DealerResult:
    dealer: str
    website: str
    count: int | None
    status: str
    detail: str


class BrowserSession:
    def __init__(self, headed: bool = False):
        self.headed = headed
        self._playwright = None
        self.browser = None
        self.page = None

    def __enter__(self) -> "BrowserSession":
        if sync_playwright is None:
            raise ScrapeError("Playwright is not installed in the current environment.")
        self._playwright = sync_playwright().start()
        self.browser = self._playwright.chromium.launch(headless=not self.headed)
        self.page = self.browser.new_page(viewport={"width": 1400, "height": 900})
        self.page.set_default_timeout(30000)
        self.page.set_default_navigation_timeout(30000)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.browser:
            self.browser.close()
        if self._playwright:
            self._playwright.stop()


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


def get_soup(session: requests.Session, url: str, *, timeout: int = 30) -> BeautifulSoup:
    response = session.get(url.strip(), timeout=timeout)
    response.raise_for_status()
    return BeautifulSoup(response.text, "lxml")


def get_text_lines_from_html(html: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    return [
        line.strip()
        for line in soup.get_text("\n").splitlines()
        if line.strip()
    ]


def get_text_lines(session: requests.Session, url: str, *, timeout: int = 30) -> list[str]:
    response = session.get(url.strip(), timeout=timeout)
    response.raise_for_status()
    return get_text_lines_from_html(response.text)


def extract_app_id_from_motorace(session: requests.Session) -> str | None:
    url = "https://motorace.ph/wp-json/wp/v2/pages?slug=branches"
    response = session.get(url, timeout=30)
    response.raise_for_status()
    payload = response.json()
    if not payload:
        return None
    rendered = payload[0].get("content", {}).get("rendered", "")
    match = re.search(r"elfsight-app-([a-f0-9-]{36})", rendered, re.I)
    return match.group(1) if match else None


def get_motorace_widget_boot_reason(session: requests.Session, app_id: str) -> str | None:
    response = session.get(
        "https://core.service.elfsight.com/p/boot/",
        params={"w": app_id, "page": "https://motorace.ph/branches/"},
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    return payload.get("data", {}).get("widgets", {}).get(app_id, {}).get("reason")


def scrape_cdo(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://www.cdo2cycles.com/cdo-2-cycles"
    lines = get_text_lines(session, url)
    text = "\n".join(lines)
    counts = [int(num) for num in re.findall(r"\((\d+)\)\s+Yamaha", text)]
    if not counts:
        raise ScrapeError("Could not find explicit Yamaha branch counts on the page.")
    return sum(counts)


def scrape_dueksam(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://www.dueksam.com.ph/includes/branches.php"
    headers = {"X-Requested-With": "XMLHttpRequest"}
    total = 0
    page_id = 0

    while True:
        response = session.post(
            url,
            data={"type": "fetch", "pageId": str(page_id)},
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()
        html = payload.get("data", "")
        page_rows = html.lower().count("<tr><td>")
        if page_rows == 0:
            break
        total += page_rows
        if "View More" not in payload.get("page", ""):
            break
        page_id += 1

    if total == 0:
        raise ScrapeError("No branch rows were returned from the Dueksam AJAX endpoint.")
    return total


def scrape_guanzon(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://www.guanzongroup.com.ph/branches/motorcycle-branches/"
    soup = get_soup(session, url)
    count = len(soup.select(".br-item-total p"))
    if count == 0:
        raise ScrapeError("Could not find numbered Guanzon branch entries.")
    return count


def scrape_kservico(_dealer: str, _website: str, _session: requests.Session, browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    if browser is None:
        raise ManualReviewNeeded(
            "KServico requires a headed browser because HTTP requests hit Cloudflare. "
            "Rerun with `.venv/bin/python scripts/scrape_dealer_branches.py --only \"KServico Trade, Inc.\" --headed`."
        )

    browser.page.goto("https://kservico.com.ph/store/", wait_until="domcontentloaded")
    browser.page.wait_for_timeout(5000)
    text = browser.page.locator("body").inner_text()

    if "security verification" in text.lower():
        raise ManualReviewNeeded(
            "KServico is showing a Cloudflare verification page in Playwright. "
            "Rerun with `.venv/bin/python scripts/scrape_dealer_branches.py --only \"KServico Trade, Inc.\" --headed` "
            "and solve the challenge manually if prompted."
        )

    count = text.count("Directions")
    if count == 0:
        raise ScrapeError("Rendered KServico page did not expose any store direction entries.")
    return count


def scrape_motoxpress(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://motoxpress.ph/location"
    lines = get_text_lines(session, url)
    text = "\n".join(lines)
    match = re.search(r"SHOWING\s+\d+\s+OF\s+(\d+)\s+BRANCHES", text, re.I)
    if not match:
        raise ScrapeError("MotoXpress did not expose the total branch count summary.")
    return int(match.group(1))


def scrape_motorace(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    app_id = extract_app_id_from_motorace(session)
    detail = "Motor Ace branch page is an Elfsight Store Locator widget"
    if app_id:
        detail += f" (app id: {app_id})"
        boot_reason = get_motorace_widget_boot_reason(session, app_id)
        if boot_reason:
            detail += f"; widget boot returned {boot_reason}"
    raise ManualReviewNeeded(
        detail + ". Skipping Motor Ace until an official machine-readable source is available again."
    )


def scrape_motorcentral(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://motorcentralgroup.com/branches/"
    response = session.get(url, timeout=30)
    response.raise_for_status()
    count = response.text.count("Contact Details:")
    if count == 0:
        raise ScrapeError("Could not find Motorcentral branch markers.")
    return count


def scrape_motorcyclecity(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://motorcyclecity.com.ph/store-locator/"
    soup = get_soup(session, url)
    iframe = soup.select_one("iframe[src*='google.com/maps/d/u/0/embed']")
    if not iframe:
        raise ScrapeError("MotorcycleCity did not expose the embedded Google My Maps iframe.")

    iframe_src = iframe.get("src", "")
    query = parse_qs(urlparse(iframe_src).query)
    mid = query.get("mid", [None])[0]
    if not mid:
        raise ScrapeError("Could not extract the Google My Maps map id for MotorcycleCity.")

    kml_url = f"https://www.google.com/maps/d/kml?mid={mid}&forcekml=1"
    response = session.get(kml_url, timeout=30)
    response.raise_for_status()
    count = response.text.count("<Placemark")
    if count == 0:
        raise ScrapeError("MotorcycleCity KML export did not contain any placemarks.")
    return count


def extract_motortrade_nonce(html: str) -> str:
    match = re.search(r'var ASL_REMOTE = \{"ajax_url":"[^"]+","nonce":"([^"]+)"\}', html)
    if not match:
        raise ScrapeError("Could not extract the Motortrade locator nonce.")
    return match.group(1)


def scrape_motortrade(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    page_url = "https://motortrade.com.ph/find-a-store/"
    response = session.get(page_url, timeout=30)
    response.raise_for_status()
    nonce = extract_motortrade_nonce(response.text)

    ajax_url = "https://motortrade.com.ph/wp-admin/admin-ajax.php"
    params = {
        "action": "asl_load_stores",
        "nonce": nonce,
        "load_all": "0",
        "layout": "1",
        "lat": "12.8797",
        "lng": "121.7740",
        "nw[]": ["21.5", "116.0"],
        "se[]": ["4.5", "127.5"],
    }
    ajax_response = session.get(ajax_url, params=params, timeout=60)
    ajax_response.raise_for_status()
    payload = ajax_response.json()
    count = len(payload)
    if count == 0:
        raise ScrapeError("Motortrade AJAX locator did not return any stores.")
    return count


def scrape_premiumbikes(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    endpoint = "https://premiumbikes.ph/wp-json/branch-locator/v1/branches"
    response = session.get(endpoint, timeout=30)
    response.raise_for_status()
    payload = response.json()
    count = len(payload)
    if count == 0:
        raise ScrapeError("Premiumbikes branch endpoint returned an empty dataset.")
    return count


def scrape_ropali(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    page_url = "https://ropali.ph/our-stores/"
    page_response = session.get(page_url, timeout=30)
    page_response.raise_for_status()
    match = re.search(r'https://ropali\.ph/wp-content/cache/autoptimize/js/[^"\']+\.js', page_response.text)
    if not match:
        raise ScrapeError("Could not locate the Ropali store-data script.")

    js_response = session.get(match.group(0), timeout=30)
    js_response.raise_for_status()
    count = js_response.text.count("'type':'Feature'")
    if count == 0:
        raise ScrapeError("Could not find any Ropali map features in the embedded store data.")
    return count


def scrape_unique_motor(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://uniquemotorcenter.com/contact/"
    lines = get_text_lines(session, url)
    branch_lines = [
        line
        for line in lines
        if (
            line.startswith("Unique Motor Center")
            or line.startswith("Safety Enterprise")
        )
        and re.search(r"(Main|Revzone|3s|Putik)$", line)
    ]
    count = len(branch_lines)
    if count == 0:
        raise ScrapeError("Unique Motor branch lines were not detected.")
    return count


def scrape_wheeltek(_dealer: str, _website: str, _session: requests.Session, browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    if browser is None:
        raise ManualReviewNeeded("Wheeltek requires Playwright because the branch list is rendered dynamically.")

    browser.page.goto("https://wheeltek.com.ph/branch-locator/", wait_until="domcontentloaded")
    browser.page.wait_for_timeout(5000)
    count = browser.page.locator(".sl-item").count()
    if count == 0:
        text = browser.page.locator("body").inner_text()
        count = text.count("DIRECTIONS")
    if count == 0:
        raise ScrapeError("Wheeltek did not expose any rendered branch entries.")
    return count


def scrape_motobikes(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, args: argparse.Namespace) -> int:
    url = "https://motobikesph.com/contact-us-8262"
    lines = get_text_lines(session, url)
    if "...Our Branches..." not in lines:
        raise ScrapeError("Motobikes branch section was not found.")

    start = lines.index("...Our Branches...")
    end = lines.index("Questions? We Have Answers!") if "Questions? We Have Answers!" in lines else len(lines)
    section = lines[start + 1 : end]

    active = 0
    pending = 0
    for index, line in enumerate(section):
        if not line.endswith("Branch"):
            continue
        previous_line = section[index - 1] if index > 0 else ""
        if previous_line == "Soon To Open":
            pending += 1
        else:
            active += 1

    total = active + pending if args.include_pending_motobikes else active
    if total == 0:
        raise ScrapeError("Motobikes branch entries were not detected.")
    return total


def scrape_starbike(_dealer: str, _website: str, session: requests.Session, _browser: BrowserSession | None, _args: argparse.Namespace) -> int:
    url = "https://starbikeph.com/store/"
    lines = get_text_lines(session, url)
    branch_lines = [line for line in lines if line.endswith("Branch")]
    count = len(branch_lines)
    if count == 0:
        raise ScrapeError("Starbike branch headings were not found.")
    return count


Handler = Callable[[str, str, requests.Session, BrowserSession | None, argparse.Namespace], int]


HANDLERS: dict[str, Handler] = {
    "CDO 2 Cycles Marketing Corp": scrape_cdo,
    "Du Ek Sam Inc.": scrape_dueksam,
    "Guanzon Motorcycle Group": scrape_guanzon,
    "KServico Trade, Inc.": scrape_kservico,
    "MotoXpress": scrape_motoxpress,
    "Motor Ace Philippines, Inc.": scrape_motorace,
    "Motorcentral Group of Companies": scrape_motorcentral,
    "MotorcycleCity": scrape_motorcyclecity,
    "Motortrade": scrape_motortrade,
    "Premiumbikes": scrape_premiumbikes,
    "Ropali Motorcycles": scrape_ropali,
    "Unique Motor": scrape_unique_motor,
    "Wheeltek": scrape_wheeltek,
    "Motobikes": scrape_motobikes,
    "Starbike": scrape_starbike,
}


def load_rows(workbook_path: Path):
    workbook = load_workbook(workbook_path)
    worksheet = workbook.active
    return workbook, worksheet


def process_workbook(workbook_path: Path, args: argparse.Namespace) -> list[DealerResult]:
    workbook, worksheet = load_rows(workbook_path)
    session = build_session()
    results: list[DealerResult] = []
    needs_browser = {"KServico Trade, Inc.", "Wheeltek"}
    targeted_dealers: set[str] = set()

    for row_index in range(2, worksheet.max_row + 1):
        dealer = worksheet.cell(row=row_index, column=1).value
        website = worksheet.cell(row=row_index, column=2).value
        if not dealer or not website:
            continue
        if args.only and dealer != args.only:
            continue
        targeted_dealers.add(str(dealer))

    browser: BrowserSession | None = None
    if targeted_dealers & needs_browser:
        browser = BrowserSession(headed=args.headed).__enter__()

    try:
        for row_index in range(2, worksheet.max_row + 1):
            dealer = worksheet.cell(row=row_index, column=1).value
            website = worksheet.cell(row=row_index, column=2).value
            if not dealer or not website:
                continue
            if args.only and dealer != args.only:
                continue

            handler = HANDLERS.get(str(dealer))
            if handler is None:
                results.append(
                    DealerResult(str(dealer), str(website), None, "skipped", "No handler is defined for this dealer.")
                )
                continue

            try:
                count = handler(str(dealer), str(website), session, browser, args)
                if not args.dry_run:
                    worksheet.cell(row=row_index, column=3).value = count
                results.append(DealerResult(str(dealer), str(website), count, "ok", "Count written."))
            except ManualReviewNeeded as exc:
                results.append(DealerResult(str(dealer), str(website), None, "manual", str(exc)))
            except Exception as exc:
                results.append(DealerResult(str(dealer), str(website), None, "error", str(exc)))

        if not args.dry_run:
            workbook.save(workbook_path)
    finally:
        if browser is not None:
            browser.__exit__(None, None, None)

    return results


def print_summary(results: list[DealerResult]) -> int:
    ok = 0
    manual = 0
    errors = 0

    for result in results:
        prefix = {
            "ok": "[OK]",
            "manual": "[MANUAL]",
            "error": "[ERROR]",
            "skipped": "[SKIPPED]",
        }.get(result.status, "[INFO]")

        if result.status == "ok":
            ok += 1
            print(f"{prefix} {result.dealer}: {result.count}")
        else:
            if result.status == "manual":
                manual += 1
            elif result.status == "error":
                errors += 1
            print(f"{prefix} {result.dealer}: {result.detail}")

    print()
    print(f"Successful counts: {ok}")
    print(f"Needs manual review: {manual}")
    print(f"Errors: {errors}")
    return 0 if errors == 0 else 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape motorcycle dealer branch counts and write them back to Dealers.xlsx."
    )
    parser.add_argument(
        "--workbook",
        default="working/data/Dealers.xlsx",
        help="Path to the workbook to update.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do not write any values back to the workbook.",
    )
    parser.add_argument(
        "--only",
        help="Run the scraper for a single dealer name.",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Run Playwright in headed mode for sites that require browser rendering.",
    )
    parser.add_argument(
        "--include-pending-motobikes",
        action="store_true",
        help="Include Motobikes branches marked as 'Soon To Open'.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workbook_path = Path(args.workbook).resolve()
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 2

    results = process_workbook(workbook_path, args)
    return print_summary(results)


if __name__ == "__main__":
    raise SystemExit(main())
