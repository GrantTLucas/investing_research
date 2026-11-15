#!/usr/bin/env python3
"""Screen SEC-listed ticker symbols by SIC, US business address and annual 10-K history.

Standard-library only.  Run --help for options.  Uses SEC public data; no API key.

A company qualifies if it has Form 10-K (not 10-K/A) filings in at least
five DISTINCT fiscal reporting years, filed during the trailing five-year
window.  Sorting uses the calendar year of its most recent qualifying filing,
descending, and then its ticker alphabetically.
"""

import argparse
import concurrent.futures
import csv
from datetime import date, datetime, timedelta
import json
import os
from pathlib import Path
import sys
import threading
import time
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
import zipfile

TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SUBMISSIONS_URL = "https://data.sec.gov/submissions/"
STATES = set("AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC".split())


class RateLimiter:
    def __init__(self, requests_per_second):
        self.gap = 1.0 / requests_per_second
        self.next_request = 0.0
        self.lock = threading.Lock()

    def wait(self):
        with self.lock:
            now = time.monotonic()
            if now < self.next_request:
                time.sleep(self.next_request - now)
            self.next_request = time.monotonic() + self.gap


class SecClient:
    def __init__(self, email, cache_dir, refresh, rate):
        self.agent = f"InvestingResearch SIC 10K Screener ({email})"
        self.cache_dir = cache_dir
        self.refresh = refresh
        self.limiter = RateLimiter(rate)

    def get_json(self, url, cache_file):
        # Cached SEC metadata/filing histories are refreshed after 24 hours.
        if not self.refresh and cache_file.is_file():
            if time.time() - cache_file.stat().st_mtime < 86400:
                try:
                    with cache_file.open("r", encoding="utf-8") as f:
                        return json.load(f)
                except (OSError, ValueError):
                    pass

        error = None
        for attempt in range(5):
            self.limiter.wait()
            try:
                req = Request(url, headers={
                    "User-Agent": self.agent,
                    "Accept": "application/json",
                    "Accept-Encoding": "identity",
                })
                with urlopen(req, timeout=45) as response:
                    payload = response.read()
                obj = json.loads(payload)
                cache_file.parent.mkdir(parents=True, exist_ok=True)
                temp = cache_file.with_name(cache_file.name + f".{threading.get_ident()}.tmp")
                try:
                    temp.write_bytes(payload)
                    os.replace(temp, cache_file)
                finally:
                    temp.unlink(missing_ok=True)
                return obj
            except (HTTPError, URLError, TimeoutError, OSError, ValueError) as exc:
                error = exc
                if isinstance(exc, HTTPError) and exc.code not in (403, 408, 429, 500, 502, 503, 504):
                    break
                if attempt < 4:
                    time.sleep(min(30, 2 ** (attempt + 1)))
        raise RuntimeError(f"Failed to fetch {url}: {error}")

    def fetch_tickers(self):
        return self.get_json(TICKERS_URL, self.cache_dir / "company_tickers.json")

    def fetch_submission(self, filename):
        return self.get_json(SUBMISSIONS_URL + filename,
                             self.cache_dir / "submissions" / filename)


def five_years_ago(day):
    try:
        return day.replace(year=day.year - 5)
    except ValueError:  # February 29 in a non-leap year
        return day.replace(year=day.year - 5, day=28)


def ticker_universe(ticker_data):
    """Map CIK -> current ticker symbols listed in SEC company_tickers.json."""
    if isinstance(ticker_data, dict):
        records = ticker_data.values()
    elif isinstance(ticker_data, list):
        records = ticker_data
    else:
        raise ValueError("Unexpected SEC company_tickers.json format")
    ciks = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        try:
            cik = f"{int(record['cik_str']):010d}"
            ticker = str(record['ticker']).strip().upper()
        except (KeyError, TypeError, ValueError):
            continue
        if ticker:
            ciks.setdefault(cik, set()).add(ticker)
    return ciks


def filings_in_window(block, beginning, ending):
    if not isinstance(block, dict):
        return
    forms = block.get("form", [])
    dates = block.get("filingDate", [])
    reports = block.get("reportDate", [])
    for form, filed, reported in zip(forms, dates, reports):
        if form != "10-K" or not (beginning <= filed <= ending):
            continue
        try:
            filing_day = date.fromisoformat(filed)
            fiscal_year = date.fromisoformat(reported).year
        except (ValueError, TypeError):
            continue
        yield fiscal_year, filing_day


def qualifies(data, tickers, get_continuation, beginning, ending, sic_filter):
    sic = str(data.get("sic") or "").zfill(4)
    if not sic.isdecimal() or sic == "0000":
        return []
    if sic_filter and sic not in sic_filter:
        return []
    address = data.get("addresses", {}).get("business") or {}
    state = str(address.get("stateOrCountry") or "").upper()
    if state not in STATES:
        return []

    filings = data.get("filings") or {}
    annual_year_to_filing = {}

    def add_filings(block):
        for fiscal_year, filing_date in filings_in_window(block, beginning, ending):
            current = annual_year_to_filing.get(fiscal_year)
            if current is None or filing_date > current:
                annual_year_to_filing[fiscal_year] = filing_date

    add_filings(filings.get("recent", {}))
    for older in filings.get("files", []):
        # Only fetch history chunks whose advertised date range intersects our window.
        if older.get("filingTo", "") < beginning or older.get("filingFrom", "") > ending:
            continue
        name = older.get("name", "")
        if not name.startswith("CIK") or not name.endswith(".json") or "/" in name:
            continue
        add_filings(get_continuation(name))

    if len(annual_year_to_filing) < 5:
        return []

    # Show actual filing years (not fiscal report years), one for each distinct
    # annual reporting year.  A calendar year can appear twice if necessary.
    ordered_dates = sorted(annual_year_to_filing.values(), reverse=True)
    years = ", ".join(str(d.year) for d in ordered_dates)
    newest_filing_year = ordered_dates[0].year
    return [(sic, ticker, newest_filing_year, years) for ticker in sorted(tickers)]


def load_local_json(filename):
    with open(filename, "r", encoding="utf-8") as f:
        return json.load(f)


def output_results(rows, selected, output_dir, quiet):
    output_dir.mkdir(parents=True, exist_ok=True)
    groups = {}
    for row in rows:
        groups.setdefault(row[0], []).append(row)

    # Tie-break by symbol, not exact month/day, as requested.
    sort_key = lambda x: (-x[2], x[1])
    for entries in groups.values():
        entries.sort(key=sort_key)

    master_path = output_dir / "all_sic_tickers.csv"
    with master_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["SIC", "Ticker", "10-K Filing Years"])
        for row in sorted(rows, key=lambda x: (-x[2], x[1], x[0])):
            writer.writerow((row[0], row[1], row[3]))

    for sic in sorted(groups):
        with (output_dir / f"sic_{sic}.csv").open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Ticker", "10-K Filing Years"])
            for _, ticker, _, years in groups[sic]:
                writer.writerow((ticker, years))

    if not quiet:
        for sic in sorted(selected or groups):
            entries = groups.get(sic, [])
            print(f"\nSIC {sic}  ({len(entries)} tickers)")
            print(f"{'Ticker':<10}  10-K Filing Years")
            print(f"{'-' * 10}  {'-' * 28}")
            for _, ticker, _, years in entries:
                print(f"{ticker:<10}  {years}")

    print(f"\nQualified tickers: {len(rows):,} | SIC groups: {len(groups):,}", file=sys.stderr)
    print(f"CSV output: {output_dir.resolve()}", file=sys.stderr)
    return master_path


def main():
    parser = argparse.ArgumentParser(
        description="Filter SEC tickers by US business address, SIC, and 5 annual 10-K reporting years."
    )
    parser.add_argument("--email", default=os.environ.get("SEC_EMAIL", ""),
                        help="Real SEC request contact email (or set SEC_EMAIL)")
    parser.add_argument("--sic", action="append", default=[],
                        help="Limit displayed/exported SICs (repeat or comma-separate); default: all")
    parser.add_argument("--bulk-zip", type=Path, help="Use an EXISTING SEC submissions.zip rather than individual API calls")
    parser.add_argument("--tickers-json", type=Path, help="Use an existing SEC company_tickers.json")
    parser.add_argument("--cache-dir", type=Path, default=Path(".edgar_cache"))
    parser.add_argument("--output-dir", type=Path, default=Path("sic_tables"))
    parser.add_argument("--as-of", type=date.fromisoformat, default=date.today(), metavar="YYYY-MM-DD")
    parser.add_argument("--refresh", action="store_true", help="Ignore API caches and request fresh SEC data")
    parser.add_argument("--workers", type=int, default=6, help="Concurrent API workers (default: 6)")
    parser.add_argument("--rate", type=float, default=6.0, help="Maximum SEC HTTP requests per second (default: 6)")
    parser.add_argument("--quiet", action="store_true", help="Only save CSV files, do not print tables")
    args = parser.parse_args()

    if not args.email or "@" not in args.email:
        parser.error("Supply a real contact email using --email or SEC_EMAIL")
    if not 1 <= args.workers <= 10 or not 0 < args.rate <= 9:
        parser.error("Require --workers between 1 and 10 and --rate >0 and <=9")
    if args.as_of > date.today():
        parser.error("--as-of cannot be a future date")

    chosen_sics = set()
    for chunk in args.sic:
        for value in chunk.split(","):
            value = value.strip()
            if not value.isdigit() or not 1 <= len(value) <= 4:
                parser.error(f"Invalid SIC code: {value!r}")
            chosen_sics.add(value.zfill(4))

    beginning = five_years_ago(args.as_of).isoformat()
    ending = args.as_of.isoformat()
    print(f"Filing window: {beginning} to {ending}, inclusive", file=sys.stderr)

    sec = SecClient(args.email, args.cache_dir, args.refresh, args.rate)
    ticker_data = load_local_json(args.tickers_json) if args.tickers_json else sec.fetch_tickers()
    cik_tickers = ticker_universe(ticker_data)
    print(f"Current SEC ticker universe: {len(cik_tickers):,} CIKs", file=sys.stderr)
    rows = []
    failures = []

    if args.bulk_zip:
        if not args.bulk_zip.is_file():
            parser.error(f"Bulk ZIP does not exist: {args.bulk_zip}")
        with zipfile.ZipFile(args.bulk_zip) as zf:
            def get_zip_json(name):
                with zf.open(name) as f:
                    return json.load(f)

            for i, (cik, tickers) in enumerate(cik_tickers.items(), start=1):
                try:
                    data = get_zip_json(f"CIK{cik}.json")
                    rows.extend(qualifies(data, tickers, get_zip_json,
                                          beginning, ending, chosen_sics))
                except (KeyError, OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
                    failures.append((cik, str(exc)))
                if i % 500 == 0:
                    print(f"Checked {i:,}/{len(cik_tickers):,} CIKs", file=sys.stderr)
    else:
        def evaluate(item):
            cik, tickers = item
            data = sec.fetch_submission(f"CIK{cik}.json")
            return qualifies(data, tickers, sec.fetch_submission,
                             beginning, ending, chosen_sics)

        with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
            jobs = {pool.submit(evaluate, item): item[0]
                    for item in cik_tickers.items()}
            for i, future in enumerate(concurrent.futures.as_completed(jobs), start=1):
                try:
                    rows.extend(future.result())
                except (RuntimeError, KeyError, OSError, ValueError) as exc:
                    failures.append((jobs[future], str(exc)))
                if i % 250 == 0:
                    print(f"Checked {i:,}/{len(jobs):,} CIKs", file=sys.stderr)

    output_results(rows, chosen_sics, args.output_dir, args.quiet)
    if failures:
        error_file = args.output_dir / "failed_ciks.csv"
        with error_file.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(("CIK", "Error"))
            writer.writerows(failures)
        print(f"WARNING: {len(failures):,} CIKs failed. Results are INCOMPLETE. See {error_file}",
              file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
