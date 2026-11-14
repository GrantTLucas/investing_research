#!/usr/bin/env python3
"""Update three-sheet GTL-style XLSX financial spreads from SEC EDGAR XBRL.

Zero-byte .xlsx placeholders made with `touch` are initialized automatically.
Filenames use TICKER_COMPANY TITLE_GTL.xlsx; punctuation and hyphens are allowed.

Python 3.10+; standard library only. Run --help for options.
Read the companion README before using on investment workbooks.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import difflib
import gzip
import io
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zlib
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, BadZipFile, ZIP_DEFLATED

LOG = logging.getLogger("edgar_updater")
X = "http://www.w3.org/1999/xlink"
M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
API = "https://data.sec.gov"
SITE = "https://www.sec.gov"
SEC_FORMS = {"10-Q", "10-K"}
CELL_RE = re.compile(r'<c\b[^>]*?\br="([A-Z]+)(\d+)"[^>]*?(?:\s*/>|>.*?</c>)', re.S)
FORMULA_REF = re.compile(r'(?<![A-Za-z0-9_])(?P<colabs>\$?)(?P<col>[A-Z]{1,3})(?P<rowabs>\$?)(?P<row>[1-9][0-9]{0,6})(?![A-Za-z0-9_])')
CHANGE_HEADER = {'quarter': 'QoQ %', 'annual': 'YoY %'}


def change_header(period: str) -> str:
    return CHANGE_HEADER['annual' if period.endswith('FY') else 'quarter']


def prior_period(period: str) -> str:
    """Adjacent quarter for quarterly labels; previous full year for FY labels."""
    year = int(period[:4])
    if period.endswith('FY'):
        return f'{year-1}FY'
    q = int(period[-1])
    return f'{year}Q{q-1}' if q > 1 else f'{year-1}Q4'


def _formula_translate(formula: str, from_col: int, to_col: int, col_map) -> str:
    """Translate shared formula's unanchored relative columns, then map moved columns."""
    def replace(match):
        original_col = col_number(match.group('col'))
        ref_col = original_col if match.group('colabs') else original_col + (to_col - from_col)
        if ref_col < 1: raise ValueError(f'Invalid translated Excel formula reference: {formula}')
        return f"{match.group('colabs')}{col_name(col_map(ref_col))}{match.group('rowabs')}{match.group('row')}"
    return FORMULA_REF.sub(replace, formula)


def _remap_sheet_xml(source: str, col_map, add_pair_border=False) -> str:
    """Move physical worksheet cells and expand shared formulas to independent formulas.

    Uses an OOXML-preserving text transformation rather than a full XML rewrite,
    retaining worksheet extension namespaces and existing drawing relationships.
    """
    root = ET.fromstring(source)
    masters = {}
    for row in root.findall('.//{%s}sheetData/{%s}row' % (M, M)):
        for cell in row.findall('{%s}c' % M):
            fm = cell.find('{%s}f' % M)
            if fm is None or fm.get('t') != 'shared' or fm.text is None: continue
            address = re.fullmatch(r'([A-Z]+)([0-9]+)', cell.get('r', ''))
            if address:
                masters[fm.get('si')] = (col_number(address[1]), int(address[2]), fm.text)

    def transform_cell(match):
        old_col, old_row = col_number(match.group(1)), int(match.group(2))
        new_col = col_map(old_col)
        text = match.group(0)
        start_close = text.find('>')
        first = text[:start_close + 1]
        new_address = f'{col_name(new_col)}{old_row}'
        first = re.sub(r'(\br=")[A-Z]+[0-9]+(")', lambda m: m[1] + new_address + m[2], first, count=1)
        text = first + text[start_close + 1:]
        def formula_repl(fm):
            attrs, formula = fm.group(1), fm.group(2)
            if 't="shared"' in attrs:
                sid = re.search(r'\bsi="([^"]+)"', attrs)
                if not sid or sid[1] not in masters:
                    raise ValueError('Unresolved shared Excel formula in worksheet')
                master_col, master_row, master_formula = masters[sid[1]]
                # Input files have shared formulas along the same statement row.
                if master_row != old_row:
                    raise ValueError('Vertical shared formulas require manual conversion')
                return f'<f>{escape(_formula_translate(master_formula, master_col, old_col, col_map))}</f>'
            if formula is None: return fm.group(0)
            return f'<f{attrs}>{escape(_formula_translate(formula, old_col, old_col, col_map))}</f>'
        text = re.sub(r'<f([^>]*?)(?:>(.*?)</f>|/>)', formula_repl, text, flags=re.S)
        return text

    output = CELL_RE.sub(transform_cell, source)
    # Expand explicitly stored worksheet dimensions, merged headings, and column widths.
    def remap_range(match, extend_end=False):
        a, b = match.group(1), match.group(2)
        def ref(value, ending):
            mm = re.fullmatch(r'([A-Z]+)([0-9]+)', value)
            if not mm: return value
            old = col_number(mm[1]); new = col_map(old)
            if ending and extend_end and old >= 3 and new != old:
                new += 1
            return f'{col_name(new)}{mm[2]}'
        return f'{ref(a, False)}:{ref(b, True)}'
    output = re.sub(r'(?<=<dimension ref=")[A-Z]+[0-9]+:[A-Z]+[0-9]+',
                    lambda m: remap_range(re.fullmatch(r'([A-Z]+[0-9]+):([A-Z]+[0-9]+)', m.group(0)), True), output)
    output = re.sub(r'(?<=<mergeCell ref=")[A-Z]+[0-9]+:[A-Z]+[0-9]+',
                    lambda m: remap_range(re.fullmatch(r'([A-Z]+[0-9]+):([A-Z]+[0-9]+)', m.group(0)), True), output)
    output = re.sub(r'(<row\b[^>]*\bspans=")[0-9]+:[0-9]+(")',
                    lambda m: re.sub(r'[0-9]+:[0-9]+', lambda v: f'{v[0].split(":")[0]}:{col_map(int(v[0].split(":")[1]))}', m[0]), output)
    ET.fromstring(output)
    return output


def _ensure_percent_style(styles: bytes) -> tuple[bytes, int]:
    """Return styles.xml with an Excel built-in 0.00% style, plus its style index."""
    xml = styles.decode('utf-8')
    root = ET.fromstring(xml)
    xf_parent = root.find('{%s}cellXfs' % M)
    if xf_parent is None: raise ValueError('XLSX has no cellXfs styles collection')
    for i, xf in enumerate(xf_parent):
        if xf.get('numFmtId') == '10' and xf.get('applyNumberFormat') == '1':
            return styles, i
    index = len(xf_parent)
    # Clone a normal numeric style if possible, but change the number format.
    attributes = (dict(xf_parent[0].attrib) if len(xf_parent) else
                  {'fontId': '0', 'fillId': '0', 'borderId': '0', 'xfId': '0'})
    attributes.update({'numFmtId': '10', 'applyNumberFormat': '1'})
    from xml.sax.saxutils import quoteattr
    xf_xml = '<xf ' + ' '.join(f'{k}={quoteattr(v)}' for k,v in attributes.items()) + '/>'
    new_count = index + 1
    xml, n = re.subn(r'(<cellXfs\b[^>]*\bcount=")[0-9]+(")',
                     lambda m: m[1] + str(new_count) + m[2], xml, count=1)
    if n != 1 or '</cellXfs>' not in xml: raise ValueError('Cannot append percent format to XLSX styles')
    xml = xml.replace('</cellXfs>', xf_xml + '</cellXfs>', 1)
    ET.fromstring(xml)
    return xml.encode('utf-8'), index

ROW_RE = re.compile(r'<row\b[^>]*\br="(\d+)"[^>]*(?:/>|>.*?</row>)', re.S)
PERIOD_RE = re.compile(r"^(20\d\d)(Q[1-4]|FY)$")
SKIP_LABELS = {
    "assets", "liabilities", "equity", "expenses", "revenues", "other",
    "current assets", "current liabilities", "stockholders equity",
    "operating expenses", "other income and taxes", "noncurrent assets",
    "stockholders' equity", "reconciliation checks all should equal zero",
}
CHECK_TOKENS = ("check", "reconciliation", "difference", "should equal zero")

# Explicit, conservative synonyms for standard US-GAAP concepts.
# Custom company-specific items are primarily found in filing instance + label linkbase.
ALIASES = {
    "cash and cash equivalents": ["us-gaap:CashAndCashEquivalentsAtCarryingValue"],
    "restricted cash": ["us-gaap:RestrictedCashAndCashEquivalents", "us-gaap:RestrictedCash"],
    "goodwill": ["us-gaap:Goodwill"],
    "total assets": ["us-gaap:Assets"],
    "total liabilities": ["us-gaap:Liabilities"],
    "total current assets": ["us-gaap:AssetsCurrent"],
    "total current liabilities": ["us-gaap:LiabilitiesCurrent"],
    "total stockholders equity": ["us-gaap:StockholdersEquity"],
    "total equity": ["us-gaap:StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
    "total liabilities and stockholders equity": ["us-gaap:LiabilitiesAndStockholdersEquity"],
    "total liabilities mezzanine equity and equity": ["us-gaap:LiabilitiesAndStockholdersEquity"],
    "net income loss": ["us-gaap:ProfitLoss", "us-gaap:NetIncomeLoss"],
    "net income loss consolidated": ["us-gaap:ProfitLoss", "us-gaap:NetIncomeLoss"],
    "income loss before income taxes": ["us-gaap:IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
    "income loss before taxes": ["us-gaap:IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
    "basic earnings per share": ["us-gaap:EarningsPerShareBasic"],
    "diluted earnings per share": ["us-gaap:EarningsPerShareDiluted"],
    "basic weighted average shares millions": ["us-gaap:WeightedAverageNumberOfSharesOutstandingBasic"],
    "diluted weighted average shares millions": ["us-gaap:WeightedAverageNumberOfDilutedSharesOutstanding"],
    "additional paid in capital": ["us-gaap:AdditionalPaidInCapital"],
    "treasury stock at cost": ["us-gaap:TreasuryStockValue"],
    "long term borrowings": ["us-gaap:LongTermDebtNoncurrent", "us-gaap:LongTermDebt"],
    "accounts payable and accrued expenses": ["us-gaap:AccountsPayableAndAccruedLiabilitiesCurrent"],
    "net cash provided by used in operating activities": ["us-gaap:NetCashProvidedByUsedInOperatingActivities"],
    "net cash provided by used in investing activities": ["us-gaap:NetCashProvidedByUsedInInvestingActivities"],
    "net cash provided by used in financing activities": ["us-gaap:NetCashProvidedByUsedInFinancingActivities"],
}


def norm(value: str) -> str:
    value = re.sub(r"\([^)]*\b(?:when separately presented|pre-sab 122)\b[^)]*\)", "", str(value), flags=re.I)
    value = value.lower().replace("&", " and ").replace("–", " ").replace("—", " ")
    value = value.replace("non-controlling", "noncontrolling").replace("non-current", "noncurrent")
    value = re.sub(r"\b(?:usd|millions|thousands|shares|share)\b", lambda m: m.group(0), value)
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def tag_prefix(ns: str, declared: dict[str, str]) -> str:
    for prefix, uri in declared.items():
        if uri == ns:
            return prefix
    if "fasb.org/us-gaap" in ns: return "us-gaap"
    if "xbrl.sec.gov/dei" in ns: return "dei"
    if "xbrl.sec.gov/srt" in ns: return "srt"
    return "ext"


def tag_key(tag: str, prefixes: dict[str, str]) -> str:
    if tag.startswith("{"):
        namespace, local = tag[1:].split("}", 1)
        return f"{tag_prefix(namespace, prefixes)}:{local}"
    return tag


def xml_prefixes(xml_bytes: bytes) -> dict[str, str]:
    prefixes = {}
    for _, item in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        prefix, ns = item
        if prefix and prefix not in prefixes:
            prefixes[prefix] = ns
    return prefixes


def parse_filename(path: Path):
    """Parse TICKER_COMPANY TITLE_OWNER.xlsx without changing company punctuation.

    Only the first underscore (ticker separator) and the last underscore
    (owner separator) are structural. Additional underscores and hyphens
    inside the company name are left intact.
    """
    if path.suffix.lower() != ".xlsx":
        return None
    ticker_part, first_separator, remainder = path.stem.partition("_")
    if not first_separator:
        return None
    company, last_separator, owner = remainder.rpartition("_")
    if not last_separator or not company.strip() or not owner.strip():
        return None
    ticker = ticker_part.strip().upper().replace(".", "-")
    if not re.fullmatch(r"[A-Z][A-Z0-9-]{0,9}", ticker) or ticker == "TICKER":
        return None
    owner = owner.strip().upper()
    if owner == "OWNER":
        return None
    return ticker, company, owner, path.suffix[1:].lower()


def col_number(name: str) -> int:
    result = 0
    for char in name:
        result = result * 26 + ord(char) - ord("A") + 1
    return result


def col_name(number: int) -> str:
    s = ""
    while number:
        number, rem = divmod(number - 1, 26)
        s = chr(rem + 65) + s
    return s


class SECClient:
    def __init__(self, agent: str, cache: Path, interval: float = 0.5, retries: int = 4):
        self.agent = agent
        self.cache = cache
        self.interval = max(interval, 0.2)
        self.retries = retries
        self.last_request = 0.0
        cache.mkdir(parents=True, exist_ok=True)

    def fetch(self, url: str, *, ttl: int | None = None) -> bytes:
        from hashlib import sha256
        key = sha256(url.encode()).hexdigest()
        cache_file = self.cache / (key + ".gz")
        if cache_file.exists() and (ttl is None or time.time() - cache_file.stat().st_mtime < ttl):
            with gzip.open(cache_file, "rb") as handle: return handle.read()
        for attempt in range(self.retries):
            elapsed = time.monotonic() - self.last_request
            if elapsed < self.interval: time.sleep(self.interval - elapsed)
            self.last_request = time.monotonic()
            request = urllib.request.Request(url, headers={
                "User-Agent": self.agent,
                "Accept-Encoding": "gzip, deflate",
                "Accept": "application/json, application/xml, text/xml, */*",
            })
            try:
                with urllib.request.urlopen(request, timeout=40) as response:
                    raw = response.read()
                    encoding = (response.headers.get("Content-Encoding") or "").lower()
                    if encoding == "gzip": raw = gzip.decompress(raw)
                    elif encoding == "deflate": raw = zlib.decompress(raw)
                with gzip.open(cache_file, "wb") as handle: handle.write(raw)
                return raw
            except urllib.error.HTTPError as exc:
                if exc.code in (403, 429, 500, 502, 503, 504) and attempt < self.retries-1:
                    pause = min(60, 2 ** (attempt+1)); LOG.warning("SEC HTTP %s; retry in %ss: %s", exc.code, pause, url)
                    time.sleep(pause)
                    continue
                raise
            except (urllib.error.URLError, TimeoutError) as exc:
                if attempt < self.retries-1:
                    pause = 2 ** (attempt+1); LOG.warning("SEC network problem: %s; retry in %ss", exc, pause)
                    time.sleep(pause)
                    continue
                raise
        raise RuntimeError(f"Unable to retrieve {url}")

    def get_json(self, url: str, *, ttl: int | None = None) -> dict:
        return json.loads(self.fetch(url, ttl=ttl))

    def cik_for(self, ticker: str) -> int:
        data = self.get_json(f"{SITE}/files/company_tickers.json", ttl=86400)
        for company in data.values():
            if company["ticker"].upper().replace(".", "-") == ticker:
                return int(company["cik_str"])
        raise ValueError(f"Ticker {ticker} not found in SEC company_tickers.json (verify ticker/CIK)")

    def submissions(self, cik: int) -> dict:
        return self.get_json(f"{API}/submissions/CIK{cik:010d}.json", ttl=0)

    def companyfacts(self, cik: int) -> dict:
        # Re-fetch once per polling run: facts may appear after submissions listing.
        return self.get_json(f"{API}/api/xbrl/companyfacts/CIK{cik:010d}.json", ttl=0)


@dataclass(frozen=True)
class Filing:
    accn: str
    form: str
    report_date: str
    filed: str
    document: str
    cik: int

    @property
    def base_url(self):
        return f"{SITE}/Archives/edgar/data/{self.cik}/{self.accn.replace('-', '')}"

    @property
    def source_url(self):
        return f"{self.base_url}/{self.document}" if self.document else f"{self.base_url}/"

    @property
    def period(self):
        y, q = quarter(self.report_date)
        return f"{y}FY" if self.form == "10-K" else f"{y}Q{q}"


def quarter(date: str) -> tuple[int, int]:
    d = dt.date.fromisoformat(date)
    return d.year, (d.month - 1) // 3 + 1


def recent_filings(submissions: dict, cik: int, limit: int) -> list[Filing]:
    recent = submissions.get("filings", {}).get("recent", {})
    keys = ["form", "accessionNumber", "reportDate", "filingDate", "primaryDocument"]
    if any(key not in recent for key in keys): raise ValueError("Missing submissions recent filing metadata")
    found = []
    for values in zip(*(recent[key] for key in keys)):
        form, accn, report, filed, document = values
        if form not in SEC_FORMS or not report: continue
        try: year, q = quarter(report)
        except ValueError: continue
        # Template columns are calendar quarters ending in March, June, Sept, Dec.
        report_end = dt.date.fromisoformat(report)
        if report_end.month not in {3, 6, 9, 12}:
            LOG.warning("Skipping non-calendar quarter-end filing %s (%s); custom fiscal-year mapping required", accn, report)
            continue
        found.append(Filing(accn, form, report, filed, document, cik))
        if len(found) >= limit: break
    return found


@dataclass(frozen=True)
class Fact:
    key: str
    start: str | None
    end: str
    unit: str
    value: Decimal
    accn: str
    filed: str
    source: str


def parse_companyfacts(data: dict, accns: set[str]) -> tuple[dict[str, list[Fact]], dict[str, list[str]]]:
    facts = defaultdict(list)
    labels = defaultdict(list)
    for taxonomy, concepts in data.get("facts", {}).items():
        if taxonomy not in ("us-gaap", "dei", "srt"): continue
        for tag, item in concepts.items():
            key = f"{taxonomy}:{tag}"
            for label in (item.get("label"), item.get("description")):
                if label and len(label) < 180: labels[key].append(label)
            for unit, entries in item.get("units", {}).items():
                if unit not in ("USD", "shares", "USD/shares"): continue
                for entry in entries:
                    if entry.get("accn") not in accns or not entry.get("end"): continue
                    if entry.get("form") not in SEC_FORMS: continue
                    try: value = Decimal(str(entry["val"]))
                    except (KeyError, InvalidOperation): continue
                    facts[key].append(Fact(key, entry.get("start"), entry["end"], unit,
                                           value, entry["accn"], entry.get("filed", ""), "companyfacts"))
    return facts, labels


def parse_instance(data: bytes, filing: Filing) -> dict[str, list[Fact]]:
    prefixes = xml_prefixes(data)
    root = ET.fromstring(data)
    contexts = {}
    units = {}
    for child in root:
        if child.tag.endswith("}context"):
            cid = child.get("id")
            if child.find(".//{*}segment") is not None or child.find(".//{*}scenario") is not None:
                continue  # dimensional disclosures cannot safely become consolidated statement figures
            instant = child.find(".//{*}period/{*}instant")
            begin = child.find(".//{*}period/{*}startDate")
            end = child.find(".//{*}period/{*}endDate")
            if instant is not None: contexts[cid] = (None, (instant.text or "").strip())
            elif begin is not None and end is not None:
                contexts[cid] = ((begin.text or "").strip(), (end.text or "").strip())
        elif child.tag.endswith("}unit"):
            measures = [x.text or "" for x in child.findall(".//{*}measure")]
            if len(measures) == 1:
                label = measures[0].split(":")[-1]
            elif "USD" in " ".join(measures) and "shares" in " ".join(measures):
                label = "USD/shares"
            else:
                label = "other"
            units[child.get("id")] = label
    facts = defaultdict(list)
    ignored = {"context", "unit", "schemaRef", "roleRef", "arcroleRef", "footnoteLink"}
    for elem in root:
        if elem.tag.split("}")[-1] in ignored: continue
        context = contexts.get(elem.get("contextRef"))
        if not context or elem.get("nil") == "true": continue
        unit = units.get(elem.get("unitRef"), "other")
        if unit not in ("USD", "shares", "USD/shares"): continue
        if not context[1] or not elem.text: continue
        try: value = Decimal(elem.text.strip().replace(",", ""))
        except InvalidOperation: continue
        key = tag_key(elem.tag, prefixes)
        facts[key].append(Fact(key, context[0], context[1], unit, value,
                               filing.accn, filing.filed, "filing_instance"))
    return facts


def parse_labels(data: bytes) -> dict[str, list[str]]:
    root = ET.fromstring(data)
    labels = defaultdict(list)
    for link in root.findall(".//{*}labelLink"):
        locations = {}
        resources = {}
        for item in link:
            kind = item.tag.split("}")[-1]
            if kind == "loc":
                href = item.get(f"{{{X}}}href", "")
                fragment = href.rsplit("#", 1)[-1]
                if "_" in fragment:
                    ns, name = fragment.split("_", 1)
                    locations[item.get(f"{{{X}}}label")] = f"{ns}:{name}"
            elif kind == "label":
                role = item.get(f"{{{X}}}role", "").split("/")[-1]
                if role in ("label", "terseLabel", "verboseLabel", "totalLabel", "periodStartLabel", "periodEndLabel"):
                    text = " ".join("".join(item.itertext()).split())
                    if text and len(text) < 220:
                        resources[item.get(f"{{{X}}}label")] = (role, text)
        for arc in link.findall("{*}labelArc"):
            key = locations.get(arc.get(f"{{{X}}}from"))
            payload = resources.get(arc.get(f"{{{X}}}to"))
            if key and payload:
                role, label = payload
                if label not in labels[key]: labels[key].append(label)
    return labels


def parse_presentation(data: bytes) -> dict[str, set[str]]:
    root = ET.fromstring(data)
    result = {"balance": set(), "income": set(), "cashflow": set()}
    for presentation in root.findall(".//{*}presentationLink"):
        role = presentation.get(f"{{{X}}}role", "").rsplit("/", 1)[-1]
        token = re.sub(r"[^a-z]", "", role.lower())
        if any(word in token for word in ("parenthetical", "detail", "disclosure", "note", "schedule", "supplemental")):
            continue
        if any(word in token for word in ("balancesheet", "financialposition")):
            typ = "balance"
        elif any(word in token for word in ("cashflow", "cashflows")):
            typ = "cashflow"
        elif any(word in token for word in ("statementofoperation", "statementofincome", "income statement", "statementofearnings", "statementofloss", "statementofcomprehensiveincome", "statementsofoperations")):
            typ = "income"
        else:
            continue
        for loc in presentation.findall("{*}loc"):
            fragment = loc.get(f"{{{X}}}href", "").rsplit("#", 1)[-1]
            if "_" not in fragment: continue
            ns, name = fragment.split("_", 1)
            result[typ].add(f"{ns}:{name}")
    return result


def extension_filenames(index_json: dict, document: str) -> dict[str, str]:
    names = [obj["name"] for obj in index_json.get("directory", {}).get("item", []) if "name" in obj]
    name_by_lower = {name.lower(): name for name in names}
    primary = document.rsplit(".", 1)[0] + "_htm.xml" if document else ""
    result = {}
    for kind, suffix in (("instance", "_htm.xml"), ("labels", "_lab.xml"), ("presentation", "_pre.xml")):
        matches = [name for name in names if name.lower().endswith(suffix)]
        if kind == "instance" and primary.lower() in name_by_lower:
            result[kind] = name_by_lower[primary.lower()]
        elif matches:
            result[kind] = matches[0]
    return result


def gather_filing(client: SECClient, filing: Filing):
    """Return instance facts, labels and statement presentation role concepts."""
    try:
        index = client.get_json(f"{filing.base_url}/index.json", ttl=3600)
        names = extension_filenames(index, filing.document)
        facts = defaultdict(list); labels = defaultdict(list)
        roles = {"balance": set(), "income": set(), "cashflow": set()}
        if "instance" in names:
            facts.update(parse_instance(client.fetch(f"{filing.base_url}/{names['instance']}"), filing))
        if "labels" in names:
            labels.update(parse_labels(client.fetch(f"{filing.base_url}/{names['labels']}")))
        if "presentation" in names:
            roles = parse_presentation(client.fetch(f"{filing.base_url}/{names['presentation']}"))
        return facts, labels, roles
    except Exception as exc:
        LOG.warning("Could not download complete filing XBRL for %s: %s. Standard GAAP fallback remains available.", filing.accn, exc)
        return {}, {}, {"balance": set(), "income": set(), "cashflow": set()}


def merge_dict_lists(destination, source):
    for key, values in source.items():
        for value in values:
            if value not in destination[key]: destination[key].append(value)


class XMLSheet:
    """Patch only XML cells/rows, preserving the original XLSX formatting and metadata."""
    def __init__(self, xml_text: str, shared: list[str], name: str):
        self.name = name
        self.original = xml_text
        self.shared = shared
        self.values: dict[tuple[int, int], object] = {}
        self.existing: dict[tuple[int, int], str] = {}
        self.styles: dict[tuple[int, int], str] = {}
        self.pending: dict[tuple[int, int], tuple[str, object]] = {}
        self.percent_style = None
        self.report_styles = {}
        self.row_roles = {}
        self.layout_changed = False
        self._read()

    def _read(self):
        root = ET.fromstring(self.original)
        for row in root.findall(".//{%s}sheetData/{%s}row" % (M, M)):
            for cell in row.findall("{%s}c" % M):
                ref = cell.get("r")
                m = re.match(r"([A-Z]+)(\d+)$", ref or "")
                if not m: continue
                pos = (int(m.group(2)), col_number(m.group(1)))
                raw_val = cell.find("{%s}v" % M)
                inline = cell.find("{%s}is" % M)
                formula = cell.find("{%s}f" % M)
                content = raw_val.text if raw_val is not None else None
                if cell.get("t") == "s" and content is not None:
                    content = self.shared[int(content)]
                elif inline is not None:
                    content = "".join(x.text or "" for x in inline.iter("{%s}t" % M))
                elif content is not None and cell.get("t") not in ("str", "e", "b"):
                    try: content = Decimal(content)
                    except InvalidOperation: pass
                self.values[pos] = content
                self.styles[pos] = cell.get("s", "")
                if formula is not None: self.existing[pos] = "formula"

    def enable_change_columns(self):
        """One-time migration from C:R period values to value/percent pairs."""
        periods = self.periods()
        if not periods:
            return
        # Do not expand twice. In already expanded files, headers follow the values.
        if any(self.get(3, col+1) in CHANGE_HEADER.values() for col in periods.values()):
            return
        if self.pending:
            raise ValueError('Enable percent columns before populating an XLSX worksheet')
        self.original = _remap_sheet_xml(self.original,
                                         lambda c: c if c < 3 else 3 + 2*(c-3))
        self.values.clear(); self.existing.clear(); self.styles.clear()
        self._read()
        self.layout_changed = True

    def ensure_change_headers(self):
        for period, col in self.periods().items():
            text = change_header(period)
            if self.get(3, col+1) != text:
                self.put(3, col+1, text, 's', overwrite=True)

    def ensure_percent_formulas(self):
        """Formula is blank for missing/zero priors; one delta after each value."""
        periods = self.periods()
        for row, label in self.rows().items():
            if row <= 3: continue
            normalized = norm(label)
            if normalized in SKIP_LABELS or 'additional sec line items' in normalized: continue
            if any(token in normalized for token in CHECK_TOKENS): continue
            if not any(isinstance(self.get(row, col), Decimal) or
                       (row,col) in self.existing for col in periods.values()):
                continue
            for period, col in periods.items():
                prior = prior_period(period)
                if prior not in periods: continue
                previous_col = periods[prior]
                if (row,col+1) in self.existing or self.has_value(row,col+1): continue
                curr_ref, previous_ref = f'{col_name(col)}{row}', f'{col_name(previous_col)}{row}'
                formula = (f'IF(OR(NOT(ISNUMBER({curr_ref})),NOT(ISNUMBER({previous_ref})),'
                           f'{previous_ref}=0),"",({curr_ref}-{previous_ref})/ABS({previous_ref}))')
                self.put(row, col+1, formula, kind='f')

    def get(self, row: int, col: int):
        if (row, col) in self.pending: return self.pending[row, col][1]
        return self.values.get((row, col))

    def has_value(self, row: int, col: int) -> bool:
        value = self.get(row, col)
        if value is None: return False
        return str(value).strip().lower() not in ("", "-", "–", "—", "n/a", "na")

    def put(self, row: int, col: int, value, kind: str = "n", overwrite=False) -> bool:
        if not overwrite and (self.has_value(row, col) or (row, col) in self.existing): return False
        old = self.get(row, col)
        if old == value: return False
        self.pending[(row, col)] = (kind, value)
        return True

    def periods(self) -> dict[str, int]:
        found = {}
        for col in range(3, 500):
            text = self.get(3, col)
            if text is not None and PERIOD_RE.fullmatch(str(text).strip()):
                found[str(text).strip()] = col
        return found

    def seed_periods(self, reference: dict[str, int]):
        """Give an initially empty worksheet the workbook's existing period headers."""
        if self.periods(): return
        for label, col in sorted(reference.items(), key=lambda item: item[1]):
            self.put(3, col, label, "s")

    def add_period(self, label: str) -> int:
        periods = self.periods()
        if label in periods: return periods[label]
        later = [col for name, col in periods.items() if period_sort_key(name) > period_sort_key(label)]
        col = min(later) if later else max(periods.values(), default=1) + (2 if periods else 2)
        if later:
            self._insert_change_pair(col)
        self.put(3, col, label, 's')
        self.put(3, col+1, change_header(label), 's')
        return col

    def _insert_change_pair(self, start: int):
        """Insert two physical columns, updating old values and Excel formulas."""
        pending = self.pending.copy()
        self.pending.clear()
        if self.layout_changed:
            # Materialize the one-time conversion before another insertion.
            self.original = self.output()
            self.layout_changed = False
            self.values.clear(); self.existing.clear(); self.styles.clear(); self._read()
        self.original = _remap_sheet_xml(self.original,
                                         lambda c: c if c < start else c+2)
        self.values.clear(); self.existing.clear(); self.styles.clear(); self._read()
        self.pending = {(r, c if c < start else c+2): item for (r,c),item in pending.items()}
        self.layout_changed = True

    def rows(self) -> dict[int, str]:
        result = {}
        for (row, col), value in {**self.values, **{k: v[1] for k, v in self.pending.items()}}.items():
            if col == 2 and value and str(value).strip(): result[row] = str(value).strip()
        return dict(sorted(result.items()))

    def new_item_row(self, label: str, section="ADDITIONAL SEC LINE ITEMS — REVIEW") -> int:
        rows = self.rows()
        header = next((row for row, text in rows.items() if text == section), None)
        if header is None:
            row = max([4, *rows.keys()]) + 2
            self.put(row, 2, section, "s")
            header = row
        return max([header, *rows.keys()]) + 1

    def _cell_style(self, row, col) -> str:
        if (row, col) in self.styles: return self.styles[row, col]
        if row == 3: return self.styles.get((3, 3), "")
        # Prefer same row numeric style; otherwise previous row same column/adjacent numeric column.
        for candidate in ((row, 3), (row, col-1), (row-1, col), (row-1, 3), (5, 3)):
            if candidate in self.styles: return self.styles[candidate]
        if col == 2:
            for (r, c), style in sorted(self.styles.items()):
                if c == 2 and r > 3: return style
        return ""

    def _cell_xml(self, row, col, kind, value) -> str:
        address = f"{col_name(col)}{row}"
        style = self._cell_style(row, col)
        s = f' s="{escape(style)}"' if style else ""
        if kind == "s":
            text = escape(str(value), {'"': '&quot;'})
            return f'<c r="{address}"{s} t="inlineStr"><is><t>{text}</t></is></c>'
        if kind == "f":
            style_id = (self.report_styles.get(('total_percent' if self.row_roles.get(row) == 'grand' else 'subtotal_percent' if self.row_roles.get(row) == 'subtotal' else 'review_percent' if self.row_roles.get(row) == 'review' else 'percent'))
                        if self.report_styles else self.percent_style)
            pct_style = f' s="{style_id}"' if style_id is not None else s
            return f'<c r="{address}"{pct_style}><f>{escape(str(value))}</f><v></v></c>'
        if isinstance(value, Decimal):
            value = format(value, "f")
        return f'<c r="{address}"{s}><v>{escape(str(value))}</v></c>'

    def output(self) -> str:
        if not self.pending: return self.original
        changes = defaultdict(list)
        for (row, col), (kind, value) in self.pending.items(): changes[row].append((col, kind, value))
        text = self.original
        # An empty worksheet may encode sheetData as a self-closing tag.
        text = re.sub(r'<sheetData\s*/>', '<sheetData></sheetData>', text, count=1)
        for row in sorted(changes, reverse=True):
            row_match = next((m for m in ROW_RE.finditer(text) if int(m.group(1)) == row), None)
            if row_match:
                old_row = row_match.group(0)
                if old_row.endswith("/>"):
                    body = ""; open_tag = old_row[:-2] + ">"
                else:
                    open_tag = old_row[:old_row.index(">") + 1]
                    body = old_row[len(open_tag):-len("</row>")]
                cells = list(CELL_RE.finditer(body))
                existing_cells = {col_number(m.group(1)): m for m in cells}
                for col, kind, value in sorted(changes[row], reverse=True):
                    replacement = self._cell_xml(row, col, kind, value)
                    if col in existing_cells:
                        m = existing_cells[col]
                        body = body[:m.start()] + replacement + body[m.end():]
                    else:
                        later = next((m for m in CELL_RE.finditer(body) if col_number(m.group(1)) > col), None)
                        if later: body = body[:later.start()] + replacement + body[later.start():]
                        else: body += replacement
                    existing_cells = {col_number(m.group(1)): m for m in CELL_RE.finditer(body)}
                new_row = open_tag + body + "</row>"
                text = text[:row_match.start()] + new_row + text[row_match.end():]
            else:
                contents = "".join(self._cell_xml(row, col, kind, val) for col, kind, val in sorted(changes[row]))
                new_row = f'<row r="{row}">{contents}</row>'
                sd_start = text.find("<sheetData")
                if sd_start < 0: raise ValueError(f"Invalid sheet XML: {self.name} missing sheetData")
                after = text.find(">", sd_start) + 1
                following = next((m for m in ROW_RE.finditer(text, after) if int(m.group(1)) > row), None)
                insert = following.start() if following else text.find("</sheetData>", after)
                if insert < 0: raise ValueError(f"Invalid sheet XML: {self.name} missing /sheetData")
                text = text[:insert] + new_row + text[insert:]
        # Validate XML before writing ZIP.
        ET.fromstring(text)
        return text


def new_empty_workbook_parts() -> dict[str, bytes]:
    """Construct a genuine three-sheet OOXML workbook using only the standard library.

    Supports zero-byte files created using `touch`. There is no need to ship a
    separate template or to install spreadsheet-writing dependencies on a VPS.
    Existing nonempty XLSX files are never replaced by this initialization.
    """
    sheets = ("Balance Sheets", "Income Statements", "Cashflow")
    pkg_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    doc_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sheet_rel = doc_rel + "/worksheet"
    document_rel = doc_rel + "/officeDocument"
    style_rel = doc_rel + "/styles"
    def content(text: str) -> bytes:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + text).encode('utf-8')

    overrides = ''.join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, 4))
    types = content(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + overrides + '</Types>')
    root_rels = content(
        f'<Relationships xmlns="{pkg_rel}">'
        f'<Relationship Id="rId1" Type="{document_rel}" Target="xl/workbook.xml"/>'
        '</Relationships>')
    workbook_rels = content(
        f'<Relationships xmlns="{pkg_rel}">'
        + ''.join(f'<Relationship Id="rId{i}" Type="{sheet_rel}" '
                  f'Target="worksheets/sheet{i}.xml"/>' for i in range(1, 4))
        + f'<Relationship Id="rId4" Type="{style_rel}" Target="styles.xml"/>'
        + '</Relationships>')
    workbook = content(
        f'<workbook xmlns="{M}" xmlns:r="{R}"><sheets>'
        + ''.join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                  for i, name in enumerate(sheets, 1))
        + '</sheets><calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>')
    # Two styles: general values and a bold title/header. Workbook() appends its
    # existing 0.00%-format style for percent-change formulas on first opening.
    styles = content(
        f'<styleSheet xmlns="{M}">'
        '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font>'
        '<font><b/><sz val="11"/><name val="Aptos"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>')
    parts = {
        '[Content_Types].xml': types,
        '_rels/.rels': root_rels,
        'xl/workbook.xml': workbook,
        'xl/_rels/workbook.xml.rels': workbook_rels,
        'xl/styles.xml': styles,
    }
    for i, name in enumerate(sheets, 1):
        title = escape(name)
        sheet_xml = content(
            f'<worksheet xmlns="{M}">'
            '<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
            '<pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>'
            '</sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            '<cols><col min="1" max="1" width="3" customWidth="1"/>'
            '<col min="2" max="2" width="42" customWidth="1"/>'
            '<col min="3" max="500" width="17" customWidth="1"/></cols>'
            '<sheetData>'
            f'<row r="1"><c r="B1" s="1" t="inlineStr"><is><t>{title}</t></is></c></row>'
            '<row r="3"><c r="B3" s="1" t="inlineStr"><is><t>Line Item</t></is></c></row>'
            '</sheetData></worksheet>')
        parts[f'xl/worksheets/sheet{i}.xml'] = sheet_xml
    return parts


def validate_xlsx(path: Path, *, allow_empty: bool = False) -> None:
    """Reject mislabeled/partial XLSX files with a useful local-file diagnosis.

    XLSX is an OOXML ZIP archive; a .xlsx suffix by itself does not make one.
    This check runs before contacting SEC, so malformed files do not waste
    requests or prevent other companies from being processed.
    """
    path = Path(path)
    if allow_empty and path.stat().st_size == 0:
        # Empty user-created placeholders are valid inputs to Workbook(), which
        # builds an actual OOXML workbook on the first successful write.
        return
    try:
        with path.open("rb") as file:
            beginning = file.read(256)
        if beginning.startswith(b"version https://git-lfs.github.com/spec/v1"):
            cause = "Git LFS pointer, not the downloaded workbook (try git lfs pull)"
        elif beginning.startswith(bytes.fromhex("d0cf11e0a1b11ae1")):
            cause = "legacy .xls or an encrypted Office document, not an OOXML .xlsx file"
        elif beginning.lstrip().lower().startswith((b"<!doctype html", b"<html")):
            cause = "HTML page, possibly downloaded instead of the workbook"
        elif beginning.startswith((b"PK\x03\x04", b"PK\x05\x06")):
            cause = "ZIP archive is damaged or incomplete"
        else:
            cause = "file content is not an XLSX ZIP archive (possibly CSV/text or a renamed file)"
        if not beginning.startswith((b"PK\x03\x04", b"PK\x05\x06")):
            raise ValueError(cause)
        try:
            with ZipFile(path) as archive:
                names = set(archive.namelist())
                required = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
                missing = sorted(required - names)
                if missing:
                    raise ValueError("missing required XLSX content: " + ", ".join(missing))
                broken_entry = archive.testzip()
                if broken_entry:
                    raise ValueError(f"corrupted ZIP entry: {broken_entry}")
        except BadZipFile as exc:
            raise ValueError(cause) from exc
    except (OSError, ValueError) as exc:
        raise ValueError(
            f"Cannot open {path.name!r} as an Excel .xlsx workbook: {exc}. "
            "Run `file` and `unzip -t` on this file. Restore/export a genuine "
            "Excel Workbook (.xlsx); do not just rename a .csv or .xls file."
        ) from exc


class Workbook:
    def __init__(self, path: Path):
        validate_xlsx(path, allow_empty=True)
        self.path = path
        self.is_placeholder = path.stat().st_size == 0
        self.sheets: dict[str, XMLSheet] = {}
        self.paths: dict[str, str] = {}
        if self.is_placeholder:
            self.raw = new_empty_workbook_parts()
        else:
            with ZipFile(path) as z:
                self.raw = {name: z.read(name) for name in z.namelist()}
        self.percent_style = None
        self.report_styles = {}
        if "xl/styles.xml" in self.raw:
            self.raw["xl/styles.xml"], self.percent_style = _ensure_percent_style(self.raw["xl/styles.xml"])
            self.raw["xl/styles.xml"], self.report_styles = report_styles(self.raw["xl/styles.xml"])
        # Recalculate newly written change formulas when Excel opens the file.
        if "xl/workbook.xml" in self.raw:
            xml = self.raw["xl/workbook.xml"].decode("utf-8")
            if '<calcPr' in xml:
                xml = re.sub(r'<calcPr\b[^>]*/>', '<calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/>', xml, count=1)
            else:
                xml = xml.replace('</workbook>', '<calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>')
            self.raw["xl/workbook.xml"] = xml.encode("utf-8")
        shared = []
        if "xl/sharedStrings.xml" in self.raw:
            root = ET.fromstring(self.raw["xl/sharedStrings.xml"])
            shared = ["".join(item.text or "" for item in si.iter("{%s}t" % M)) for si in root.findall("{%s}si" % M)]
        root = ET.fromstring(self.raw["xl/workbook.xml"])
        rels = ET.fromstring(self.raw["xl/_rels/workbook.xml.rels"])
        targets = {rel.get("Id"): rel.get("Target", "") for rel in rels}
        for node in root.findall("{%s}sheets/{%s}sheet" % (M, M)):
            name = node.get("name")
            target = targets[node.get("{%s}id" % R)].lstrip("/")
            if not target.startswith("xl/"): target = "xl/" + target
            if target not in self.raw: raise ValueError("Missing worksheet XML: " + target)
            self.paths[name] = target
            self.sheets[name] = XMLSheet(self.raw[target].decode("utf-8"), shared, name)
            self.sheets[name].percent_style = self.percent_style
            self.sheets[name].report_styles = self.report_styles

    def changed(self) -> bool:
        return self.is_placeholder or any(sheet.pending or sheet.layout_changed for sheet in self.sheets.values())

    def save(self, *, dry_run=False, backup=True) -> bool:
        if dry_run or not self.changed(): return False
        if self.is_placeholder and self.path.stat().st_size != 0:
            raise ValueError(f"Refusing to overwrite {self.path}: placeholder was modified by another process")
        fd, filename = tempfile.mkstemp(suffix=".xlsx", dir=self.path.parent)
        os.close(fd)
        try:
            if self.is_placeholder:
                with ZipFile(filename, "w", compression=ZIP_DEFLATED) as target:
                    for part, original in self.raw.items():
                        data = original
                        for name, sheet_path in self.paths.items():
                            if part == sheet_path:
                                data = self.sheets[name].output().encode("utf-8")
                                break
                        target.writestr(part, data)
            else:
                with ZipFile(filename, "w") as target, ZipFile(self.path) as source:
                    for info in source.infolist():
                        data = self.raw[info.filename]
                        for name, sheet_path in self.paths.items():
                            if info.filename == sheet_path:
                                data = self.sheets[name].output().encode("utf-8")
                                break
                        target.writestr(info, data)
            with ZipFile(filename) as test:
                if test.testzip() is not None: raise ValueError("Broken ZIP archive produced")
                for path in self.paths.values(): ET.fromstring(test.read(path))
            if not self.is_placeholder and Path(filename).read_bytes() == self.path.read_bytes():
                return False  # Idempotent reformat must not create repeated backups.
            if backup and not self.is_placeholder:
                stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
                backup_path = self.path.with_name(self.path.name + f".bak.{stamp}")
                if backup_path.exists():
                    backup_path = self.path.with_name(self.path.name + f".bak.{stamp}-{os.getpid()}")
                shutil.copy2(self.path, backup_path)
                LOG.info("Backup %s", backup_path)
            os.replace(filename, self.path)
            return True
        finally:
            if os.path.exists(filename): os.unlink(filename)


# Canonical presentation is intentionally an analytical presentation, not an assertion
# that every SEC registrant is required to use a single GAAP account order.
REPORT_LAYOUT_VERSION = 3
REPORT_STYLE_MARKER = 'EDGAR_REPORT_STYLES_V2'
REPORT_COLORS = {
    'navy': '17324D', 'soft': 'E8F0F8', 'total': 'D7E5F2',
    'review': 'FFF0CB', 'muted': '49647C', 'blue': '1E40AF',
}


def report_styles(styles: bytes) -> tuple[bytes, dict[str, int]]:
    """Add reusable OOXML cell styles, returning the existing indices on later runs."""
    data = styles.decode('utf-8')
    marker = re.search(r'<!-- ' + REPORT_STYLE_MARKER + r' (\{[^<]+\}) -->', data)
    if marker:
        return styles, {name: int(index) for name, index in json.loads(marker[1]).items()}
    root = ET.fromstring(data)
    def count(name):
        parent = root.find('{%s}%s' % (M, name))
        if parent is None: raise ValueError('Unsupported XLSX styles.xml: missing ' + name)
        return len(parent)
    font_start, fill_start, border_start, xf_start = (count(s) for s in ('fonts', 'fills', 'borders', 'cellXfs'))
    fmt_ids = [int(n.get('numFmtId')) for n in root.findall('.//{%s}numFmt' % M) if n.get('numFmtId')]
    fmt_start = max([163, *fmt_ids]) + 1
    fmts = {
        'number': '#,##0.00;[Red](#,##0.00);"–"',
        'per_share': '0.0000;[Red](0.0000);"–"',
        'percent': '0.0%;[Red](0.0%);"–"',
    }
    from xml.sax.saxutils import quoteattr
    formats_xml = ''.join(f'<numFmt numFmtId="{fmt_start+i}" formatCode={quoteattr(fmt)}/>'
                          for i, fmt in enumerate(fmts.values()))
    if '<numFmts' in data:
        data = re.sub(r'(<numFmts\b[^>]*\bcount=")[0-9]+(")',
                      lambda m: m[1] + str(count('numFmts') + len(fmts)) + m[2], data, count=1)
        data = data.replace('</numFmts>', formats_xml + '</numFmts>', 1)
    else:
        data = data.replace('<fonts', f'<numFmts count="{len(fmts)}">{formats_xml}</numFmts><fonts', 1)
    new_fonts = [
        '<font><sz val="10"/><name val="Aptos"/><color rgb="FF25384A"/></font>',
        '<font><sz val="10"/><name val="Aptos"/><color rgb="FF1E40AF"/></font>',
        '<font><b/><sz val="10"/><name val="Aptos"/><color rgb="FFFFFFFF"/></font>',
        '<font><b/><sz val="10"/><name val="Aptos"/><color rgb="FF17324D"/></font>',
        '<font><b/><sz val="10"/><name val="Aptos"/><color rgb="FF49647C"/></font>',
    ]
    new_fills = [
        f'<fill><patternFill patternType="solid"><fgColor rgb="FF{REPORT_COLORS[c]}"/><bgColor indexed="64"/></patternFill></fill>'
        for c in ('navy', 'soft', 'total', 'review')
    ]
    new_borders = [
        '<border><left/><right/><top style="thin"><color rgb="FF8EA4B9"/></top><bottom/><diagonal/></border>',
        '<border><left/><right/><top style="medium"><color rgb="FF17324D"/></top><bottom style="double"><color rgb="FF17324D"/></bottom><diagonal/></border>',
    ]
    def insert_collection(name, start, addition):
        nonlocal data
        data = re.sub(r'(<'+name+r'\b[^>]*\bcount=")[0-9]+(")',
                      lambda m: m[1] + str(start + len(addition)) + m[2], data, count=1)
        data = data.replace('</'+name+'>', ''.join(addition) + '</'+name+'>', 1)
    insert_collection('fonts', font_start, new_fonts)
    insert_collection('fills', fill_start, new_fills)
    insert_collection('borders', border_start, new_borders)
    # Styles have separate variants for headings, details, subtotals and grand totals.
    rules = {
        'label':           (0, 0, 0, 0, 0),
        'number':          (1, 0, 0, fmt_start, 0),
        'per_share':       (1, 0, 0, fmt_start+1, 0),
        'percent':         (0, 0, 0, fmt_start+2, 0),
        'header_label':    (2, 0, 0, 0, 0),
        'header_number':   (2, 0, 0, 0, 0),
        'section_label':   (2, 0, 0, 0, 0),
        'section_number':  (2, 0, 0, 0, 0),
        'subtotal_label':  (3, 1, 0, 0, 1),
        'subtotal_number': (3, 1, 0, fmt_start, 1),
        'subtotal_percent':(3, 1, 0, fmt_start+2, 1),
        'total_label':     (3, 2, 1, 0, 1),
        'total_number':    (3, 2, 1, fmt_start, 1),
        'total_percent':   (3, 2, 1, fmt_start+2, 1),
        'review_label':    (4, 3, 0, 0, 0),
        'review_number':   (1, 3, 0, fmt_start, 0),
        'review_percent':  (0, 3, 0, fmt_start+2, 0),
    }
    # Rules' font/fill/border coordinates above are relative to newly appended collections.
    # Use a common minimal cellXf to avoid inheriting a company's custom number format.
    new_xfs = []
    result = {}
    for index, (name, (font, fill, border, num_fmt, align)) in enumerate(rules.items()):
        result[name] = xf_start + index
        alignment = ('<alignment horizontal="right"/>' if name.endswith(('number', 'percent'))
                     or name == 'per_share' else '<alignment horizontal="left"/>')
        attrs = (f'numFmtId="{num_fmt}" fontId="{font_start+font}" '
                 f'fillId="{0 if name in ("label","number","per_share","percent") else fill_start+fill}" borderId="{border_start+border}" '
                 'xfId="0" applyFont="1" applyFill="1" applyBorder="1" '
                 'applyNumberFormat="1" applyAlignment="1"')
        new_xfs.append(f'<xf {attrs}>{alignment}</xf>')
    insert_collection('cellXfs', xf_start, new_xfs)
    data = data.replace('</styleSheet>',
                        '<!-- '+REPORT_STYLE_MARKER+' '+json.dumps(result, separators=(',',':'))+' -->'
                        '</styleSheet>', 1)
    ET.fromstring(data)
    return data.encode('utf-8'), result


# These are classification rules for analytical spreads, not a replacement for a
# registrant's GAAP notes. Companies with unclassified balance sheets are not
# assigned arbitrary current/noncurrent totals.
def _item_kind(label: str, typ: str) -> str:
    key = norm(label)
    if 'reconciliation' in key or 'less reported total' in key or any(token in key for token in CHECK_TOKENS): return 'check'
    if typ == 'balance':
        if (key.startswith('total liabilities') and ('equity' in key or 'stockholders' in key or
                                                    'shareholders' in key)): return 'grand'
        if key in ('total assets', 'total liabilities', 'total equity', 'total stockholders equity',
                   'total shareholders equity', 'total equity including noncontrolling interests'): return 'grand'
        if key.startswith('total ') or key.startswith('net investment in real estate') and False: return 'subtotal'
    elif typ == 'income':
        if (key.startswith('net income') or key.startswith('net earnings') or
                key.startswith('profit loss') or key.startswith('total net income')): return 'grand'
        if (key.startswith('total ') or 'gross profit' in key or
                key.startswith('operating income') or key.startswith('income loss from operations') or
                'before income taxes' in key or 'before taxes' in key): return 'subtotal'
    elif typ == 'cashflow':
        if ('cash and cash equivalents end' in key or 'cash cash equivalents and restricted cash end' in key or
                key.startswith('cash at end') or 'cash at end of period' in key): return 'grand'
        if ((key.startswith('net cash') or key.startswith('cash flows from')) and any(x in key for x in
                ('operating activities','investing activities','financing activities'))): return 'subtotal'
        if key.startswith('net increase decrease in cash') or key.startswith('net change in cash'): return 'subtotal'
    return 'detail'


def _sort_priority(label: str, typ: str, group: str) -> tuple[int, str]:
    k = norm(label)
    def first(matches):
        for index, pattern in enumerate(matches):
            if re.search(pattern, k): return index
        return len(matches)
    if typ == 'balance':
        if group.startswith('asset'):
            patterns = [r'cash and cash equivalents', r'cash .* segregated', r'restricted cash',
                        r'short term investment', r'marketab', r'securities borrowed',
                        r'receiv', r'deposit', r'inventor', r'prepaid', r'deferred',
                        r'property', r'land', r'real estate', r'building', r'development',
                        r'investment', r'intangible', r'goodwill', r'lease', r'other']
        elif group.startswith('liab'):
            patterns = [r'accounts payable', r'payable', r'accrued', r'securities loaned',
                        r'deferred revenue', r'unearned', r'customer', r'lease',
                        r'mortgage', r'borrow', r'loan', r'debt', r'reserve', r'tax', r'other']
        else:
            patterns = [r'common stock', r'preferred stock', r'additional paid', r'treasury',
                        r'retained', r'accumulated deficit', r'comprehensive', r'noncontrolling',r'other']
    elif typ == 'income':
        patterns = ([r'transaction',r'product',r'service',r'subscription',r'rental',
                     r'land',r'premium',r'interest',r'other'] if group=='revenue' else
                    [r'cost of sales',r'cost of revenue',r'operating costs',r'claim',
                     r'brokerage',r'technology',r'operation',r'marketing',
                     r'general and administrative',r'depreciation',r'other'])
    else:
        patterns = [r'net income',r'depreciation',r'amortization',r'stock based',r'deferred',
                    r'gain',r'loss',r'accounts receivable',r'accounts payable',
                    r'change',r'purchase',r'proceeds',r'payment',r'other']
    return first(patterns), k


def _is_heading(label: str) -> bool:
    k = norm(label)
    return k in {
        'assets','liabilities','equity','stockholders equity','shareholders equity',
        'current assets','noncurrent assets','current liabilities','noncurrent liabilities',
        'revenues','revenue','income','expenses','operating expenses','other',
        'other income and taxes','other income expenses','other income expense',
        'cash flows from operating activities','cash flows from investing activities',
        'cash flows from financing activities','operating activities', 'investing activities',
        'financing activities', 'earnings per share usd share','earnings per share',
        'weighted average shares outstanding millions','weighted average shares outstanding',
        'supplemental equity detail reconciliations','supplemental equity detail and reconciliations','reconciliation checks all should equal zero',
        'additional sec line items review','uncategorized items requiring review',
        'supplemental cash flow information', 'operating expenses and other costs',
        'cost of revenue', 'gross profit', 'operating profit', 'pretax income',
        'net income','income taxes','other income and expenses', 'weighted average shares',
        'cash reconciliation','supplemental disclosures','reconciliation checks',
        'mezzanine equity',
    }


def _balance_bucket(label: str, context: str, current: str, has_current_assets: bool,
                    has_current_liab: bool) -> str:
    k = norm(label)
    if 'reconciliation' in k or 'less reported total' in k or 'component sum difference' in k or any(t in k for t in CHECK_TOKENS): return 'checks'
    if k.startswith('assets liabilities') or 'assets less liabilities' in k: return 'checks'
    if 'total liabilities' in k and ('equity' in k or 'stockholders' in k or 'shareholders' in k):
        return 'balance_total'
    if k in ('total assets','total asset'): return 'assets_total'
    if k in ('total liabilities',): return 'liab_total'
    if k in ('total current assets',): return 'assets_current'
    if k in ('total current liabilities',): return 'liab_current'
    if k.startswith('total ') and ('stockholders equity' in k or 'shareholders equity' in k or k=='total equity'):
        return 'equity'
    if 'mezzanine equity' in k or 'redeemable noncontrolling' in k: return 'mezzanine'
    if context == 'review':
        if any(t in k for t in ('asset','receivable','cash and cash','investment','property','goodwill','intangible','deposit','inventory')):
            context = 'assets'
        elif any(t in k for t in ('liability','payable','debt','borrowing','lease obligation','accrued')):
            context = 'liab'
        elif any(t in k for t in ('equity','stock','retained earnings','accumulated deficit','noncontrolling')):
            context = 'equity'
        else: return 'review'
    if context == 'assets' and ('equity securit' in k or 'equity method investment' in k or
                                'investments in equity' in k):
        return 'assets_noncurrent' if has_current_assets and current=='noncurrent' else 'assets_other'
    if ('equity' in k or 'stockholders' in k or 'shareholders' in k or
            'noncontrolling interests' in k or 'retained earnings' in k or
            'paid in capital' in k or 'accumulated comprehensive' in k or 'accumulated deficit' in k or
            'treasury stock' in k or 'common stock' in k): return 'equity'
    if context in ('assets','liab','equity'):
        if context == 'assets':
            if has_current_assets:
                if 'noncurrent' in k or 'non current' in k: return 'assets_noncurrent'
                if 'current' in k or current == 'current': return 'assets_current'
                return 'assets_noncurrent'
            return 'assets_other'
        if context == 'liab':
            if has_current_liab:
                if 'noncurrent' in k or 'non current' in k: return 'liab_noncurrent'
                if 'current' in k or current == 'current': return 'liab_current'
                return 'liab_noncurrent'
            return 'liab_other'
        return 'equity'
    return 'review'


def _income_bucket(label: str, context: str) -> str:
    k = norm(label)
    if 'reconciliation' in k or any(t in k for t in CHECK_TOKENS): return 'checks'
    if ('weighted average' in k or 'weighted average shares' in k): return 'shares'
    if ('earnings per share' in k or 'income per share' in k or
            re.search(r'^(basic|diluted) (net )?(earnings|income|loss) per share', k)): return 'eps'
    if ('noncontrolling' in k and ('net income' in k or 'net earnings' in k)): return 'netincome'
    if ('net income' in k or 'net earnings' in k or 'net loss' in k or
            'profit loss' == k): return 'netincome'
    if ('before taxes' in k or 'before income taxes' in k or 'income before tax' in k or
            'earnings before tax' in k): return 'pretax'
    if ('income tax' in k or 'tax expense' in k or 'benefit from taxes' in k): return 'taxes'
    if ('operating income' in k or 'income loss from operations' in k or
            'income from operations' in k or 'operating profit' in k): return 'operating_profit'
    if ('gross profit' in k or 'gross margin' in k): return 'gross'
    if ('cost of sales' in k or 'cost of revenue' in k or 'cost of goods' in k or
            'cost of services' in k): return 'cost'
    if ('total operating expense' in k or 'total expense' in k or 'total costs' in k): return 'opex'
    if ('interest expense' in k or 'interest income' in k or 'interest revenue' in k and context != 'revenue'
        or 'sale or disposal' in k or 'extinguishment' in k or 'equity in earnings' in k or
        'investment gain' in k or 'other income' in k or 'total other' in k):
        return 'nonop' if context != 'revenue' else 'revenue'
    if ('total revenue' in k or 'net revenues' in k or 'total net revenue' in k): return 'revenue'
    if context == 'review':
        if any(t in k for t in ('revenue','sales','premium','land sales','rental')): return 'revenue'
        if any(t in k for t in ('expense','cost','marketing','research','development')): return 'opex'
        return 'review'
    return context if context in ('revenue','cost','opex','nonop') else 'review'


def _cashflow_bucket(label: str, context: str) -> str:
    k = norm(label)
    if ('net increase' in k or 'net decrease' in k or 'net change' in k or
            'effect of exchange rate' in k or 'effect of foreign exchange' in k or
            'cash and cash equivalents' in k and ('beginning' in k or 'end of' in k)): return 'reconciliation'
    for group in ('operating','investing','financing'):
        if group+' activities' in k or group+' cash flows' in k:
            return group
    if any(t in k for t in ('cash paid for interest','cash paid for income taxes','supplemental disclosure')):
        return 'supplemental'
    if context == 'review':
        if any(t in k for t in ('depreciation','amortization','working capital','noncash','non cash',
                                'receivables','accounts payable','net income','stock based compensation')): return 'operating'
        if (any(t in k for t in ('purchase of property','purchases of property','sale of property','sales of property',
                                'purchase of investment','purchases of investment','acquisition of business',
                                'capital expenditure','payments to acquire','payments for investments'))): return 'investing'
        if any(t in k for t in ('repayment','dividends paid','issuance of common stock','issuance of debt',
                                'proceeds from debt','proceeds from borrowing','proceeds from issuance',
                                'proceeds from issuing','shares repurchased')): return 'financing'
        return 'review'
    return context if context in ('operating','investing','financing') else 'review'


def _classify_rows(sheet, typ: str):
    rows = sheet.rows()
    normalized = {norm(v) for v in rows.values()}
    has_ca = 'total current assets' in normalized
    has_cl = 'total current liabilities' in normalized
    ctx = 'assets' if typ=='balance' else ('revenue' if typ=='income' else 'operating')
    current = ''
    results = []
    for row,label in rows.items():
        if row <= 3: continue
        key = norm(label)
        has_content = any(sheet.has_value(row,c) or (row,c) in sheet.existing
                          for c in range(3,max(sheet.periods().values(),default=3)+2))
        # The existing filing section captions provide context for poorly named
        # company-specific items; once standardized, captions supply it again.
        if typ == 'balance':
            if key in ('assets','current assets','noncurrent assets'):
                ctx = 'assets';current='current' if key == 'current assets' else 'noncurrent' if key=='noncurrent assets' else ''
            elif key in ('liabilities','current liabilities','noncurrent liabilities'):
                ctx = 'liab';current='current' if key == 'current liabilities' else 'noncurrent' if key == 'noncurrent liabilities' else ''
            elif key in ('equity','stockholders equity','shareholders equity'): ctx='equity';current=''
            elif key in ('additional sec line items review','uncategorized items requiring review'): ctx='review'
            elif key=='total current assets':ctx='assets';current='noncurrent'
            elif key=='total current liabilities':ctx='liab';current='noncurrent'
            elif key=='total assets':ctx='liab';current=''
            elif key=='total liabilities':ctx='equity';current=''
        elif typ=='income':
            if key in ('revenues','revenue'):ctx='revenue'
            elif key in ('expenses','operating expenses','operating expenses and other costs'):ctx='opex'
            elif key=='cost of revenue':ctx='cost'
            elif key=='gross profit':ctx='gross'
            elif key=='operating profit':ctx='operating_profit'
            elif key in ('other income and taxes','other income expenses','other income and expenses') or (key=='other' and not has_content):ctx='nonop'
            elif key in ('additional sec line items review','uncategorized items requiring review','uncategorized items review'):ctx='review'
        else:
            if key in ('cash flows from operating activities','operating activities'):ctx='operating'
            elif key in ('cash flows from investing activities','investing activities'):ctx='investing'
            elif key in ('cash flows from financing activities','financing activities'):ctx='financing'
            elif key in ('additional sec line items review','uncategorized items requiring review','uncategorized items review'):ctx='review'
        # Preserve every labeled data item, but replace old section captions.
        if _is_heading(label) and not has_content: continue
        if typ == 'balance': group=_balance_bucket(label,ctx,current,has_ca,has_cl)
        elif typ == 'income': group=_income_bucket(label,ctx)
        else: group=_cashflow_bucket(label,ctx)
        kind = _item_kind(label,typ)
        results.append((row,label,group,kind))
    return results,has_ca,has_cl


def _section_sort(typ: str, rows, has_ca=False, has_cl=False):
    if typ=='balance':
        rank = {'assets_current':10,'assets_noncurrent':20,'assets_other':20,'assets_total':29,
                'liab_current':30,'liab_noncurrent':40,'liab_other':40,'liab_total':49,
                'mezzanine':50,'equity':60,'balance_total':70,'checks':80,'review':90}
    elif typ=='income':
        rank = {x:i*10 for i,x in enumerate(('revenue','cost','gross','opex','operating_profit',
                'nonop','pretax','taxes','netincome','eps','shares','checks','review'),1)}
    else:
        rank = {x:i*10 for i,x in enumerate(('operating','investing','financing',
                'reconciliation','supplemental','checks','review'),1)}
    def key(item):
        old_row,label,group,kind = item
        # Totals come after their components in every section, including cashflow.
        kind_order = 3 if kind=='grand' else 2 if kind=='subtotal' else 0
        if typ=='income' and group=='netincome':
            norm_label=norm(label)
            if 'attributable to' in norm_label and 'noncontrolling' not in norm_label:kind_order=3
            elif 'noncontrolling' in norm_label:kind_order=1
            elif 'consolidated' in norm_label:kind_order=0
            else:kind_order=2
        if typ=='balance' and group=='equity' and kind in ('subtotal','grand'):
            kind_order=3 if norm(label)=='total equity' else 2
        return (rank.get(group,99),kind_order,_sort_priority(label,typ,group),old_row)
    return sorted(rows,key=key)


def _section_headings(typ: str, group: str, seen: set[str], has_ca: bool, has_cl: bool) -> list[str]:
    headings = []
    if typ=='balance':
        if group.startswith('assets') and 'ASSETS' not in seen:headings.append('ASSETS')
        if group=='assets_current' and has_ca:headings.append('CURRENT ASSETS')
        if group=='assets_noncurrent' and has_ca:headings.append('NONCURRENT ASSETS')
        if group.startswith('liab') and 'LIABILITIES' not in seen:headings.append('LIABILITIES')
        if group=='liab_current' and has_cl:headings.append('CURRENT LIABILITIES')
        if group=='liab_noncurrent' and has_cl:headings.append('NONCURRENT LIABILITIES')
        headings += {'mezzanine':['MEZZANINE EQUITY'],'equity':["STOCKHOLDERS' EQUITY"],
                     'checks':['RECONCILIATION CHECKS'],'review':['UNCATEGORIZED ITEMS — REVIEW']}.get(group,[])
    elif typ=='income':
        heading = {'revenue':'REVENUE','cost':'COST OF REVENUE','gross':'GROSS PROFIT',
                   'opex':'OPERATING EXPENSES AND OTHER COSTS','operating_profit':'OPERATING PROFIT',
                   'nonop':'OTHER INCOME AND EXPENSES','pretax':'PRETAX INCOME',
                   'taxes':'INCOME TAXES','netincome':'NET INCOME',
                   'eps':'EARNINGS PER SHARE','shares':'WEIGHTED-AVERAGE SHARES',
                   'checks':'RECONCILIATION CHECKS','review':'UNCATEGORIZED ITEMS — REVIEW'}.get(group)
        if heading:headings.append(heading)
    else:
        heading = {'operating':'OPERATING ACTIVITIES','investing':'INVESTING ACTIVITIES',
                   'financing':'FINANCING ACTIVITIES','reconciliation':'CASH RECONCILIATION',
                   'supplemental':'SUPPLEMENTAL DISCLOSURES','checks':'RECONCILIATION CHECKS',
                   'review':'UNCATEGORIZED ITEMS — REVIEW'}.get(group)
        if heading:headings.append(heading)
    return [h for h in headings if h not in seen]


def _new_heading_row(row: int, title: str, last_col: int, styles: dict[str,int]) -> str:
    text = escape(title)
    # Continuous deep-navy band; no merges that could break nearby period cells.
    cells = [f'<c r="B{row}" s="{styles["section_label"]}" t="inlineStr"><is><t>{text}</t></is></c>']
    cells.extend(f'<c r="{col_name(c)}{row}" s="{styles["section_number"]}"/>' for c in range(3,last_col+1))
    return f'<row r="{row}" ht="19" customHeight="1">'+''.join(cells)+'</row>'


def _update_formula_rows(formula: str, row_map: dict[int,int]) -> str:
    """Update existing Excel formula refs and preserve original SUM components."""
    single_sum = re.compile(r'\bSUM\(\s*(\$?[A-Z]{1,3})(\$?)(\d+):(\$?[A-Z]{1,3})(\$?)(\d+)\s*\)', re.I)
    protected = []
    def sum_replace(m):
        first_col, first_row = col_number(m[1].lstrip('$')), int(m[3])
        last_col, last_row = col_number(m[4].lstrip('$')), int(m[6])
        if first_col != last_col or last_row < first_row or last_row-first_row > 300:
            return m[0]
        refs = [m[1]+m[2]+str(row_map[i]) for i in range(first_row,last_row+1) if i in row_map]
        if not refs: return m[0]
        index = len(protected)
        protected.append('SUM('+','.join(refs)+')')
        return f'__SEC_RANGE_{index}__'
    formula = single_sum.sub(sum_replace, formula)
    def replace(m):
        old = int(m['row'])
        if old not in row_map: return m[0]
        return f'{m["colabs"]}{m["col"]}{m["rowabs"]}{row_map[old]}'
    formula = FORMULA_REF.sub(replace, formula)
    for i,value in enumerate(protected):
        formula = formula.replace(f'__SEC_RANGE_{i}__',value)
    return formula


def _restyle_data_row(source: str, old_row: int, new_row: int, kind: str,
                      label: str, last_col: int, value_columns: set[int], styles: dict[str,int],
                      row_map: dict[int,int]) -> str:
    kind = 'review' if kind=='review' else kind
    base = 'total' if kind=='grand' else 'subtotal' if kind=='subtotal' else 'review' if kind=='review' else ''
    m = re.match(r'<row\b[^>]*>',source)
    if not m:raise ValueError('Invalid row XML')
    opening = re.sub(r'(\br=")[0-9]+(")',lambda m:m[1]+str(new_row)+m[2],m[0],count=1)
    opening = re.sub(r'\bspans="[^"]*"','',opening)
    if kind in ('grand','subtotal'):
        opening = re.sub(r'\s+ht="[^"]*"|\s+customHeight="[^"]*"','',opening[:-1])+' ht="19" customHeight="1">'
    cells = {}
    for match in CELL_RE.finditer(source[len(m[0]):-len('</row>')]):
        c = col_number(match[1]);cell = match[0]
        if c in value_columns and c-1 in value_columns:raise ValueError('Duplicate period columns')
        # Rebuild comparison formulas after reorder using the new row addresses.
        if c > 2 and c not in value_columns:continue
        if c == 2:role = base+'_label' if base else 'label'
        elif 'per share' in norm(label) and 'percent' not in norm(label):role='per_share'
        else:role = base+'_number' if base else 'number'
        style=styles.get(role,styles['number'])
        # Normalize OOXML attribute order for idempotent offline formatting.
        old_open = re.match(r'<c\b[^>]*>',cell)
        if not old_open:raise ValueError('Invalid cell XML')
        o = old_open[0]
        extras = [(k,v) for k,v in re.findall(r'\s+([A-Za-z_:][A-Za-z0-9_:]*)="([^"]*)"',o)
                  if k not in ('r','s')]
        opener = f'<c r="{col_name(c)}{new_row}" s="{style}"' + ''.join(
            f' {k}="{v}"' for k,v in extras) + ('/>' if o.endswith('/>') else '>')
        cell = opener + cell[len(o):]
        def formula_update(f):
            old = f[2]
            return f[1]+escape(_update_formula_rows(old,row_map))+f[3]
        cell = re.sub(r'(<f(?:\s+[^>]*)?>)(.*?)(</f>)',formula_update,cell,flags=re.S)
        cells[c]=cell
    for col in range(2,last_col+1):
        if col in cells:continue
        if col == 2:style = styles[base+'_label'] if base else styles['label']
        elif col in value_columns:
            style = styles[base+'_number'] if base else (styles['per_share'] if 'per share' in norm(label) else styles['number'])
        else:
            style = styles[base+'_percent'] if base else styles['percent']
        cells[col]=f'<c r="{col_name(col)}{new_row}" s="{style}"/>'
    return opening+''.join(cells[c] for c in sorted(cells))+'</row>'


def standardize_statement(sheet, typ: str, style_ids: dict[str,int]) -> int:
    """Reorder EXISTING statement data, without changing its stored financial facts.

    Only data rows are rearranged; values are copied intact, existing source
    formulas are row-remapped, and the QoQ/YoY formulas are rebuilt.
    """
    # Materialize SEC-updated values and any inserted future columns first.
    sheet.original = sheet.output()
    sheet.pending.clear();sheet.values.clear();sheet.existing.clear();sheet.styles.clear();sheet._read()
    sheet.original = _remap_sheet_xml(sheet.original,lambda col:col)
    sheet.values.clear();sheet.existing.clear();sheet.styles.clear();sheet._read()
    items,has_ca,has_cl = _classify_rows(sheet,typ)
    ordered = _section_sort(typ,items,has_ca,has_cl)
    # Retain numeric or formula-only rows even without a label (append at review).
    known_rows={r for r,_,_,_ in ordered}
    root=ET.fromstring(sheet.original)
    for r in root.findall('.//{%s}sheetData/{%s}row' % (M,M)):
        index=int(r.get('r','0'))
        if index <=3 or index in known_rows:continue
        if any(c.get('r','').split('!')[-1][:-len(str(index))] not in ('A','B') and
               (c.find('{%s}v' % M) is not None or c.find('{%s}f' % M) is not None)
               for c in r.findall('{%s}c' % M)):
            ordered.append((index,f'UNLABELED ORIGINAL ROW {index}','review','detail'))
    # Blank templates still show a clean scaffold, so a touch-created XLSX
    # becomes immediately usable without waiting for EDGAR's next filing.
    if not ordered:ordered=[]
    last_col=max([4,*(c+1 for c in sheet.periods().values())])
    value_columns=set(sheet.periods().values())
    if not value_columns:value_columns=set(range(3,last_col+1,2))
    layout=[];seen=set();previous_group=None;row=4;row_map={}
    for old,label,group,kind in ordered:
        if group!=previous_group:
            row += 1
            for heading in _section_headings(typ,group,seen,has_ca,has_cl):
                layout.append((row,'heading',heading,None,None))
                seen.add(heading);row+=1
        row_map[old]=row
        layout.append((row,'data',label,old,'review' if group=='review' else kind))
        row +=1
        previous_group=group
    # If input rows were already styled and sorted, equal XML is a true no-op.
    source=sheet.original
    sd=re.search(r'<sheetData\b[^>]*>(.*?)</sheetData>',source,re.S)
    if not sd:raise ValueError(f'{sheet.name}: sheetData element is missing')
    existing_by_row={int(m[1]):m[0] for m in ROW_RE.finditer(sd[1])}
    old_top=[xml for r,xml in sorted(existing_by_row.items()) if r<=3]
    out=[]
    for r,what,label,old,kind in layout:
        if what=='heading':out.append(_new_heading_row(r,label,last_col,style_ids))
        else:
            orig=existing_by_row.get(old)
            if orig is None:raise ValueError(f'Missing original data row {old}')
            out.append(_restyle_data_row(orig,old,r,kind,label,last_col,value_columns,
                                         style_ids,row_map))
    # Set row-3 caption and style, preserving its existing period labels.
    if old_top:
        fixed=[]
        for orig in old_top:
            if re.search(r'<row\b[^>]*\br="3"',orig):
                orig=re.sub(r'<c\b[^>]*\br="B3"[^>]*(?:/>|>.*?</c>)',
                    f'<c r="B3" s="{style_ids["header_label"]}" t="inlineStr"><is><t>Line Item</t></is></c>',orig,count=1,flags=re.S)
                if 'r="B3"' not in orig:orig=orig.replace('</row>',f'<c r="B3" s="{style_ids["header_label"]}" t="inlineStr"><is><t>Line Item</t></is></c></row>')
                def style_header(m):
                    c=col_number(m[1])
                    index = (style_ids['header_label'] if c==2 else style_ids['header_number'])
                    cell = re.sub(r'\s+s="[^"]+"','',m[0],count=1)
                    return cell.replace('<c ',f'<c s="{index}" ',1)
                orig=CELL_RE.sub(style_header,orig)
            fixed.append(orig)
        old_top=fixed
    body=''.join(old_top+out)
    result=source[:sd.start(1)]+body+source[sd.end(1):]
    # Reset body-range coordinate metadata that would point to old row positions.
    result=re.sub(r'(<dimension\b[^>]*\bref=")[^"]+("[\s/>])',
                  lambda m:m[1]+f'B1:{col_name(last_col)}{max(row-1,3)}'+m[2],result,count=1)
    # Superseded row-spanning template merges must not cover new section bands.
    result = re.sub(r'<mergeCell\b[^>]*\bref="[A-Z]+(\d+):[A-Z]+(\d+)"[^>]*/>',
                    lambda m: m[0] if int(m[1]) <= 3 and int(m[2]) <= 3 else '',result)
    result = re.sub(r'<mergeCells\b[^>]*>\s*</mergeCells>', '',result)
    if '<mergeCells' in result:
        number = len(re.findall(r'<mergeCell\b',result))
        result = re.sub(r'(<mergeCells\b[^>]*\bcount=")[0-9]+(")',
                        lambda m:m[1]+str(number)+m[2],result,count=1)
    # Workbook drawings can be harmless empty spreadsheet artifacts. Keep their
    # relationships intact; other cell-range-linked objects require a dedicated
    # reanchoring implementation, so refuse rather than silently damage them.
    if any(token in result for token in ('<tableParts ', '<hyperlinks>', '<dataValidations')):
        raise ValueError(f'{sheet.name}: tables, hyperlinks or validations require manual row-migration handling')
    ET.fromstring(result)
    different = result!=sheet.original
    if different:
        sheet.original=result
        sheet.values.clear();sheet.existing.clear();sheet.styles.clear();sheet._read()
        sheet.layout_changed=True
    sheet.report_styles=style_ids
    sheet.row_roles={new_row: ('review' if kind=='review' else kind)
                     for new_row,what,_,_,kind in layout if what=='data'}
    return len(ordered) if different else 0


def sheet_type(name: str) -> str | None:
    s = norm(name)
    if "balance" in s or "financial position" in s: return "balance"
    if "cashflow" in s or "cash flow" in s: return "cashflow"
    if "operation" in s or "income" in s or "earnings" in s: return "income"
    return None


def workbook_scale(workbook: Workbook) -> dict[str, Decimal]:
    """Fallback divisor inferred from historical worksheet labels; override via JSON."""
    hints = []
    for sheet in workbook.sheets.values():
        for pos in [(3, 2), (2, 2), (1, 2)]:
            value = sheet.get(*pos)
            if value: hints.append(str(value).lower())
    combined = " ".join(hints)
    if "million" in combined: scale = Decimal(1_000_000)
    elif "thousand" in combined or "'000" in combined or " 000s" in combined: scale = Decimal(1_000)
    else: scale = Decimal(1)   # Brand-new blank templates use native SEC XBRL units.
    return {name: scale for name in workbook.sheets}


def load_xlsx_units(directory: Path, ticker: str, path: Path, sheets: dict[str, XMLSheet]) -> dict[str, Decimal]:
    """Non-private, explicit numeric scale in historical XLSX rows.

    XBRL monetary facts are raw USD. HHH/HOOD examples contain thousands/millions.
    Visible B3 unit labels may be deleted without changing interpretation once
    explicit unit divisors are configured in edgar_xlsx_units.json.
    """
    units_file = directory / 'edgar_xlsx_units.json'
    config = json.loads(units_file.read_text(encoding='utf-8')) if units_file.exists() else {}
    if not isinstance(config, dict): raise ValueError('edgar_xlsx_units.json must be an object')
    chosen = config.get(path.name, config.get(ticker))
    guessed = workbook_scale(type('WorkbookScale', (), {'sheets': sheets})())
    def divisor(value):
        synonyms = {'dollars': '1', 'thousands': '1000', 'millions': '1000000'}
        n = Decimal(synonyms.get(str(value).lower(), str(value)))
        if n <= 0: raise ValueError(f'Invalid unit divisor: {value}')
        return n
    if chosen is None: return guessed
    if isinstance(chosen, dict):
        return {name: divisor(chosen.get(name, guessed.get(name, Decimal(1)))) for name in sheets}
    return {name: divisor(chosen) for name in sheets}



def remember_xlsx_units(directory: Path, ticker: str, path: Path,
                        units: dict[str, Decimal], *, dry_run=False) -> None:
    """Persist inferred legacy scaling before removing source unit captions.

    Do not overwrite any user-provided mapping. The config is ordinary public
    metadata, unlike the separate private SEC user_agent.conf.
    """
    if dry_run:return
    config_file=directory/'edgar_xlsx_units.json'
    config=json.loads(config_file.read_text(encoding='utf-8')) if config_file.exists() else {}
    if not isinstance(config,dict):raise ValueError('edgar_xlsx_units.json must be a JSON object')
    if path.name in config or ticker in config:return
    unique=set(units.values())
    config[path.name]=(str(next(iter(unique))) if len(unique)==1 else
                      {name:str(divisor) for name,divisor in units.items()})
    write_json_atomic(config_file,config)
    LOG.info('%s: persisted numeric unit scaling in edgar_xlsx_units.json',path.name)


def period_sort_key(period: str) -> tuple[int, int]:
    match = PERIOD_RE.fullmatch(period)
    if not match: raise ValueError(f"Invalid reporting period label: {period}")
    return int(match[1]), 5 if match[2] == "FY" else int(match[2][-1])


def write_json_atomic(path: Path, data: dict) -> None:
    fd, temp = tempfile.mkstemp(suffix=".json", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2, sort_keys=True)
            handle.write("\n")
        os.replace(temp, path)
    finally:
        if os.path.exists(temp): os.unlink(temp)


def adjust_sign(value: Decimal, label: str, override) -> Decimal:
    """Match explicit spreadsheet presentation signs; otherwise keep XBRL signs.

    A positive XBRL 'interest expense' or 'accumulated depreciation' may be
    presented as a negative line in the user's spreadsheet.
    """
    if isinstance(override, dict) and "sign" in override:
        sign = int(override["sign"])
        if sign not in (-1, 1): raise ValueError("Override sign must be -1 or 1")
        return abs(value) * sign
    normalized = norm(label)
    if normalized.startswith("less accumulated depreciation") or normalized in (
        "interest expense", "treasury stock at cost"
    ):
        return -abs(value)
    return value


def unit_for(label: str) -> str:
    lower = norm(label)
    if "per share" in lower or "earnings per share" in lower: return "USD/shares"
    if "weighted average" in lower and "share" in lower: return "shares"
    if "shares outstanding" in lower: return "shares"
    return "USD"


def match_concept(label: str, typ: str, facts, labels, presentation, overrides) -> tuple[str | None, str]:
    key_label = norm(label)
    if key_label in overrides:
        explicit = overrides[key_label]
        key = explicit.get("tag") if isinstance(explicit, dict) else explicit
        return (key, "manual") if key in facts else (None, "manual_not_found")
    allowed_unit = unit_for(label)
    exact = [key for key, variants in labels.items()
             if key in facts and any(rec.unit == allowed_unit for rec in facts[key])
             and any(norm(variant) == key_label for variant in variants)]
    preferred_exact = [key for key in exact if key in presentation.get(typ, set())]
    exact = preferred_exact or exact
    if len(exact) == 1: return exact[0], "exact_label"
    if len(exact) > 1: return None, "ambiguous_exact:" + ",".join(exact[:4])
    aliases = [key for key in ALIASES.get(key_label, []) if key in facts
               and any(rec.unit == allowed_unit for rec in facts[key])]
    preferred_aliases = [key for key in aliases if key in presentation.get(typ, set())]
    aliases = preferred_aliases or aliases
    if len(aliases) == 1: return aliases[0], "alias"
    if len(aliases) > 1: return None, "ambiguous_alias:" + ",".join(aliases[:4])
    # Only use concepts with actual facts of the expected unit and match nearby names.
    candidates = []
    for key, records in facts.items():
        if not any(r.unit == allowed_unit for r in records): continue
        variants = labels.get(key, [])
        if not variants: continue
        similarity = max((difflib.SequenceMatcher(None, key_label, norm(other)).ratio() for other in variants), default=0.0)
        if similarity < 0.88: continue
        in_statement = key in presentation.get(typ, set())
        score = similarity + (0.06 if in_statement else 0.0)
        candidates.append((score, similarity, key))
    candidates.sort(reverse=True)
    if not candidates: return None, "no_match"
    if len(candidates) > 1 and candidates[0][0] - candidates[1][0] < 0.035:
        return None, "ambiguous:" + ",".join(c[-1] for c in candidates[:3])
    if candidates[0][1] < 0.94 and candidates[0][2] not in presentation.get(typ, set()):
        return None, "low_confidence"
    return candidates[0][2], "matched_label"


def period_date(label: str) -> dt.date:
    match = PERIOD_RE.fullmatch(label)
    if not match: raise ValueError(label)
    year = int(match.group(1))
    if match.group(2) == "FY": return dt.date(year, 12, 31)
    q = int(match.group(2)[1])
    month = 3*q
    return dt.date(year, month+1, 1) - dt.timedelta(days=1) if month < 12 else dt.date(year, 12, 31)


def same_period_end(a: str, target: dt.date) -> bool:
    try: d = dt.date.fromisoformat(a)
    except ValueError: return False
    return abs((d-target).days) <= 10


def choose_fact(records: list[Fact], end: dt.date, unit: str, mode: str) -> tuple[Decimal, Fact] | None:
    """Prefer a fact reported in its own filing, then the most recent re-reporting."""
    valid = []
    for fact in records:
        if fact.unit != unit or not same_period_end(fact.end, end): continue
        if mode == "instant":
            if fact.start is not None: continue
        else:
            if fact.start is None: continue
            try: days = (dt.date.fromisoformat(fact.end) - dt.date.fromisoformat(fact.start)).days + 1
            except ValueError: continue
            if mode == "quarter" and not 65 <= days <= 115: continue
            if mode == "annual" and not 345 <= days <= 380: continue
            if mode == "ytd":
                expected = (end - dt.date(end.year, 1, 1)).days + 1
                if abs(days - expected) > 18: continue
                if not 65 <= days <= 380: continue
        valid.append(fact)
    if not valid: return None
    # Restrict to the best (typically own-period) filing: avoid combining multiple restatements.
    def rank(f: Fact):
        end_delta = abs((dt.date.fromisoformat(f.end) - end).days)
        return (end_delta, abs((dt.date.fromisoformat(f.filed or f.end) - end).days), 0 if f.source == "filing_instance" else 1)
    valid.sort(key=rank)
    best_rank = rank(valid[0])
    top = [f for f in valid if rank(f) == best_rank]
    if len({f.value for f in top}) > 1:
        return None  # same concept, date & duration but inconsistent numerical disclosure
    return top[0].value, top[0]


def value_for(key: str, records: list[Fact], period: str, typ: str, unit: str):
    """Return value, source fact(s), derivation. No EPS/share-count subtraction."""
    end = period_date(period)
    if typ == "balance":
        hit = choose_fact(records, end, unit, "instant")
        return (hit[0], [hit[1]], "as_of") if hit else None
    is_fy = period.endswith("FY")
    if is_fy:
        hit = choose_fact(records, end, unit, "annual")
        return (hit[0], [hit[1]], "annual") if hit else None
    direct = choose_fact(records, end, unit, "quarter")
    if direct: return direct[0], [direct[1]], "direct_quarter"
    if unit in ("shares", "USD/shares"):
        return None  # EPS and average shares should not be simply subtracted across YTD periods.
    q = int(period[-1])
    if q == 1:
        ytd = choose_fact(records, end, unit, "ytd")
        return (ytd[0], [ytd[1]], "q1_ytd") if ytd else None
    if q == 4:
        annual = choose_fact(records, end, unit, "annual")
        previous = choose_fact(records, dt.date(end.year, 9, 30), unit, "ytd")
    else:
        annual = choose_fact(records, end, unit, "ytd")
        previous = choose_fact(records, dt.date(end.year, (q-1)*3+1, 1) - dt.timedelta(days=1), unit, "ytd")
    if annual and previous:
        # For Q2/Q3 use same annual year, and Q4 FY minus first nine months.
        if annual[1].start and previous[1].start and abs((dt.date.fromisoformat(annual[1].start) - dt.date.fromisoformat(previous[1].start)).days) > 18:
            return None
        return annual[0] - previous[0], [annual[1], previous[1]], "derived_from_ytd"
    return None


def period_source(facts: list[Fact], filings: list[Filing]) -> str:
    matched = {fact.accn for fact in facts}
    return "; ".join(f.source_url for f in filings if f.accn in matched)


def find_candidate_periods(filings: list[Filing]):
    periods = set()
    for f in filings:
        year, q = quarter(f.report_date)
        if f.form == "10-K":
            periods.add(f"{year}FY"); periods.add(f"{year}Q4")
        else: periods.add(f"{year}Q{q}")
    return sorted(periods, key=lambda p: (int(p[:4]), 5 if p.endswith("FY") else int(p[-1])))


def append_log(path: Path, changes: list[dict], dry_run: bool):
    if dry_run or not changes: return
    columns = ["run_time", "filename", "sheet", "period", "row", "line_item", "xbrl_concept", "old", "new", "method", "source", "action"]
    existing = path.exists()
    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        if not existing: writer.writeheader()
        writer.writerows(changes)


def save_review(path: Path, rows: list[dict], dry_run: bool):
    if dry_run: return
    with path.open("w", newline="", encoding="utf-8") as handle:
        fields = ["filename", "sheet", "line_item", "problem", "suggested_concept", "filing_url"]
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader(); writer.writerows(rows)


def load_overrides(path: Path, ticker: str) -> dict[str, str]:
    if not path.exists(): return {}
    config = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(config, dict): raise ValueError("Mapping file must be a JSON object")
    company = config.get(ticker, {})
    if not isinstance(company, dict): raise ValueError(f"Mapping for {ticker} must be an object")
    for label, entry in company.items():
        if isinstance(entry, dict) and (not isinstance(entry.get("tag"), str) or entry.get("sign", 1) not in (1, -1)):
            raise ValueError(f"Invalid mapping for {ticker} / {label}; use {{'tag':'ns:Name','sign':-1}}")
        if not isinstance(entry, (dict, str)):
            raise ValueError(f"Invalid mapping for {ticker} / {label}")
    return {norm(k): v for k, v in company.items()}


def update_workbook(path: Path, ticker: str, filings: list[Filing], facts, labels, roles, args):
    if path.suffix.lower() != '.xlsx':
        raise ValueError(f'Only multi-sheet .xlsx workbooks are supported: {path.name}')
    book = Workbook(path)
    units = load_xlsx_units(Path(args.directory), ticker, path, book.sheets)
    overrides = load_overrides(Path(args.directory) / args.mappings_file, ticker)
    periods = find_candidate_periods(filings)
    # Migrate existing C:R period headers and cells to alternating value/change
    # columns exactly once, before adding new EDGAR periods and numeric facts.
    for sheet in book.sheets.values():
        sheet.enable_change_columns()
    reference_periods = max((s.periods() for s in book.sheets.values()), key=len, default={})
    for sheet in book.sheets.values():
        sheet.seed_periods(reference_periods)
        sheet.ensure_change_headers()
    now = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    changes = []
    reviews = []
    for name, sheet in book.sheets.items():
        typ = sheet_type(name)
        if not typ: continue
        scale = units[name]
        if not scale:
            LOG.warning("Skip %s / %s: units not found (pass --default-units)", path.name, name)
            continue
        # Existing historical/future periods stay intact; insert value/change pairs chronologically.
        for period in periods:
            sheet.add_period(period)
        # Do not put new item candidates from arbitrary XBRL note tables into income/balance sheet.
        existing_rows = sheet.rows()
        present_keys = set(roles.get(typ, set()))
        matched = set()
        for row, label in existing_rows.items():
            label_norm = norm(label)
            if label_norm in SKIP_LABELS or any(token in label_norm for token in CHECK_TOKENS): continue
            if "additional sec line items" in label_norm: continue
            key, method = match_concept(label, typ, facts, labels, roles, overrides)
            if not key:
                if row <= 4 or not any(ch.isalpha() for ch in label): continue
                reviews.append({"filename": path.name, "sheet": name, "line_item": label,
                                "problem": method, "suggested_concept": "", "filing_url": filings[0].source_url if filings else ""})
                continue
            matched.add(key)
            for period, col in sheet.periods().items():
                if period not in periods or sheet.has_value(row, col): continue
                selected = value_for(key, facts[key], period, typ, unit_for(label))
                if selected is None: continue
                raw_value, selected_facts, how = selected
                raw_value = adjust_sign(raw_value, label, overrides.get(norm(label)))
                divisor = Decimal(1) if unit_for(label) == "USD/shares" else (Decimal(1_000_000) if unit_for(label) == "shares" and "million" in norm(label) else (Decimal(1) if unit_for(label) == "shares" else scale))
                value = raw_value / divisor
                if sheet.put(row, col, value):
                    old = sheet.values.get((row, col))
                    changes.append({"run_time": now, "filename": path.name, "sheet": name, "period": period,
                                    "row": row, "line_item": label, "xbrl_concept": key, "old": str(old) if old is not None else "", "new": str(value),
                                    "method": how, "source": period_source(selected_facts, filings), "action": "filled"})
        # New rows: only components actually listed in the filing presentation of this statement,
        # and only after the user explicitly enabled discovery.
        if args.discover_rows:
            already = {norm(text) for text in sheet.rows().values()}
            discovered = 0
            for key in sorted(present_keys):
                if key in matched or key not in facts: continue
                proposed_labels = labels.get(key, [])
                if not proposed_labels: continue
                label = proposed_labels[0]
                normalized = norm(label)
                if normalized in already or normalized in SKIP_LABELS or "abstract" in key.lower(): continue
                if len(label) > 95 or not label.strip() or not any(c.isalpha() for c in label): continue
                if "[" in label or "]" in label or "axis" in normalized or "domain" in normalized: continue
                fact_units = {f.unit for f in facts[key]}
                unit = "USD/shares" if "USD/shares" in fact_units and "USD" not in fact_units else (
                       "shares" if "shares" in fact_units and "USD" not in fact_units else "USD")
                values = {}
                for period in periods:
                    found = value_for(key, facts[key], period, typ, unit)
                    if found:
                        raw_value, source_facts, how = found
                        raw_value = adjust_sign(raw_value, label, overrides.get(norm(label)))
                        divisor = (Decimal(1) if unit == "USD/shares" else
                                   (Decimal(1_000_000) if unit == "shares" and "million" in normalized else
                                    (Decimal(1) if unit == "shares" else scale)))
                        values[period] = (raw_value / divisor, source_facts, how)
                if not values: continue
                if args.max_new_rows and discovered >= args.max_new_rows:
                    reviews.append({"filename": path.name, "sheet": name, "line_item": label,
                                    "problem": "max_new_rows_reached", "suggested_concept": key,
                                    "filing_url": filings[0].source_url if filings else ""})
                    continue
                row = sheet.new_item_row(label)
                sheet.put(row, 2, label, "s")
                already.add(normalized); discovered += 1
                for period, (value, used, how) in values.items():
                    col = sheet.periods()[period]
                    if sheet.put(row, col, value):
                        changes.append({"run_time": now, "filename": path.name, "sheet": name, "period": period,
                                        "row": row, "line_item": label, "xbrl_concept": key, "old": "", "new": str(value),
                                        "method": "new_review_row:" + how, "source": period_source(used, filings), "action": "new_row"})
        if typ == "balance":
            # Validation: do not silently accept a non-balancing consolidated BS.
            normalized_rows = {norm(text): row for row, text in sheet.rows().items()}
            ar = normalized_rows.get("total assets")
            lr = (normalized_rows.get("total liabilities and stockholders equity") or
                  normalized_rows.get("total liabilities mezzanine equity and equity"))
            if ar and lr:
                for period in periods:
                    col = sheet.periods().get(period)
                    if col is None: continue
                    av, lv = sheet.get(ar, col), sheet.get(lr, col)
                    if isinstance(av, Decimal) and isinstance(lv, Decimal) and abs(av - lv) > Decimal(2):
                        reviews.append({"filename": path.name, "sheet": name,
                                        "line_item": f"{period}: Total assets - total liabilities and equity = {av-lv}",
                                        "problem": "balance_sheet_does_not_reconcile", "suggested_concept": "",
                                        "filing_url": filings[0].source_url if filings else ""})
    for sheet in book.sheets.values():
        typ = sheet_type(sheet.name)
        if typ:
            sheet.ensure_change_headers()
            standardize_statement(sheet, typ, book.report_styles)
            sheet.ensure_percent_formulas()
    book.save(dry_run=args.dry_run, backup=not args.no_backup)
    remember_xlsx_units(Path(args.directory), ticker, path, units, dry_run=args.dry_run)
    return changes, reviews


def tickers_in_directory(directory: Path):
    selected = []
    for path in sorted(directory.iterdir()):
        if not path.is_file() or path.suffix.lower() != ".xlsx": continue
        parsed = parse_filename(path)
        if not parsed: continue
        selected.append((path, *parsed))
    return selected


def format_only(args):
    """Offline migration to paired value/percentage columns. No SEC access."""
    directory = Path(args.directory).expanduser().resolve()
    if not directory.is_dir(): raise ValueError(f'Not a directory: {directory}')
    count = 0
    for path, ticker, company, owner, fmt in tickers_in_directory(directory):
        if args.owner and owner != args.owner: continue
        try:
            wb = Workbook(path)
            units = load_xlsx_units(directory, ticker, path, wb.sheets)
            if wb.is_placeholder:
                LOG.info("%s: initializing blank 3-sheet workbook%s", path.name,
                         " (dry-run only)" if args.dry_run else "")
            for sheet in wb.sheets.values():
                sheet.enable_change_columns()
            reference = max((sh.periods() for sh in wb.sheets.values()), key=len, default={})
            for sheet in wb.sheets.values():
                sheet.seed_periods(reference)
                sheet.ensure_change_headers()
                typ = sheet_type(sheet.name)
                if typ:
                    standardize_statement(sheet, typ, wb.report_styles)
                sheet.ensure_percent_formulas()
            changed = wb.changed() if args.dry_run else False
            wrote = wb.save(dry_run=args.dry_run, backup=not args.no_backup)
            remember_xlsx_units(directory, ticker, path, units, dry_run=args.dry_run)
            if changed or wrote:
                count += 1
                LOG.info('%s: %s standard account order, formatting and percentage columns', path.name,
                         'would apply' if args.dry_run else 'applied')
        except Exception:
            LOG.exception('Failed formatting %s; other workbooks continue', path.name)
    LOG.info('Offline format-only: %d workbooks %s', count,
             'proposed' if args.dry_run else 'updated')


def run_once(args):
    directory = Path(args.directory).expanduser().resolve()
    if not directory.is_dir(): raise ValueError(f"Not a directory: {directory}")
    files = [(path, ticker, company, owner, fmt)
             for path, ticker, company, owner, fmt in tickers_in_directory(directory)
             if fmt == "xlsx"]
    if not files:
        LOG.warning("No company XLSX workbooks in %s. Expected TICKER_COMPANY TITLE_GTL.xlsx", directory)
        return
    valid_files = []
    for entry in files:
        try:
            validate_xlsx(entry[0], allow_empty=True)
            if entry[0].stat().st_size == 0:
                LOG.info("%s: empty placeholder; will initialize 3-sheet XLSX on first write%s",
                         entry[0].name, " (dry-run: no write)" if args.dry_run else "")
            valid_files.append(entry)
        except ValueError as exc:
            LOG.error("Skipping invalid workbook: %s", exc)
    if not valid_files:
        LOG.error("No valid .xlsx workbooks to update; fix the rejected files and retry")
        return
    files = valid_files
    client = SECClient(args.user_agent, directory / args.cache_dir, args.request_spacing)
    review_rows = []
    all_changes = []
    by_ticker = defaultdict(list)
    for path, ticker, company, owner, _ in files:
        if args.owner and args.owner != owner: continue
        by_ticker[ticker].append(path)
    state_file = directory / ".edgar_seen.json"
    try:
        state = json.loads(state_file.read_text(encoding="utf-8")) if state_file.exists() else {}
        if not isinstance(state, dict): state = {}
    except (OSError, ValueError):
        state = {}
    for ticker, paths in by_ticker.items():
        try:
            cik = client.cik_for(ticker)
            submissions = client.submissions(cik)
            if str(submissions.get("fiscalYearEnd", "1231")) != "1231":
                LOG.warning("%s fiscal year end %s: only calendar-year spreads supported; skipping", ticker, submissions.get("fiscalYearEnd"))
                continue
            filings = recent_filings(submissions, cik, args.backfill_filings)
            if not filings:
                LOG.warning("No recent 10-Q / 10-K filings: %s", ticker)
                continue
            current_accns = [f.accn for f in filings]
            ticker_state = state.get(ticker, {})
            recorded = ticker_state.get("files", {})
            if not isinstance(recorded, dict): recorded = {}
            pending_files = []
            for path in paths:
                previous = recorded.get(path.name, {})
                same_filings = previous.get("accessions", []) == current_accns
                same_settings = (bool(previous.get("discovered_rows")) == bool(args.discover_rows)
                                 and previous.get("layout_version") == REPORT_LAYOUT_VERSION)
                mapping_file = directory / args.mappings_file
                mapping_mtime = mapping_file.stat().st_mtime if mapping_file.exists() else None
                same_mapping = previous.get("mapping_mtime") == mapping_mtime
                if not args.recheck_existing and same_filings and same_settings and same_mapping:
                    LOG.info("%s: no newly detected 10-Q / 10-K (unchanged since previous successful scan)", path.name)
                    continue
                pending_files.append((path, mapping_mtime))
            if not pending_files: continue
            added = [f for f in filings if f.accn not in set(next(iter(recorded.values())).get("accessions", []))] if recorded else filings
            LOG.info("%s: %d newly detected or not-yet-processed filings", ticker, len(added))
            # Combine SEC standardized companyfacts and filing-specific facts for custom disclosures.
            data = client.companyfacts(cik)
            facts, labels = parse_companyfacts(data, {f.accn for f in filings})
            roles = {"balance": set(), "income": set(), "cashflow": set()}
            for index, filing in enumerate(filings):
                extra, extra_labels, extra_roles = gather_filing(client, filing)
                merge_dict_lists(facts, extra)
                merge_dict_lists(labels, extra_labels)
                # Union all inspected filings, not only the newest: a disclosure
                # may disappear from later statements but should remain in history.
                for typ in roles: roles[typ].update(extra_roles[typ])
            LOG.info("%s CIK %s: %d filings, %d concepts, most recent: %s %s", ticker, cik, len(filings), len(facts), filings[0].form, filings[0].report_date)
            newest_facts_ready = any(f.accn == filings[0].accn for rows in facts.values() for f in rows)
            newest_statement_roles_ready = any(roles.values())
            can_mark_processed = newest_facts_ready and (not args.discover_rows or newest_statement_roles_ready)
            if not can_mark_processed:
                LOG.warning("%s: latest SEC filing XBRL/presentation is not fully available yet; will retry on next check", ticker)
            for path, mapping_mtime in pending_files:
                changes, review = update_workbook(path, ticker, filings, facts, labels, roles, args)
                all_changes.extend(changes)
                review_rows.extend(review)
                LOG.info("%s: %d new values / rows; %d unresolved line items", path.name, len(changes), len(review))
                if args.dry_run:
                    for item in changes[:30]:
                        LOG.info("[DRY RUN] %s / %s / %s / %s = %s [%s]",
                                 item["filename"], item["sheet"], item["period"],
                                 item["line_item"], item["new"], item["xbrl_concept"])
                    if len(changes) > 30:
                        LOG.info("[DRY RUN] %d additional proposed changes not printed", len(changes)-30)
                    for item in review[:8]:
                        LOG.info("[REVIEW] %s: %s (%s)", item["sheet"], item["line_item"], item["problem"])
                if not args.dry_run and can_mark_processed:
                    recorded[path.name] = {"accessions": current_accns, "discovered_rows": bool(args.discover_rows),
                                           "mapping_mtime": mapping_mtime, "layout_version": REPORT_LAYOUT_VERSION,
                                           "processed_utc": dt.datetime.now(dt.timezone.utc).isoformat()}
                    state[ticker] = {"files": recorded}
        except Exception:
            LOG.exception("Failed processing ticker %s; other companies continue", ticker)
    if not args.dry_run:
        write_json_atomic(state_file, state)
    append_log(directory / args.log_file, all_changes, args.dry_run)
    save_review(directory / args.review_file, review_rows, args.dry_run)
    LOG.info("Completed: %d values, %d review entries; dry_run=%s", len(all_changes), len(review_rows), args.dry_run)


def load_sec_user_agent(path: Path) -> str:
    """Load a User-Agent from a local file without echoing its contents.

    Accepted formats:
        SEC_USER_AGENT=GTL Research contact@example.com
        GTL Research contact@example.com

    Comment and empty lines are ignored; one setting only is permitted.
    Missing files return an empty string so environment fallback can work.
    """
    try:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    except FileNotFoundError:
        return ""
    except OSError as exc:
        raise ValueError(f"Unable to read User-Agent config {path}: {exc.strerror}") from exc

    values = []
    for line_no, line in enumerate(lines, start=1):
        entry = line.strip()
        if not entry or entry.startswith(("#", ";")):
            continue
        if entry.startswith("export "):
            entry = entry[7:].strip()
        if "=" in entry:
            key, entry = entry.split("=", 1)
            if key.strip() != "SEC_USER_AGENT":
                raise ValueError(f"Unexpected setting in {path} on line {line_no}; expected SEC_USER_AGENT")
            entry = entry.strip()
        if len(entry) >= 2 and entry[0] == entry[-1] and entry[0] in ("'", '"'):
            entry = entry[1:-1]
        if not entry.strip():
            raise ValueError(f"Empty SEC_USER_AGENT in {path} on line {line_no}")
        values.append(entry.strip())
    if len(values) > 1:
        raise ValueError(f"Multiple User-Agent entries in {path}; use only one")
    return values[0] if values else ""


def main(argv=None):
    parser = argparse.ArgumentParser(description="Monitor SEC EDGAR and standardize three-sheet XLSX spreads with a consistent account hierarchy and QoQ / YoY formulas.")
    parser.add_argument("--directory", default=".", help="Folder with TICKER_COMPANY TITLE_GTL.xlsx workbooks")
    parser.add_argument("--user-agent", default=None, help="Override User-Agent from config file / environment (avoid on shared servers: command lines may be visible)")
    parser.add_argument("--user-agent-file", default=None, help="Path to private config; defaults to user_agent.conf beside this script")
    parser.add_argument("--interval-hours", type=float, default=0, help="0 = one run; positive = repeat forever at this interval")
    parser.add_argument("--backfill-filings", type=int, default=8, help="Number of most recent 10-K / 10-Q filings per ticker (default 8)")
    parser.add_argument("--request-spacing", type=float, default=0.5, help="Seconds between SEC HTTP requests; default 0.5 (2 requests/sec)")
    parser.add_argument("--owner", default="GTL", help="Only update workbooks with this owner suffix; empty string = all owners")
    parser.add_argument("--discover-rows", action="store_true", default=True, help="Add new main-statement XBRL fields (enabled by default)")
    parser.add_argument("--no-discover-rows", action="store_false", dest="discover_rows", help="Disable discovery of new statement rows")
    parser.add_argument("--max-new-rows", type=int, default=0, help="Max discovered rows per statement, 0 (default) = unlimited")
    parser.add_argument("--format-only", action="store_true", help="Offline migration: organize existing statement rows and add percentage columns, without contacting EDGAR")
    parser.add_argument("--dry-run", action="store_true", help="Analyze but do not modify workbooks or logs")
    parser.add_argument("--no-backup", action="store_true", help="Disable timestamped original workbook backups")
    parser.add_argument("--recheck-existing", action="store_true", help="Reprocess even when no new SEC filing is detected")
    parser.add_argument("--cache-dir", default=".edgar_cache")
    parser.add_argument("--mappings-file", default="edgar_mappings.json")
    parser.add_argument("--log-file", default="edgar_changes.csv")
    parser.add_argument("--review-file", default="edgar_review.csv")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if args.format_only:
        format_only(args)
        return
    # Resolve relative to the SCRIPT, not the working directory. This ensures
    # scheduled jobs read the private config regardless of WorkingDirectory.
    config_path = (Path(args.user_agent_file).expanduser() if args.user_agent_file
                   else Path(__file__).resolve().parent / "user_agent.conf")
    try:
        file_agent = load_sec_user_agent(config_path) if args.user_agent is None else ""
    except ValueError as exc:
        parser.error(str(exc))
    # Explicit CLI override > private .conf > environment variable.
    args.user_agent = (args.user_agent if args.user_agent is not None else
                       (file_agent or os.getenv("SEC_USER_AGENT", ""))).strip()
    if not args.user_agent or "@" not in args.user_agent:
        parser.error("Add 'SEC_USER_AGENT=Organization contact@example.com' to user_agent.conf beside the script, "
                     "or set SEC_USER_AGENT in the environment")
    if args.backfill_filings < 1: parser.error("--backfill-filings must be at least 1")
    if args.interval_hours < 0: parser.error("--interval-hours must be >= 0")
    if args.max_new_rows < 0: parser.error("--max-new-rows must be >= 0")
    while True:
        run_once(args)
        if not args.interval_hours: return
        LOG.info("Next check in %.2f hours", args.interval_hours)
        time.sleep(args.interval_hours*3600)


if __name__ == "__main__":
    main()
