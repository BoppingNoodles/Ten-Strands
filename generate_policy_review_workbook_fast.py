"""
Fast policy review workbook generator for the Caden tab.

This script keeps the original workbook untouched. It checks rows 9-315 and
columns E-BH, proposes policy value/year/link changes in a new workbook,
bolds changed cells, highlights revisions/found policies green, highlights
expired links red, and adds a Summary of Changes column.

Performance notes:
- Row-level web work runs in a ThreadPoolExecutor.
- Worker threads do not touch openpyxl objects.
- The cache is shared through a lock and saved in batches.
- The workbook is edited only on the main thread and saved atomically in batches.
"""

from __future__ import annotations

import argparse
import json
import re
import threading
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from html import unescape
from pathlib import Path
from typing import Iterable

import requests

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing dependency: openpyxl\n"
        "Install it with: python -m pip install openpyxl"
    ) from exc


DEFAULT_INPUT = Path(
    r"C:\Users\caden\Downloads\Summer 2026 Board Policy Indicator Refresh Data Tracker.xlsx"
)
DEFAULT_OUTPUT = Path("caden_policy_review_rows_9_315_fast.xlsx")
DEFAULT_CACHE = Path("policy_review_cache.json")

SHEET_NAME = "Caden"
HEADER_ROW = 2
START_ROW = 9
END_ROW = 315
START_COL = 5  # E
END_COL = 60  # BH
GROUP_WIDTH = 4

TIMEOUT_SECONDS = 8
REQUEST_DELAY_SECONDS = 0.0
SEARCH_MAX_RESULTS = 4
DATABASE_SEARCH_MAX_RESULTS = 3
DEFAULT_WORKERS = 8
DEFAULT_CACHE_SAVE_INTERVAL = 10
DEFAULT_WORKBOOK_SAVE_INTERVAL = 10

GREEN_FILL_NAME = "green"
RED_FILL_NAME = "red"
YELLOW_FILL_NAME = "yellow"

GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
CHANGED_FONT = Font(bold=True)

SEARCH_URL = "https://duckduckgo.com/html/?q={query}"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"
)
POLICY_HOST_MARKERS = [
    "simbli.eboardsolutions.com",
    "go.boarddocs.com",
    "gamutonline.net",
    "boardpolicyonline.com",
    "agendaonline.net",
]

thread_local = threading.local()


@dataclass(frozen=True)
class PolicyGroup:
    value_col: int
    adopted_col: int
    revised_col: int
    link_col: int
    label: str
    code: str
    title: str


@dataclass
class FetchResult:
    ok: bool
    final_url: str
    status: int | None
    text: str
    error: str | None = None


@dataclass(frozen=True)
class RowSnapshot:
    row: int
    district: str
    values: dict[int, str]


@dataclass(frozen=True)
class CellUpdate:
    col: int
    value: str
    fill: str


@dataclass(frozen=True)
class CellMark:
    col: int
    fill: str


@dataclass
class RowResult:
    row: int
    notes: list[str] = field(default_factory=list)
    updates: list[CellUpdate] = field(default_factory=list)
    marks: list[CellMark] = field(default_factory=list)


class ThreadSafeCache:
    def __init__(self, path: Path):
        self.path = path
        self.lock = threading.Lock()
        self.dirty = False
        self.data = self._load()

    def _load(self) -> dict:
        if not self.path.exists():
            return {}
        return json.loads(self.path.read_text(encoding="utf-8"))

    def get(self, key: str):
        with self.lock:
            return self.data.get(key)

    def set(self, key: str, value) -> None:
        with self.lock:
            self.data[key] = value
            self.dirty = True

    def save_if_dirty(self) -> None:
        with self.lock:
            if not self.dirty:
                return
            payload = json.dumps(self.data, indent=2, sort_keys=True)
            self.dirty = False
        tmp_path = self.path.with_name(f"{self.path.stem}.tmp{self.path.suffix}")
        tmp_path.write_text(payload, encoding="utf-8")
        tmp_path.replace(self.path)


def get_session() -> requests.Session:
    session = getattr(thread_local, "session", None)
    if session is None:
        session = requests.Session()
        session.headers.update({"User-Agent": USER_AGENT})
        thread_local.session = session
    return session


def normalize_cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def is_url(value: str) -> bool:
    return value.lower().startswith(("http://", "https://"))


def extract_policy_parts(label: str) -> tuple[str, str]:
    base = re.split(r"\s+[-\u2013]\s+", label, maxsplit=1)[0].strip()
    match = re.match(r"^(BP|AR)\s+([0-9.]+)\s+(.+)$", base, re.IGNORECASE)
    if not match:
        return "", base
    return f"{match.group(1).upper()} {match.group(2)}", match.group(3).strip()


def request_url(url: str, timeout: int) -> FetchResult:
    try:
        response = get_session().get(url, timeout=timeout, allow_redirects=True)
        response.encoding = response.encoding or "utf-8"
        return FetchResult(
            ok=200 <= response.status_code < 400,
            final_url=response.url,
            status=response.status_code,
            text=response.text[:800_000],
        )
    except Exception as exc:
        return FetchResult(False, url, None, "", str(exc))


def cached_fetch(url: str, cache: ThreadSafeCache, timeout: int, delay: float) -> FetchResult:
    key = f"fetch::{url}"
    cached = cache.get(key)
    if cached is not None:
        return FetchResult(**cached)
    result = request_url(url, timeout)
    cache.set(key, result.__dict__)
    if delay:
        time.sleep(delay)
    return result


def page_looks_expired(result: FetchResult) -> bool:
    if not result.ok:
        return True
    text = result.text.lower()
    expired_markers = [
        "404",
        "not found",
        "page cannot be found",
        "policy not found",
        "policy is no longer available",
        "expired",
        "access denied",
        "permission denied",
        "private?open",
    ]
    return any(marker in text for marker in expired_markers)


def visible_text(html: str) -> str:
    html = re.sub(r"(?is)<script.*?</script>|<style.*?</style>", " ", html)
    html = re.sub(r"(?s)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", unescape(html)).strip()


def extract_years(text: str) -> tuple[str | None, str | None]:
    adopted = None
    revised = None
    adopted_patterns = [
        r"\bAdopted\s*:?\s*(?:[^0-9]{0,30})((?:19|20)\d{2})",
        r"\bApproved\s*:?\s*(?:[^0-9]{0,30})((?:19|20)\d{2})",
    ]
    revised_patterns = [
        r"\bRevised\s*:?\s*(?:[^0-9]{0,30})((?:19|20)\d{2})",
        r"\bLast\s+Revised\s*:?\s*(?:[^0-9]{0,30})((?:19|20)\d{2})",
        r"\bUpdated\s*:?\s*(?:[^0-9]{0,30})((?:19|20)\d{2})",
    ]
    for pattern in adopted_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            adopted = match.group(1)
            break
    revised_matches: list[str] = []
    for pattern in revised_patterns:
        revised_matches.extend(re.findall(pattern, text, re.IGNORECASE))
    if revised_matches:
        revised = max(revised_matches)
    return adopted, revised


def html_search(
    query: str,
    cache: ThreadSafeCache,
    timeout: int,
    delay: float,
    max_results: int = SEARCH_MAX_RESULTS,
) -> list[str]:
    key = f"search::{query}"
    cached = cache.get(key)
    if cached is not None:
        return cached

    url = SEARCH_URL.format(query=urllib.parse.quote_plus(query))
    result = request_url(url, timeout)
    links: list[str] = []
    if result.ok:
        for raw in re.findall(r'href="([^"]+)"', result.text):
            raw = unescape(raw)
            parsed = urllib.parse.urlparse(raw)
            if parsed.path == "/l/":
                params = urllib.parse.parse_qs(parsed.query)
                raw = params.get("uddg", [raw])[0]
            if raw.startswith(("http://", "https://")) and "duckduckgo.com" not in raw:
                links.append(raw)

    deduped = []
    seen = set()
    for link in links:
        clean = link.split("&rut=")[0]
        if clean not in seen:
            deduped.append(clean)
            seen.add(clean)
        if len(deduped) >= max_results:
            break

    cache.set(key, deduped)
    if delay:
        time.sleep(delay)
    return deduped


def search_queries(district: str, policy: PolicyGroup) -> Iterable[str]:
    quoted_district = f'"{district}"'
    yield f'{quoted_district} "{policy.code}" "{policy.title}" board policy'
    yield f'{quoted_district} "{policy.code}" site:simbli.eboardsolutions.com'
    yield f'{quoted_district} "{policy.code}" site:go.boarddocs.com'


def result_matches_policy(url: str, text: str, policy: PolicyGroup) -> bool:
    haystack = f"{url} {text}".lower()
    compact = haystack.replace(" ", "")
    code = policy.code.lower()
    code_without_space = code.replace(" ", "")
    return code in haystack or code_without_space in compact


def candidate_is_policy_system(url: str, policy: PolicyGroup) -> bool:
    lowered = url.lower()
    code = policy.code.lower().replace(" ", "")
    compact_url = lowered.replace("%20", "").replace("-", "").replace("_", "")
    return any(marker in lowered for marker in POLICY_HOST_MARKERS) or code in compact_url


def find_policy_link(
    district: str,
    policy: PolicyGroup,
    cache: ThreadSafeCache,
    timeout: int,
    delay: float,
) -> tuple[str | None, str | None, str | None]:
    for query in search_queries(district, policy):
        for candidate in html_search(query, cache, timeout, delay):
            if not candidate_is_policy_system(candidate, policy):
                continue
            fetched = cached_fetch(candidate, cache, timeout, delay)
            if page_looks_expired(fetched):
                continue
            text = visible_text(fetched.text)
            if result_matches_policy(fetched.final_url, text, policy):
                adopted, revised = extract_years(text)
                return fetched.final_url, adopted, revised
    return None, None, None


def district_policy_database_exists(
    district: str,
    cache: ThreadSafeCache,
    timeout: int,
    delay: float,
) -> bool:
    queries = [
        f'"{district}" "board policy"',
        f'"{district}" "simbli"',
        f'"{district}" "BoardDocs"',
    ]
    for query in queries:
        for link in html_search(query, cache, timeout, delay, DATABASE_SEARCH_MAX_RESULTS):
            lowered = link.lower()
            if any(host in lowered for host in POLICY_HOST_MARKERS):
                return True
    return False


def as_int_year(value: str) -> int | None:
    match = re.search(r"(19|20)\d{2}", value)
    if not match:
        return None
    return int(match.group(0))


def add_update(result: RowResult, snapshot: RowSnapshot, col: int, value: str, fill: str) -> bool:
    old = snapshot.values.get(col, "")
    if old == normalize_cell(value):
        return False
    result.updates.append(CellUpdate(col, value, fill))
    return True


def add_mark(result: RowResult, col: int, fill: str) -> None:
    result.marks.append(CellMark(col, fill))


def process_existing_policy(
    result: RowResult,
    snapshot: RowSnapshot,
    policy: PolicyGroup,
    cache: ThreadSafeCache,
    timeout: int,
    delay: float,
) -> None:
    district = snapshot.district
    link = snapshot.values.get(policy.link_col, "")
    old_adopted = snapshot.values.get(policy.adopted_col, "")
    old_revised = snapshot.values.get(policy.revised_col, "")

    if not is_url(link):
        found_link, found_adopted, found_revised = find_policy_link(district, policy, cache, timeout, delay)
        if found_link:
            add_update(result, snapshot, policy.link_col, found_link, GREEN_FILL_NAME)
            if found_adopted:
                add_update(result, snapshot, policy.adopted_col, found_adopted, GREEN_FILL_NAME)
            if found_revised:
                add_update(result, snapshot, policy.revised_col, found_revised, GREEN_FILL_NAME)
            result.notes.append(f"{policy.code}: found replacement link for existing policy")
        return

    fetched = cached_fetch(link, cache, timeout, delay)
    if page_looks_expired(fetched):
        add_mark(result, policy.value_col, RED_FILL_NAME)
        add_mark(result, policy.link_col, RED_FILL_NAME)
        found_link, found_adopted, found_revised = find_policy_link(district, policy, cache, timeout, delay)
        if found_link:
            add_update(result, snapshot, policy.link_col, found_link, RED_FILL_NAME)
            if found_adopted:
                add_update(result, snapshot, policy.adopted_col, found_adopted, RED_FILL_NAME)
            if found_revised:
                add_update(result, snapshot, policy.revised_col, found_revised, RED_FILL_NAME)
            result.notes.append(f"{policy.code}: original link expired; replacement found")
        else:
            replacement = "N/A" if district_policy_database_exists(district, cache, timeout, delay) else "*"
            add_update(result, snapshot, policy.link_col, replacement, RED_FILL_NAME)
            result.notes.append(f"{policy.code}: original link expired; no replacement found")
        return

    text = visible_text(fetched.text)
    adopted, revised = extract_years(text)
    changed_parts: list[str] = []
    if adopted and old_adopted in {"", "N/A", "*"}:
        if add_update(result, snapshot, policy.adopted_col, adopted, GREEN_FILL_NAME):
            changed_parts.append(f"adopted year {old_adopted or 'blank'} -> {adopted}")
    if revised:
        old_year = as_int_year(old_revised)
        new_year = as_int_year(revised)
        if new_year and (old_year is None or new_year > old_year):
            if add_update(result, snapshot, policy.revised_col, revised, GREEN_FILL_NAME):
                changed_parts.append(f"revised year {old_revised or 'blank'} -> {revised}")
    if fetched.final_url and fetched.final_url != link:
        if add_update(result, snapshot, policy.link_col, fetched.final_url, GREEN_FILL_NAME):
            changed_parts.append("canonical link updated")

    if changed_parts:
        add_mark(result, policy.value_col, GREEN_FILL_NAME)
        result.notes.append(f"{policy.code}: " + "; ".join(changed_parts))


def process_missing_policy(
    result: RowResult,
    snapshot: RowSnapshot,
    policy: PolicyGroup,
    cache: ThreadSafeCache,
    timeout: int,
    delay: float,
) -> None:
    found_link, adopted, revised = find_policy_link(snapshot.district, policy, cache, timeout, delay)
    if not found_link:
        return
    add_update(result, snapshot, policy.value_col, "1", GREEN_FILL_NAME)
    add_update(result, snapshot, policy.link_col, found_link, GREEN_FILL_NAME)
    if adopted:
        add_update(result, snapshot, policy.adopted_col, adopted, GREEN_FILL_NAME)
    if revised:
        add_update(result, snapshot, policy.revised_col, revised, GREEN_FILL_NAME)
    result.notes.append(f"{policy.code}: policy found and proposed as passed")


def process_row(
    snapshot: RowSnapshot,
    groups: list[PolicyGroup],
    cache: ThreadSafeCache,
    timeout: int,
    delay: float,
) -> RowResult:
    result = RowResult(row=snapshot.row)
    for policy in groups:
        value = snapshot.values.get(policy.value_col, "")
        if value in {"1", "1.0"}:
            process_existing_policy(result, snapshot, policy, cache, timeout, delay)
        else:
            process_missing_policy(result, snapshot, policy, cache, timeout, delay)
    return result


def build_policy_groups(ws) -> list[PolicyGroup]:
    groups: list[PolicyGroup] = []
    for col in range(START_COL, END_COL + 1, GROUP_WIDTH):
        label = normalize_cell(ws.cell(HEADER_ROW, col).value)
        code, title = extract_policy_parts(label)
        groups.append(
            PolicyGroup(
                value_col=col,
                adopted_col=col + 1,
                revised_col=col + 2,
                link_col=col + 3,
                label=label,
                code=code,
                title=title,
            )
        )
    return groups


def build_snapshots(ws, start_row: int, end_row: int) -> list[RowSnapshot]:
    snapshots: list[RowSnapshot] = []
    for row in range(start_row, end_row + 1):
        district = normalize_cell(ws.cell(row, 4).value)
        if not district:
            continue
        values = {col: normalize_cell(ws.cell(row, col).value) for col in range(START_COL, END_COL + 1)}
        snapshots.append(RowSnapshot(row=row, district=district, values=values))
    return snapshots


def fill_for_name(name: str) -> PatternFill:
    if name == GREEN_FILL_NAME:
        return GREEN_FILL
    if name == RED_FILL_NAME:
        return RED_FILL
    if name == YELLOW_FILL_NAME:
        return YELLOW_FILL
    raise ValueError(f"Unknown fill name: {name}")


def mark_cell(cell, fill_name: str) -> None:
    cell.font = CHANGED_FONT
    cell.fill = fill_for_name(fill_name)


def apply_row_result(ws, result: RowResult, summary_col: int) -> None:
    for mark in result.marks:
        mark_cell(ws.cell(result.row, mark.col), mark.fill)
    for update in result.updates:
        cell = ws.cell(result.row, update.col)
        cell.value = update.value
        mark_cell(cell, update.fill)
    if result.notes:
        summary_cell = ws.cell(result.row, summary_col)
        summary_cell.value = " | ".join(result.notes)
        mark_cell(summary_cell, YELLOW_FILL_NAME)


def save_workbook_atomic(wb, output_path: Path) -> None:
    tmp_path = output_path.with_name(f"{output_path.stem}.tmp{output_path.suffix}")
    wb.save(tmp_path)
    tmp_path.replace(output_path)


def generate_review(
    input_path: Path,
    output_path: Path,
    cache_path: Path,
    start_row: int,
    end_row: int,
    workers: int,
    cache_save_interval: int,
    workbook_save_interval: int,
    timeout: int,
    delay: float,
) -> None:
    wb = load_workbook(input_path)
    if SHEET_NAME not in wb.sheetnames:
        raise SystemExit(f"Sheet not found: {SHEET_NAME}")
    ws = wb[SHEET_NAME]
    groups = build_policy_groups(ws)
    snapshots = build_snapshots(ws, start_row, end_row)
    cache = ThreadSafeCache(cache_path)

    summary_col = END_COL + 1
    ws.cell(HEADER_ROW, summary_col).value = "Summary of Changes"
    ws.cell(HEADER_ROW, summary_col).font = Font(bold=True)
    ws.column_dimensions[get_column_letter(summary_col)].width = 80

    completed = 0
    print(f"Queued {len(snapshots)} district rows with {workers} workers.", flush=True)
    try:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_snapshot = {
                executor.submit(process_row, snapshot, groups, cache, timeout, delay): snapshot
                for snapshot in snapshots
            }
            for future in as_completed(future_to_snapshot):
                snapshot = future_to_snapshot[future]
                try:
                    result = future.result()
                except Exception as exc:
                    print(f"Row {snapshot.row}: {snapshot.district} failed: {exc}", flush=True)
                    continue

                apply_row_result(ws, result, summary_col)
                completed += 1
                print(
                    f"Completed {completed}/{len(snapshots)}: "
                    f"row {snapshot.row} {snapshot.district} "
                    f"({len(result.notes)} note(s))",
                    flush=True,
                )

                if completed % cache_save_interval == 0:
                    cache.save_if_dirty()
                if completed % workbook_save_interval == 0:
                    save_workbook_atomic(wb, output_path)
    finally:
        cache.save_if_dirty()
        save_workbook_atomic(wb, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate Caden policy review workbook quickly.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    parser.add_argument("--start-row", type=int, default=START_ROW)
    parser.add_argument("--end-row", type=int, default=END_ROW)
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    parser.add_argument("--cache-save-interval", type=int, default=DEFAULT_CACHE_SAVE_INTERVAL)
    parser.add_argument("--workbook-save-interval", type=int, default=DEFAULT_WORKBOOK_SAVE_INTERVAL)
    parser.add_argument("--timeout", type=int, default=TIMEOUT_SECONDS)
    parser.add_argument("--delay", type=float, default=REQUEST_DELAY_SECONDS)
    args = parser.parse_args()

    generate_review(
        input_path=args.input,
        output_path=args.output,
        cache_path=args.cache,
        start_row=args.start_row,
        end_row=args.end_row,
        workers=args.workers,
        cache_save_interval=args.cache_save_interval,
        workbook_save_interval=args.workbook_save_interval,
        timeout=args.timeout,
        delay=args.delay,
    )
    print(f"Review workbook written to: {args.output.resolve()}", flush=True)


if __name__ == "__main__":
    main()
