"""
discover.py — Discovers Simbli or BoardDocs IDs for districts that have no
existing links in the spreadsheet, using a real browser Google search to
bypass anti-bot detection.
"""

import re
import asyncio
import time
import unicodedata

_SIMBLI_URL_RE = re.compile(
    r'https?://[^\s"\'>]*simbli\.eboardsolutions\.com[^\s"\'>]*[?&]S=(\d{5,})',
    re.IGNORECASE
)
_BOARDDOCS_URL_RE = re.compile(
    r'https?://[^\s"\'>]*boarddocs\.com/[^/]+/([A-Za-z0-9_-]{2,})/[^\s"\'>]*',
    re.IGNORECASE
)


def _search_google_sync(driver, query: str) -> str:
    """Navigate to Google search and return the page source. Runs synchronously."""
    search_url = "https://www.google.com/search?q=" + query.replace(' ', '+') + "&num=5"
    driver.get(search_url)
    time.sleep(2)  # Let the page settle
    return driver.page_source


def _extract_simbli_id(html: str) -> str | None:
    """Parse Simbli S=XXXXXXXX from full simbli.eboardsolutions.com URLs in page source."""
    m = _SIMBLI_URL_RE.search(html)
    return m.group(1) if m else None


def _extract_boarddocs_slug(html: str) -> str | None:
    """Parse BoardDocs slug from full boarddocs.com URLs in page source."""
    m = _BOARDDOCS_URL_RE.search(html)
    return m.group(1) if m else None


def _normalize_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"\b(school|district|schools|elementary|unified|union|joint|high)\b", " ", value, flags=re.I)
    value = re.sub(r"[^a-z0-9]+", " ", value.lower())
    return " ".join(value.split())


def _name_matches(candidate_name: str, page_text: str) -> bool:
    candidate = _normalize_name(candidate_name)
    page = _normalize_name(page_text)
    if not candidate or not page:
        return False
    candidate_terms = set(candidate.split())
    page_terms = set(page.split())
    if candidate in page:
        return True
    return bool(candidate_terms) and candidate_terms.issubset(page_terms)


def _verify_simbli_id_sync(driver, district_name: str, simbli_id: str) -> bool:
    url = f"https://simbli.eboardsolutions.com/Policy/PolicyListing.aspx?S={simbli_id}"
    try:
        driver.get(url)
        time.sleep(2)
    except Exception:
        return False
    page_text = " ".join([driver.title or "", driver.page_source[:5000]])
    return _name_matches(district_name, page_text)


async def discover_platform(district, driver, lock) -> tuple[str | None, str | None]:
    """
    For a district with no known Simbli ID or BoardDocs slug, search Google
    using the real Chrome driver (bypasses anti-bot).

    Returns (simbli_id, boarddocs_slug) — at most one will be non-None.
    """
    name = district.district_name
    print(f"  [Discover] Searching for platform: '{name}'")

    loop = asyncio.get_event_loop()

    # Search for Simbli first — no site: filter because Simbli pages are NOINDEX.
    # Instead we search by district name + platform name and find the URL in results.
    async with lock:
        simbli_query = f'"{name}" simbli eboardsolutions'
        simbli_html = await loop.run_in_executor(
            None, _search_google_sync, driver, simbli_query
        )

    simbli_id = _extract_simbli_id(simbli_html)

    if simbli_id:
        async with lock:
            is_verified = await loop.run_in_executor(
                None, _verify_simbli_id_sync, driver, name, simbli_id
            )
        if is_verified:
            print(f"  [Discover] Found verified Simbli ID for '{name}': S={simbli_id}")
            return simbli_id, None
        print(f"  [Discover] Rejected Simbli ID for '{name}' after title/name verification: S={simbli_id}")

    # If no Simbli found, search for BoardDocs
    await asyncio.sleep(1.5)
    async with lock:
        boarddocs_query = f'"{name}" boarddocs.com board policies'
        boarddocs_html = await loop.run_in_executor(
            None, _search_google_sync, driver, boarddocs_query
        )

    boarddocs_slug = _extract_boarddocs_slug(boarddocs_html)

    if boarddocs_slug:
        print(f"  [Discover] Found BoardDocs slug for '{name}': {boarddocs_slug}")
        return None, boarddocs_slug

    print(f"  [Discover] No platform found for '{name}', will stay N/A")
    return None, None


async def discover_missing_platforms(districts: list, simbli_ctx: tuple) -> None:
    """
    Mutates each district in-place: for those with no simbli_id AND no
    boarddocs_slug, attempts a Google search to discover the platform.
    Only runs for districts that have at least one 0/N/A policy (value check),
    since all-* districts genuinely have no site.
    """
    driver, lock = simbli_ctx

    def has_no_database_markers(district) -> bool:
        return any(
            str(value).strip() == "*"
            for policy in district.policies
            for value in (policy.value, policy.year_adopted, policy.year_revised, policy.link)
        )

    candidates = [
        d for d in districts
        if not d.simbli_id
        and not d.boarddocs_slug
        and not has_no_database_markers(d)
        and any(p.is_not_adopted for p in d.policies)
    ]

    if not candidates:
        print("[Discover] No link-less districts need platform discovery.")
        return

    print(f"\n[Discover] {len(candidates)} link-less districts to probe for platforms...")

    # Run sequentially to avoid hammering Google
    for district in candidates:
        simbli_id, boarddocs_slug = await discover_platform(district, driver, lock)
        if simbli_id:
            district.simbli_id = simbli_id
        elif boarddocs_slug:
            district.boarddocs_slug = boarddocs_slug
        await asyncio.sleep(2)  # Polite pause between searches

    print("[Discover] Platform discovery complete.\n")
