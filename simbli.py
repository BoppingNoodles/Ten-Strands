"""
simbli.py - undetected-chromedriver based scraper for Simbli

Strategy for check_policy (already adopted):
  1. If the spreadsheet has a direct Simbli link → visit it, scrape dates.
  2. If the direct link fails (or no link) → fall back to scanning the district's
     full policy index for the matching policy code.

Strategy for search_for_policy (not yet adopted):
  - Scan the district's policy index to see if the policy has been newly adopted.
"""
import asyncio
import re
import random
import time
from bs4 import BeautifulSoup
import undetected_chromedriver as uc
from urllib.parse import urljoin

from models import (
    DistrictRecord,
    PolicyEntry,
    ScrapeResult,
    ScapeAction,
    HighlightColor,
    apply_policy_link,
    blank_not_found_result,
)

SIMBLI_SEARCH_URL = "https://simbli.eboardsolutions.com/Policy/PolicyListing.aspx?S={simbli_id}"

# Cache keyed by simbli_id → list of row dicts from the policy index
_INDEX_CACHE = {}

# Pacing helpers — keep requests human-paced without making full runs too slow.
_NAV_PAUSE_MIN = 0.8
_NAV_PAUSE_MAX = 1.6
_ACTION_PAUSE_MIN = 1.2
_ACTION_PAUSE_MAX = 2.2


def _human_pause(min_s: float = _ACTION_PAUSE_MIN, max_s: float = _ACTION_PAUSE_MAX) -> None:
    time.sleep(random.uniform(min_s, max_s))


def _wait_for_page(driver: uc.Chrome, extra_wait: float = 4.0):
    """Wait for Cloudflare to pass and Angular to render."""
    for _ in range(20):
        title = driver.title
        src = driver.page_source
        if "Just a moment" in title or "cf-browser-verification" in src:
            time.sleep(1.5)
        else:
            break
    time.sleep(extra_wait)


def _navigate(driver: uc.Chrome, url: str) -> bool:
    """Navigate to url. Returns False on timeout/error."""
    try:
        driver.get(url)
        _human_pause(_NAV_PAUSE_MIN, _NAV_PAUSE_MAX)
        return True
    except Exception as e:
        print(f"      [Simbli] Error loading {url}: {e}")
        return False


def _first_policy_list(value):
    if isinstance(value, list):
        if value and all(isinstance(item, dict) and "Policy" in item for item in value):
            return value
        for item in value:
            found = _first_policy_list(item)
            if found:
                return found
    elif isinstance(value, dict):
        for key in ("Policies", "PolicyList", "PolicyListing", "PolicyListingList", "Data"):
            found = _first_policy_list(value.get(key))
            if found:
                return found
        for item in value.values():
            found = _first_policy_list(item)
            if found:
                return found
    return []


def _rows_from_policy_api_data(data: dict, simbli_id: str) -> list[dict]:
    rows = []
    for item in _first_policy_list(data):
        policy = item.get("Policy") or {}
        content_type = policy.get("ContentType") or {}
        code = str(policy.get("Code") or "").strip()
        type_text = str(content_type.get("Abbreviation") or content_type.get("Name") or "").strip()
        title = str(policy.get("Description") or policy.get("DisplayTitle") or policy.get("Name") or "").strip()
        adopted = policy.get("OriginalAdoptedDate") or item.get("AdoptedDate") or item.get("OrigAdoptedDate")
        revised = policy.get("LastRevisedDate") or item.get("LastReviewedDate")
        date_text = " ".join(str(v) for v in (adopted, revised) if v)
        policy_id = item.get("ID") or policy.get("Id")
        link = None
        if policy_id:
            link = (
                "https://simbli.eboardsolutions.com/Policy/ViewPolicy.aspx"
                f"?S={simbli_id}&revid={policy_id}&ptid=&secid=&PG=6&IRP=0&isPndg=false"
            )
        if code or title:
            rows.append({
                "code": f"{type_text} {code}".strip(),
                "title": title,
                "date_text": date_text,
                "link": link,
            })
    return rows


def _get_policy_listing_api_rows(driver: uc.Chrome, simbli_id: str) -> list[dict]:
    try:
        data = driver.execute_async_script(
            """
            const done = arguments[0];
            const api = "/Services/api/PolicyListing/?sct=" + sToken +
              "&ensid=" + enSID +
              "&enUID=" + enCuUID +
              "&ismobile=false" +
              "&ptid=" + enPTID +
              "&secid=" + enSectionID;
            fetch(api, { headers: { Authorization: `Bearer ${getCoreAuthTokenFromCookies()}` } })
              .then((response) => response.json())
              .then((data) => done(data))
              .catch((error) => done({ error: String(error) }));
            """
        )
    except Exception as exc:
        print(f"      [Simbli] PolicyListing API failed: {exc}")
        return []
    if isinstance(data, dict) and data.get("error"):
        print(f"      [Simbli] PolicyListing API error: {data['error']}")
        return []
    rows = _rows_from_policy_api_data(data, simbli_id)
    if rows:
        print(f"      [Simbli] PolicyListing API returned {len(rows)} policies")
    return rows


# ── Direct-link scraping ──────────────────────────────────────────────────────

def _scrape_direct_link(driver: uc.Chrome, link: str) -> dict | None:
    """
    Visit a ViewPolicy page directly and extract:
      - original_adopted_year
      - last_revised_year
    Returns a dict with those keys, or None if the page couldn't be parsed.
    """
    print(f"      [Simbli] Trying direct link: {link}")
    ok = _navigate(driver, link)
    if not ok:
        return None

    _wait_for_page(driver, extra_wait=4.0)

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # Verify it's actually a ViewPolicy page and not an error/Cloudflare page
    if "ViewPolicy" not in driver.current_url and "ViewPolicy" not in html[:500]:
        # Check title contains policy number
        if "View Policy" not in driver.title and "Policy" not in driver.title:
            print(f"      [Simbli] Direct link does not appear to be a policy page.")
            return None

    result = {}

    # Look for "Original Adopted Date:" and "Last Revised Date:" labels
    for label_text in ["Original Adopted Date:", "Adopted Date:", "Adopted:"]:
        label = soup.find(string=re.compile(re.escape(label_text), re.I))
        if label:
            parent_text = label.parent.parent.get_text(" ", strip=True) if label.parent else ""
            year = _extract_year(parent_text)
            if year:
                result['original_adopted_year'] = year
                break

    for label_text in ["Last Revised Date:", "Revised Date:", "Revised:"]:
        label = soup.find(string=re.compile(re.escape(label_text), re.I))
        if label:
            parent_text = label.parent.parent.get_text(" ", strip=True) if label.parent else ""
            year = _extract_year(parent_text)
            if year:
                result['last_revised_year'] = year
                break

    if not result:
        # Try a broader approach — look for any date near the bottom metadata area
        metadata_area = soup.find("div", class_=re.compile(r"policy.?meta|meta.?data|policyDates", re.I))
        if metadata_area:
            text = metadata_area.get_text(" ", strip=True)
            years = re.findall(r'\b(19|20)\d{2}\b', text)
            if years:
                result['last_revised_year'] = years[-1]

    print(f"      [Simbli] Direct link result: {result or 'No dates found'}")
    return result if result else None


# ── Policy index scraping ─────────────────────────────────────────────────────

def _get_policy_listing_sync(driver: uc.Chrome, simbli_id: str) -> list:
    """
    Fetches the district's policy index and returns a list of row dicts.
    Results are cached per simbli_id.
    """
    simbli_id = str(simbli_id).strip()
    if simbli_id in _INDEX_CACHE:
        return _INDEX_CACHE[simbli_id]

    url = SIMBLI_SEARCH_URL.format(simbli_id=simbli_id)
    print(f"      [Simbli] Loading policy index: {url}")

    ok = _navigate(driver, url)
    if not ok:
        _INDEX_CACHE[simbli_id] = []
        return []

    _wait_for_page(driver, extra_wait=4.0)

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    rows_data = []

    api_rows = _get_policy_listing_api_rows(driver, simbli_id)
    if api_rows:
        _INDEX_CACHE[simbli_id] = api_rows
        return api_rows

    def add_safe_routes_link_rows():
        """Capture Safe Routes rows from rendered links when table parsing misses them."""
        seen = {(r.get("code"), r.get("link")) for r in rows_data}
        for a_tag in soup.find_all("a", href=True):
            link_text = a_tag.get_text(" ", strip=True)
            context_text = ""
            parent = a_tag.find_parent(["tr", "li", "div", "td"])
            if parent:
                context_text = parent.get_text(" ", strip=True)
            combined = " ".join([link_text, context_text])
            combined_lower = combined.lower()
            has_safe_routes = "safe routes" in combined_lower and "school" in combined_lower
            has_5142_2 = re.search(r"\b5142\.2(?:\([a-z]\))?\b", combined, re.I)
            if not (has_safe_routes or has_5142_2):
                continue

            code_match = re.search(r"\b(BP|AR)\s*5142\.2(?:\([a-z]\))?\b", combined, re.I)
            code = code_match.group(0).upper() if code_match else "5142.2"
            link = urljoin("https://simbli.eboardsolutions.com/", a_tag["href"])
            key = (code, link)
            if key in seen:
                continue
            seen.add(key)
            rows_data.append({
                'code': code,
                'title': link_text or "Safe Routes to School",
                'date_text': context_text,
                'link': link
            })

    # Format 1: legacy PolicyList table
    table = soup.select_one("table.PolicyList")
    if table:
        for row in table.select("tr")[1:]:
            cols = row.select("td")
            if len(cols) >= 3:
                code_node = cols[0]
                title_node = cols[1]
                rev_node = cols[-1]
                a_tag = code_node.select_one("a")
                link = a_tag['href'] if a_tag and a_tag.has_attr('href') else None
                if link:
                    link = urljoin("https://simbli.eboardsolutions.com/Policy/", link)
                rows_data.append({
                    'code': code_node.get_text(strip=True),
                    'title': title_node.get_text(strip=True),
                    'date_text': rev_node.get_text(strip=True),
                    'link': link
                })
        add_safe_routes_link_rows()
        _INDEX_CACHE[simbli_id] = rows_data
        return rows_data

    # Format 2: Angular policyFormat table
    table = soup.select_one("table.policyFormat, div.pl-grid-wrap table")
    if not table:
        add_safe_routes_link_rows()
        _INDEX_CACHE[simbli_id] = rows_data
        return rows_data

    for row in table.select("tr"):
        cols = row.select("td")
        if len(cols) >= 5:
            # Columns: 0=Code, 1=Title+Link, 2=Type, 3=Revised, 4=Reviewed
            code_text = cols[0].get_text(strip=True)
            type_text = cols[2].get_text(strip=True)
            full_code = f"{type_text} {code_text}".strip()

            title_node = cols[1]
            title_text = title_node.get_text(strip=True)

            a_tag = title_node.select_one("a")
            link = a_tag['href'] if a_tag and a_tag.has_attr('href') else None
            if link:
                link = urljoin("https://simbli.eboardsolutions.com/", link)

            date_text = cols[3].get_text(strip=True)

            rows_data.append({
                'code': full_code,
                'title': title_text,
                'date_text': date_text,
                'link': link
            })

    add_safe_routes_link_rows()
    _INDEX_CACHE[simbli_id] = rows_data
    return rows_data


def _normalize_policy_code(value: str) -> str:
    return re.sub(r"[^a-z0-9.]", "", str(value).lower())


def _policy_number(value: str) -> str:
    match = re.search(r"\d+(?:\.\d+)+", str(value))
    return match.group(0) if match else ""


def _is_safe_routes_target(target_code: str) -> bool:
    return _policy_number(target_code) == "5142.2"


def _safe_routes_not_found_result(result: ScrapeResult) -> ScrapeResult:
    result.action = ScapeAction.UNCHANGED
    result.highlight_color = HighlightColor.NONE
    result.new_value = "0"
    result.new_year_adopted = "N/A"
    result.new_year_revised = "N/A"
    result.new_link = "N/A"
    result.notes = "Safe Routes policy not found; normalized blank cells to 0/N/A"
    return result


def _safe_routes_match(row: dict, target_code: str) -> bool:
    row_code = str(row.get("code", ""))
    row_title = str(row.get("title", ""))
    target_type = target_code.split()[0].lower() if " " in target_code else ""
    row_type = row_code.split()[0].lower() if " " in row_code else ""
    row_number = _policy_number(row_code)
    title_matches = "safe routes" in row_title.lower() and "school" in row_title.lower()
    number_matches = row_number == "5142.2"
    type_matches = not target_type or not row_type or target_type == row_type
    return type_matches and (number_matches or title_matches)


def _find_matching_policy(rows_data: list, target_code: str) -> dict | None:
    """Find a policy row in the index that matches the target code."""
    clean_target = _normalize_policy_code(target_code)
    for r in rows_data:
        clean_code = _normalize_policy_code(r.get("code", ""))
        if clean_code == clean_target or clean_code.endswith(clean_target):
            return r
        if _is_safe_routes_target(target_code) and _safe_routes_match(r, target_code):
            return r
    # Fallback: match just the number portion
    if " " in target_code:
        _, num = target_code.split(" ", 1)
        clean_num = _normalize_policy_code(num)
        for r in rows_data:
            clean_code = _normalize_policy_code(r.get("code", ""))
            if clean_code == clean_num or clean_code.startswith(clean_num):
                return r
    return None


def _extract_year(text: str) -> str | None:
    m = re.search(r'\b(19|20)\d{2}\b', text)
    return m.group(0) if m else None


def _extract_years(text: str) -> list[str]:
    return [m.group(0) for m in re.finditer(r'\b(19|20)\d{2}\b', str(text))]


# ── Core logic ────────────────────────────────────────────────────────────────

def _check_policy_sync(district: DistrictRecord, policy: PolicyEntry, driver: uc.Chrome) -> ScrapeResult:
    result = ScrapeResult(
        cds_code=district.cds_code, district_name=district.district_name,
        policy_code=policy.policy_code, action=ScapeAction.UNCHANGED,
        highlight_color=HighlightColor.NONE, old_value=policy.value,
        old_year_revised=policy.year_revised, old_link=policy.link,
        col_start=policy.col_start
    )

    if not district.simbli_id:
        if _is_safe_routes_target(policy.policy_code):
            return _safe_routes_not_found_result(result)
        result.action = ScapeAction.SKIPPED
        result.notes = "No simbli_id found"
        return result

    baseline_year = policy.max_year

    # ── Step 1: Try the direct link from the spreadsheet ──────────────────────
    direct_result = None
    if policy.has_real_link and policy.is_simbli:
        direct_result = _scrape_direct_link(driver, policy.link)

    if direct_result:
        # Determine the best year from the direct page
        found_year = direct_result.get('last_revised_year') or direct_result.get('original_adopted_year')
        if found_year and baseline_year:
            if int(found_year) > baseline_year:
                result.action = ScapeAction.REVISED
                result.highlight_color = HighlightColor.GREEN
                result.new_year_revised = found_year
                result.notes = f"Revised (via direct link): {baseline_year} → {found_year}"
            else:
                result.notes = f"Unchanged (via direct link): year={found_year}"
        else:
            result.notes = f"Direct link OK, no parseable dates found"
        return result

    # ── Step 2: Fall back to policy index scan ────────────────────────────────
    print(f"      [Simbli] Falling back to index scan for {policy.policy_code}")
    _human_pause()
    try:
        rows = _get_policy_listing_sync(driver, district.simbli_id)
        match = _find_matching_policy(rows, policy.policy_code)

        if not match:
            apply_policy_link(result, policy, None)
            result.action = ScapeAction.LINK_DEAD
            result.highlight_color = HighlightColor.RED
            result.notes = "Policy not found in index (direct link also failed)"
            return result

        years = _extract_years(match['date_text'])
        found_year = years[-1] if years else None
        apply_policy_link(result, policy, match.get("link"))

        if found_year and baseline_year:
            if int(found_year) > baseline_year:
                result.action = ScapeAction.REVISED
                result.highlight_color = HighlightColor.GREEN
                result.new_year_revised = found_year
                result.notes = f"Revised (via index): {baseline_year} → {found_year}"
            else:
                result.notes = f"Unchanged (via index): year={found_year}"
        else:
            result.notes = "Could not parse year from index"

    except Exception as exc:
        result.action = ScapeAction.ERROR
        result.notes = str(exc)

    return result


def _search_for_policy_sync(district: DistrictRecord, policy: PolicyEntry, driver: uc.Chrome) -> ScrapeResult:
    """For policies not yet adopted: scan the index to see if they've been newly adopted."""
    result = ScrapeResult(
        cds_code=district.cds_code, district_name=district.district_name,
        policy_code=policy.policy_code, action=ScapeAction.UNCHANGED,
        highlight_color=HighlightColor.NONE, old_value=policy.value,
        old_year_revised=policy.year_revised, old_link=policy.link,
        col_start=policy.col_start
    )

    if not district.simbli_id:
        result.action = ScapeAction.SKIPPED
        result.notes = "No simbli_id found"
        return result

    try:
        rows = _get_policy_listing_sync(driver, district.simbli_id)
        match = _find_matching_policy(rows, policy.policy_code)

        if match:
            years = _extract_years(match['date_text'])
            adopted_year = years[0] if years else None
            revised_year = years[-1] if years else None
            result.action = ScapeAction.NEWLY_FOUND
            result.highlight_color = HighlightColor.GREEN
            result.new_value = "1"
            if adopted_year:
                result.new_year_adopted = adopted_year
            if revised_year:
                result.new_year_revised = revised_year
            apply_policy_link(result, policy, match.get("link"))
            result.notes = f"Newly found in index! Adopted: {adopted_year}; Revised: {revised_year}"
        elif policy.is_blank_block:
            return blank_not_found_result(
                district,
                policy,
                "Policy not found; normalized blank cells to 0/N/A",
            )
        elif _is_safe_routes_target(policy.policy_code):
            _safe_routes_not_found_result(result)

    except Exception as exc:
        result.action = ScapeAction.ERROR
        result.notes = str(exc)

    return result


# ── Async wrappers ────────────────────────────────────────────────────────────

async def check_policy(district: DistrictRecord, policy: PolicyEntry, ctx, delay_min: int, delay_max: int) -> ScrapeResult:
    driver, lock = ctx
    async with lock:
        await asyncio.sleep(random.uniform(delay_min, delay_max))
        return await asyncio.to_thread(_check_policy_sync, district, policy, driver)


async def search_for_policy(district: DistrictRecord, policy: PolicyEntry, ctx, delay_min: int, delay_max: int) -> ScrapeResult:
    driver, lock = ctx
    async with lock:
        await asyncio.sleep(random.uniform(delay_min, delay_max))
        return await asyncio.to_thread(_search_for_policy_sync, district, policy, driver)
