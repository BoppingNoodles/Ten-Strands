"""
simbli.py - undetected-chromedriver based scraper for Simbli
"""
import asyncio
import re
import random
import time
from bs4 import BeautifulSoup
import undetected_chromedriver as uc
from urllib.parse import urljoin

from models import DistrictRecord, PolicyEntry, ScrapeResult, ScapeAction, HighlightColor

# The base URL for searching policies on a specific board
SIMBLI_SEARCH_URL = "https://simbli.eboardsolutions.com/Policy/PolicyListing.aspx?S={simbli_id}"

_CACHE = {}

def _get_policy_listing_sync(driver: uc.Chrome, simbli_id: str):
    """
    Fetches the policy listing page and extracts all rows using BeautifulSoup.
    Uses caching to avoid fetching the same district's page multiple times.
    """
    if simbli_id in _CACHE:
        return _CACHE[simbli_id]
        
    url = SIMBLI_SEARCH_URL.format(simbli_id=simbli_id)
    print(f"      [Simbli] Navigating to {url}")
    driver.get(url)
    
    # Wait for Cloudflare
    for _ in range(15):
        if "Just a moment" in driver.title or "cf-browser-verification" in driver.page_source:
            time.sleep(1)
        else:
            break
            
    time.sleep(3) # Give Angular a moment to render the grid
    
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    
    rows_data = []
    
    # 1. Check old PolicyList table first just in case some districts use older format
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
        _CACHE[simbli_id] = rows_data
        return rows_data
        
    # 2. Check new policyFormat table
    table = soup.select_one("table.policyFormat, div.pl-grid-wrap table")
    if not table:
        _CACHE[simbli_id] = []
        return []
        
    for row in table.select("tr"):
        cols = row.select("td")
        if len(cols) >= 5:
            # 0: Code, 1: Title+Link, 2: Type, 3: Revised, 4: Reviewed
            code_text = cols[0].get_text(strip=True)
            type_text = cols[2].get_text(strip=True)
            
            # Combine code and type for easy matching (e.g. '0000' + 'BP' = 'BP 0000')
            full_code = f"{type_text} {code_text}".strip()
            
            title_node = cols[1]
            title_text = title_node.get_text(strip=True)
            
            a_tag = title_node.select_one("a")
            link = a_tag['href'] if a_tag and a_tag.has_attr('href') else None
            if link:
                link = urljoin("https://simbli.eboardsolutions.com/", link)
                
            rev_node = cols[3]
            date_text = rev_node.get_text(strip=True)
            
            rows_data.append({
                'code': full_code,
                'title': title_text,
                'date_text': date_text,
                'link': link
            })
            
    _CACHE[simbli_id] = rows_data
    return rows_data

def _find_matching_policy(rows_data, target_code: str):
    clean_target = target_code.replace(" ", "").lower()
    for r in rows_data:
        clean_code = r['code'].replace(" ", "").lower()
        if clean_code == clean_target or clean_code.endswith(clean_target):
            return r
            
    if " " in target_code:
        prefix, num = target_code.split(" ", 1)
        clean_num = num.replace(" ", "").lower()
        for r in rows_data:
            clean_code = r['code'].replace(" ", "").lower()
            if clean_num == clean_code:
                return r
    return None

def _extract_year(date_text: str) -> str | None:
    m = re.search(r'\b(19|20)\d{2}\b', date_text)
    return m.group(0) if m else None

def _check_policy_sync(district: DistrictRecord, policy: PolicyEntry, driver: uc.Chrome) -> ScrapeResult:
    result = ScrapeResult(
        cds_code=district.cds_code, district_name=district.district_name,
        policy_code=policy.policy_code, action=ScapeAction.UNCHANGED,
        highlight_color=HighlightColor.NONE, old_value=policy.value,
        old_year_revised=policy.year_revised, old_link=policy.link, col_start=policy.col_start
    )
    if not district.simbli_id:
        result.action = ScapeAction.SKIPPED
        result.notes = "No simbli_id found"
        return result
        
    try:
        rows = _get_policy_listing_sync(driver, district.simbli_id)
        match = _find_matching_policy(rows, policy.policy_code)
        
        if not match:
            result.action = ScapeAction.LINK_DEAD
            result.highlight_color = HighlightColor.RED
            result.notes = "Policy not found in listing"
            return result
            
        found_year = _extract_year(match['date_text'])
        new_link = match['link']
        baseline_year = policy.max_year
        
        if found_year and baseline_year:
            if int(found_year) > baseline_year:
                result.action = ScapeAction.REVISED
                result.highlight_color = HighlightColor.GREEN
                result.new_year_revised = found_year
                if new_link and new_link != result.old_link:
                    result.new_link = new_link
                result.notes = f"Revised: {baseline_year} -> {found_year}"
        else:
            result.notes = "Could not parse years for comparison"
            
    except Exception as exc:
        result.action = ScapeAction.ERROR
        result.notes = str(exc)
        
    return result

def _search_for_policy_sync(district: DistrictRecord, policy: PolicyEntry, driver: uc.Chrome) -> ScrapeResult:
    result = ScrapeResult(
        cds_code=district.cds_code, district_name=district.district_name,
        policy_code=policy.policy_code, action=ScapeAction.UNCHANGED,
        highlight_color=HighlightColor.NONE, old_value=policy.value,
        old_year_revised=policy.year_revised, old_link=policy.link, col_start=policy.col_start
    )
    if not district.simbli_id:
        result.action = ScapeAction.SKIPPED
        result.notes = "No simbli_id found"
        return result
        
    try:
        rows = _get_policy_listing_sync(driver, district.simbli_id)
        match = _find_matching_policy(rows, policy.policy_code)
        
        if match:
            found_year = _extract_year(match['date_text'])
            new_link = match['link']
            
            result.action = ScapeAction.NEWLY_FOUND
            result.highlight_color = HighlightColor.GREEN
            result.new_value = "1"
            if found_year:
                result.new_year_revised = found_year
            if new_link:
                result.new_link = new_link
            result.notes = f"Newly found! Year: {found_year}"
            
    except Exception as exc:
        result.action = ScapeAction.ERROR
        result.notes = str(exc)
        
    return result

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
