"""
simbli.py — Playwright-based scraper for Simbli
"""

import asyncio
import re
import random
from urllib.parse import urlencode
from models import DistrictRecord, PolicyEntry, ScrapeResult, ScapeAction, HighlightColor

# The base URL for searching policies on a specific board
SIMBLI_SEARCH_URL = "https://simbli.eboardsolutions.com/Policy/PolicyListing.aspx?S={simbli_id}"

async def _get_policy_listing(page, simbli_id: str):
    """
    Navigates to the policy listing page and extracts all rows.
    Returns a list of dicts: {"code": "BP 3510", "title": "...", "revised": "2019", "link": "..."}
    """
    url = SIMBLI_SEARCH_URL.format(simbli_id=simbli_id)
    await page.goto(url, wait_until="domcontentloaded")
    
    # Wait for the table to appear, or timeout if empty/blocked
    try:
        await page.wait_for_selector("table.PolicyList", timeout=10000)
    except Exception:
        # Check if we got Cloudflare blocked
        title = await page.title()
        if "Just a moment" in title or "Cloudflare" in title:
            raise PermissionError("Cloudflare blocked")
        return [] # No table found, empty policies
        
    # Extract data from the table rows
    # Note: Simbli DOM structure needs to be verified during pilot. 
    # Assuming standard table with rows where first col is Code, second is Title, etc.
    # We will use evaluate to extract text.
    
    rows_data = await page.evaluate('''() => {
        const rows = document.querySelectorAll("table.PolicyList tr");
        const data = [];
        for (let i = 1; i < rows.length; i++) { // skip header
            const cols = rows[i].querySelectorAll("td");
            if (cols.length >= 3) {
                const codeNode = cols[0];
                const titleNode = cols[1];
                const revNode = cols[cols.length - 1]; // Assume last col is Adopted/Revised date
                
                const aTag = codeNode.querySelector("a");
                const link = aTag ? aTag.href : null;
                
                data.push({
                    code: codeNode.innerText.trim(),
                    title: titleNode.innerText.trim(),
                    date_text: revNode.innerText.trim(),
                    link: link
                });
            }
        }
        return data;
    }''')
    
    return rows_data

def _find_matching_policy(rows_data, target_code: str):
    """
    Tries to find the policy in the extracted rows.
    target_code is e.g. "BP 3510" or "AR 3514.1".
    """
    # Clean target
    clean_target = target_code.replace(" ", "").lower()
    
    for r in rows_data:
        clean_code = r['code'].replace(" ", "").lower()
        if clean_code == clean_target or clean_code.endswith(clean_target):
            return r
            
    # Sometimes codes are missing prefixes like BP or AR in the listing
    if " " in target_code:
        prefix, num = target_code.split(" ", 1)
        clean_num = num.replace(" ", "").lower()
        for r in rows_data:
            clean_code = r['code'].replace(" ", "").lower()
            if clean_num == clean_code:
                return r
                
    return None

def _extract_year(date_text: str) -> str | None:
    """Extracts a 4-digit year from a date string like '10/24/2019'."""
    m = re.search(r'\b(19|20)\d{2}\b', date_text)
    return m.group(0) if m else None

async def check_policy(district: DistrictRecord, policy: PolicyEntry, page, delay_min: int, delay_max: int) -> ScrapeResult:
    """
    Checks if an existing policy has been revised.
    """
    result = ScrapeResult(
        cds_code=district.cds_code,
        district_name=district.district_name,
        policy_code=policy.policy_code,
        action=ScapeAction.UNCHANGED,
        highlight_color=HighlightColor.NONE,
        old_value=policy.value,
        old_year_revised=policy.year_revised,
        old_link=policy.link,
        col_start=policy.col_start
    )
    
    if not district.simbli_id:
        result.action = ScapeAction.SKIPPED
        result.notes = "No simbli_id found"
        return result
        
    await asyncio.sleep(random.uniform(delay_min, delay_max))
    
    try:
        rows = await _get_policy_listing(page, district.simbli_id)
        match = _find_matching_policy(rows, policy.policy_code)
        
        if not match:
            # Policy is missing now? Link might be dead
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
                # Same or older year
                pass 
        else:
            result.notes = "Could not parse years for comparison"
            
    except PermissionError:
        result.action = ScapeAction.BOT_DETECTED
        result.notes = "Cloudflare blocked"
    except Exception as exc:
        result.action = ScapeAction.ERROR
        result.notes = str(exc)
        
    return result

async def search_for_policy(district: DistrictRecord, policy: PolicyEntry, page, delay_min: int, delay_max: int) -> ScrapeResult:
    """
    Searches for a policy that is currently marked 0/N/A to see if it was newly adopted.
    """
    result = ScrapeResult(
        cds_code=district.cds_code,
        district_name=district.district_name,
        policy_code=policy.policy_code,
        action=ScapeAction.UNCHANGED,
        highlight_color=HighlightColor.NONE,
        old_value=policy.value,
        old_year_revised=policy.year_revised,
        old_link=policy.link,
        col_start=policy.col_start
    )
    
    if not district.simbli_id:
        result.action = ScapeAction.SKIPPED
        result.notes = "No simbli_id found"
        return result
        
    await asyncio.sleep(random.uniform(delay_min, delay_max))
    
    try:
        rows = await _get_policy_listing(page, district.simbli_id)
        match = _find_matching_policy(rows, policy.policy_code)
        
        if match:
            # We found a policy that was previously 0/N/A!
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
            
    except PermissionError:
        result.action = ScapeAction.BOT_DETECTED
        result.notes = "Cloudflare blocked"
    except Exception as exc:
        result.action = ScapeAction.ERROR
        result.notes = str(exc)
        
    return result
