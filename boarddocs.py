"""
boarddocs.py — API-based scraper for BoardDocs
"""

import httpx
import re
from models import DistrictRecord, PolicyEntry, ScrapeResult, ScapeAction, HighlightColor

# BoardDocs has a public JSON API for boards
BOARDDOCS_API_URL = "https://go.boarddocs.com/ca/{slug}/Board.nsf/policies?open&format=json"

async def _fetch_policies_api(slug: str) -> list[dict]:
    """Fetches all policies for a boarddocs slug using their JSON API."""
    url = BOARDDOCS_API_URL.format(slug=slug)
    try:
        async with httpx.AsyncClient(timeout=10.0) as client:
            resp = await client.get(url)
            if resp.status_code == 200:
                # API returns JSON array of policy objects
                # e.g. [{"code": "BP 3510", "name": "...", "id": "...", "revised": "..."}]
                # Note: The exact JSON structure will need to be verified during pilot.
                # If it's not JSON, or requires auth, this will fail and we'll log it.
                data = resp.json()
                if isinstance(data, list):
                    return data
                elif isinstance(data, dict) and "policies" in data:
                    return data["policies"]
    except Exception:
        pass
    return []

def _find_matching_policy(policies: list[dict], target_code: str):
    clean_target = target_code.replace(" ", "").lower()
    for p in policies:
        code = p.get("code", "") or p.get("name", "")
        clean_code = code.replace(" ", "").lower()
        if clean_target in clean_code:
            return p
    return None

def _extract_year(date_text: str) -> str | None:
    if not date_text: return None
    m = re.search(r'\b(19|20)\d{2}\b', str(date_text))
    return m.group(0) if m else None

async def check_policy(district: DistrictRecord, policy: PolicyEntry, page, delay_min: int, delay_max: int) -> ScrapeResult:
    """Check a BoardDocs policy."""
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
    
    if not district.boarddocs_slug:
        result.action = ScapeAction.SKIPPED
        result.notes = "No boarddocs_slug found"
        return result
        
    policies = await _fetch_policies_api(district.boarddocs_slug)
    if not policies:
        # Fallback to checking if the link is alive using generic HTTP check
        # because the API might not be public or we have the wrong endpoint
        result.notes = "BoardDocs API returned no data, doing basic link check"
        import generic
        return await generic.check_link(district, policy)
        
    match = _find_matching_policy(policies, policy.policy_code)
    if not match:
        result.action = ScapeAction.LINK_DEAD
        result.highlight_color = HighlightColor.RED
        result.notes = "Policy not found in API"
        return result
        
    # Check revision date
    rev_date = match.get("revised") or match.get("last_modified") or match.get("adopted")
    found_year = _extract_year(rev_date)
    
    baseline_year = policy.max_year
    if found_year and baseline_year and int(found_year) > baseline_year:
        result.action = ScapeAction.REVISED
        result.highlight_color = HighlightColor.GREEN
        result.new_year_revised = found_year
        result.notes = f"Revised: {baseline_year} -> {found_year}"
        
        # Build direct link if we have the ID
        pid = match.get("id")
        if pid:
            result.new_link = f"https://go.boarddocs.com/ca/{district.boarddocs_slug}/Board.nsf/goto?open&id={pid}"
            
    return result

async def search_for_policy(district: DistrictRecord, policy: PolicyEntry, page, delay_min: int, delay_max: int) -> ScrapeResult:
    """Search for a 0/N/A policy in BoardDocs."""
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
    
    if not district.boarddocs_slug:
        result.action = ScapeAction.SKIPPED
        return result
        
    policies = await _fetch_policies_api(district.boarddocs_slug)
    if not policies:
        result.notes = "API empty"
        return result
        
    match = _find_matching_policy(policies, policy.policy_code)
    if match:
        rev_date = match.get("revised") or match.get("last_modified") or match.get("adopted")
        found_year = _extract_year(rev_date)
        
        result.action = ScapeAction.NEWLY_FOUND
        result.highlight_color = HighlightColor.GREEN
        result.new_value = "1"
        if found_year:
            result.new_year_revised = found_year
            
        pid = match.get("id")
        if pid:
            result.new_link = f"https://go.boarddocs.com/ca/{district.boarddocs_slug}/Board.nsf/goto?open&id={pid}"
            
        result.notes = f"Newly found! Year: {found_year}"
        
    return result
