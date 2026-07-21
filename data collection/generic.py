"""
generic.py — Checks standard URLs (PDFs, district websites, Google Drive, etc)
"""

import httpx
from models import DistrictRecord, PolicyEntry, ScrapeResult, ScapeAction, HighlightColor

async def check_link(district: DistrictRecord, policy: PolicyEntry) -> ScrapeResult:
    """
    Checks if a generic URL is still alive (HTTP 200).
    """
    url = str(policy.link).strip()
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
    
    try:
        async with httpx.AsyncClient(follow_redirects=True, timeout=10.0) as client:
            resp = await client.head(url)
            
            # Google Drive often returns 403 or redirects to login if restricted, or 200 if open.
            # 404/410 means dead
            if resp.status_code in [404, 410]:
                result.action = ScapeAction.LINK_DEAD
                result.highlight_color = HighlightColor.RED
                result.notes = f"HTTP {resp.status_code} - Link dead"
            elif resp.status_code >= 400:
                result.action = ScapeAction.ERROR
                result.notes = f"HTTP {resp.status_code}"
            else:
                # If it redirected to a different domain or path
                if str(resp.url) != url:
                    result.action = ScapeAction.LINK_REDIRECT
                    result.new_link = str(resp.url)
                    result.notes = "URL redirected"
                    
    except httpx.RequestError as exc:
        result.action = ScapeAction.LINK_DEAD
        result.highlight_color = HighlightColor.RED
        result.notes = f"Connection error: {type(exc).__name__}"
        
    return result
