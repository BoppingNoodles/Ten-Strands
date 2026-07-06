"""
boarddocs.py — API-based scraper for BoardDocs
"""

import httpx
import re
from models import (
    DistrictRecord,
    PolicyEntry,
    ScrapeResult,
    ScapeAction,
    HighlightColor,
    apply_policy_link,
    blank_not_found_result,
)

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


def _safe_routes_match(policy: dict, target_code: str) -> bool:
    code = str(policy.get("code", "") or "")
    name = str(policy.get("name", "") or policy.get("title", "") or "")
    target_type = target_code.split()[0].lower() if " " in target_code else ""
    row_type = code.split()[0].lower() if " " in code else ""
    type_matches = not target_type or not row_type or target_type == row_type
    number_matches = _policy_number(code) == "5142.2" or _policy_number(name) == "5142.2"
    title_matches = "safe routes" in name.lower() and "school" in name.lower()
    return type_matches and (number_matches or title_matches)


def _find_matching_policy(policies: list[dict], target_code: str):
    clean_target = _normalize_policy_code(target_code)
    for p in policies:
        code = p.get("code", "") or p.get("name", "")
        clean_code = _normalize_policy_code(code)
        if clean_target in clean_code:
            return p
        if _is_safe_routes_target(target_code) and _safe_routes_match(p, target_code):
            return p
    return None


def _boarddocs_policy_link(district: DistrictRecord, match: dict) -> str | None:
    pid = match.get("id")
    if pid:
        return f"https://go.boarddocs.com/ca/{district.boarddocs_slug}/Board.nsf/goto?open&id={pid}"
    return None


def _extract_year(date_text: str) -> str | None:
    if not date_text: return None
    m = re.search(r'\b(19|20)\d{2}\b', str(date_text))
    return m.group(0) if m else None

async def check_policy(district: DistrictRecord, policy: PolicyEntry, session, delay_min: int, delay_max: int) -> ScrapeResult:
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
        apply_policy_link(result, policy, None, via="BoardDocs API")
        if not policy.has_real_link:
            result.notes = "BoardDocs API returned no data"
            return result
        result.notes = "BoardDocs API returned no data, doing basic link check"
        import generic
        return await generic.check_link(district, policy)
        
    match = _find_matching_policy(policies, policy.policy_code)
    if not match:
        # The policy wasn't matched in the API listing. That alone is NOT proof
        # the on-file link is dead (the matcher can miss naming variants). If we
        # have a real link, verify it with an HTTP check before flagging red;
        # otherwise flag for review rather than assuming dead.
        apply_policy_link(result, policy, None, via="BoardDocs API")
        if policy.has_real_link:
            import generic
            return await generic.check_link(district, policy)
        result.action = ScapeAction.LINK_DEAD
        result.highlight_color = HighlightColor.RED
        result.notes = "Policy not found in BoardDocs API (no link on file)"
        return result
        
    rev_date = match.get("revised") or match.get("last_modified") or match.get("adopted")
    found_year = _extract_year(rev_date)
    apply_policy_link(result, policy, _boarddocs_policy_link(district, match), via="BoardDocs API")

    baseline_year = policy.max_year
    if found_year and baseline_year and int(found_year) > baseline_year:
        result.action = ScapeAction.REVISED
        result.highlight_color = HighlightColor.GREEN
        result.new_year_revised = found_year
        result.notes = f"Revised: {baseline_year} -> {found_year}"

    return result

async def search_for_policy(district: DistrictRecord, policy: PolicyEntry, session, delay_min: int, delay_max: int) -> ScrapeResult:
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
        if _is_safe_routes_target(policy.policy_code):
            return _safe_routes_not_found_result(result)
        result.action = ScapeAction.SKIPPED
        return result
        
    policies = await _fetch_policies_api(district.boarddocs_slug)
    if not policies:
        if policy.is_blank_block:
            return blank_not_found_result(
                district,
                policy,
                "BoardDocs API empty; normalized blank cells to 0/N/A",
            )
        if _is_safe_routes_target(policy.policy_code):
            return _safe_routes_not_found_result(result)
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
            result.new_year_adopted = found_year
            result.new_year_revised = found_year

        apply_policy_link(result, policy, _boarddocs_policy_link(district, match), via="BoardDocs API")
        result.notes = f"Newly found! Year: {found_year}"
    elif policy.is_blank_block:
        return blank_not_found_result(
            district,
            policy,
            "Policy not found in BoardDocs; normalized blank cells to 0/N/A",
        )
    elif _is_safe_routes_target(policy.policy_code):
        _safe_routes_not_found_result(result)

    return result
