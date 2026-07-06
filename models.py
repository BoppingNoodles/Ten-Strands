"""
models.py — Shared dataclasses and enums for the policy scraper.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class ScapeAction(str, Enum):
    """Result of a single policy scrape attempt."""
    REVISED        = "revised"        # New revision year found (1 → highlight green)
    UNCHANGED      = "unchanged"      # Policy exists, not revised since stored year
    LINK_DEAD      = "link_dead"      # Stored URL returns 4xx/5xx (highlight red)
    LINK_REDIRECT  = "link_redirect"  # URL 301/302'd to new location
    NEWLY_FOUND    = "newly_found"    # Was 0/N/A, policy now found on board site
    BOT_DETECTED   = "bot_detected"   # Anti-bot blocked us
    ERROR          = "error"          # Unexpected error
    SKIPPED        = "skipped"        # No searchable source available
    NO_DATABASE    = "no_database_skipped" # Policy value is '*'


class HighlightColor(str, Enum):
    GREEN  = "green"   # Policy revised
    RED    = "red"     # Link dead / archived
    NONE   = "none"    # No change to highlight


# ── Policy column definitions ────────────────────────────────────────────────

POLICY_DEFS: list[dict] = [
    {"code": "BP 3510",   "title": "Green Schools Operations",              "col_start": 5},
    {"code": "BP 3511",   "title": "Energy and Water Management",           "col_start": 9},
    {"code": "BP 3511.1", "title": "Integrated Waste Management",           "col_start": 13},
    {"code": "BP 3514",   "title": "Environmental Safety",                  "col_start": 17},
    {"code": "BP 3514.1", "title": "Hazardous Substances",                  "col_start": 21},
    {"code": "BP 5142.2", "title": "Safe Routes to School",                 "col_start": 25},
    {"code": "BP 6142.5", "title": "Environmental Education",               "col_start": 29},
    {"code": "BP 7110",   "title": "Facilities Master Plan",                "col_start": 33},
    {"code": "AR 3511.1", "title": "Integrated Waste Management Regulation","col_start": 37},
    {"code": "AR 3514",   "title": "Environmental Safety Regulation",       "col_start": 41},
    {"code": "AR 3514.1", "title": "Hazardous Substances",                  "col_start": 45},
    {"code": "AR 3514.2", "title": "Integrated Pest Management Regulation", "col_start": 49},
    {"code": "AR 5142.2", "title": "Safe Routes to School",                 "col_start": 53},
    {"code": "AR 7110",   "title": "Facilities Master Plan",                "col_start": 57},
    {"code": "RES-CLIMATE","title": "Climate Action or Climate Emergency Resolution", "col_start": 61},
    {"code": "RES-ENV-LIT","title": "Environmental or Climate Literacy Resolution",   "col_start": 65},
    {"code": "RES-EARTH", "title": "Earth Day Resolution",                  "col_start": 69},
    {"code": "RES-OTHER", "title": "Environmental or Climate Resolution (Other)",     "col_start": 73},
]

# Sentinels that mean "no policy / no database"
NO_DATA_SENTINELS = {"N/A", "*", None, ""}

SAFE_ROUTES_CODES = frozenset({"BP 5142.2", "AR 5142.2"})
SAFE_ROUTES_COL_STARTS = (25, 53)


def is_safe_routes_policy_code(policy_code: str) -> bool:
    return policy_code in SAFE_ROUTES_CODES


@dataclass
class PolicyEntry:
    """One policy quad (value / year_adopted / year_revised / link) for a district."""
    col_start: int          # 1-based column index of the value cell in the sheet
    policy_code: str        # e.g. "BP 3510"
    policy_title: str       # e.g. "Green Schools Operations"
    value: Optional[str]    # "1", "0", "*", "N/A", or None
    year_adopted: Optional[str]
    year_revised: Optional[str]
    link: Optional[str]

    @property
    def is_adopted(self) -> bool:
        """True if the policy is currently marked as adopted (value == '1' or 1)."""
        return str(self.value).strip() == "1" or str(self.value).strip() == "1.0"

    @property
    def is_no_database(self) -> bool:
        """True if the district has no board policy database at all."""
        return str(self.value).strip() == "*"

    @property
    def is_not_adopted(self) -> bool:
        """True if policy is not adopted (0, N/A, or blank)."""
        v = str(self.value).strip()
        return v in {"0", "0.0", "N/A", "None", ""}

    @property
    def has_real_link(self) -> bool:
        """True if the link is a real URL (not N/A, *, None, or plain text)."""
        if not self.link:
            return False
        l = str(self.link).strip()
        if l in {"N/A", "*", ""}:
            return False
        return l.startswith("http")

    @property
    def is_simbli(self) -> bool:
        return self.has_real_link and "simbli" in str(self.link).lower()

    @property
    def is_boarddocs(self) -> bool:
        return self.has_real_link and "boarddocs.com" in str(self.link).lower()

    @property
    def is_blank_block(self) -> bool:
        """True if all four policy cells (value/adopted/revised/link) are empty."""
        return all(
            value is None or str(value).strip() == ""
            for value in (self.value, self.year_adopted, self.year_revised, self.link)
        )

    @property
    def has_year_data(self) -> bool:
        """True if adopted or revised contains a parseable year."""
        return self.max_year is not None

    @property
    def max_year(self) -> Optional[int]:
        """Returns the most recent year (adopted or revised) as int, or None."""
        years = []
        for y in (self.year_adopted, self.year_revised):
            try:
                val = int(str(y).strip())
                if 1900 < val < 2100:
                    years.append(val)
            except (ValueError, TypeError):
                pass
        return max(years) if years else None


@dataclass
class DistrictRecord:
    """All data for one school district row in the tracker sheet."""
    row_index: int               # 1-based row number in the sheet
    cds_code: str
    county: str
    district_name: str
    tracking_status: Optional[str]
    simbli_id: Optional[str]     # e.g. "36030670"
    boarddocs_slug: Optional[str]# e.g. "castrovalley"
    policies: list[PolicyEntry] = field(default_factory=list)


@dataclass
class ScrapeResult:
    """The outcome of scraping a single district-policy pair."""
    cds_code: str
    district_name: str
    policy_code: str
    action: ScapeAction
    highlight_color: HighlightColor
    notes: str = ""

    # Old values (from spreadsheet)
    old_value: Optional[str] = None
    old_year_revised: Optional[str] = None
    old_link: Optional[str] = None

    # New values (from scrape) — None means "no change"
    new_value: Optional[str] = None
    new_year_adopted: Optional[str] = None
    new_year_revised: Optional[str] = None
    new_link: Optional[str] = None

    # Source column info (for writer)
    col_start: int = 0

    def append_note(self, note: str) -> None:
        if note:
            self.notes = f"{self.notes}; {note}" if self.notes else note

    def to_dict(self) -> dict:
        return {
            "cds_code":       self.cds_code,
            "district_name":  self.district_name,
            "policy_code":    self.policy_code,
            "action":         self.action.value,
            "highlight_color":self.highlight_color.value,
            "notes":          self.notes,
            "old_value":      self.old_value,
            "old_year_revised":self.old_year_revised,
            "old_link":       self.old_link,
            "new_value":      self.new_value,
            "new_year_adopted":self.new_year_adopted,
            "new_year_revised":self.new_year_revised,
            "new_link":       self.new_link,
            "col_start":      self.col_start,
        }


def blank_not_found_result(
    district: DistrictRecord,
    policy: PolicyEntry,
    notes: str,
) -> ScrapeResult:
    """Normalize a blank policy block to 0 / N/A / N/A / N/A."""
    return ScrapeResult(
        cds_code=district.cds_code,
        district_name=district.district_name,
        policy_code=policy.policy_code,
        action=ScapeAction.UNCHANGED,
        highlight_color=HighlightColor.NONE,
        notes=notes,
        old_value=policy.value,
        old_year_revised=policy.year_revised,
        old_link=policy.link,
        new_value="0",
        new_year_adopted="N/A",
        new_year_revised="N/A",
        new_link="N/A",
        col_start=policy.col_start,
    )


def apply_policy_link(
    result: ScrapeResult,
    policy: PolicyEntry,
    new_link: Optional[str],
    via: str = "index",
) -> None:
    """
    Update result.new_link when a scraped URL is available, or clear invalid
    placeholder links when no URL can be resolved.
    """
    if policy.has_real_link:
        if policy.has_year_data and new_link and str(policy.link).strip() != new_link:
            result.new_link = new_link
            result.append_note(f"Link updated (via {via})")
        return

    link_text = str(policy.link).strip() if policy.link is not None else ""
    if (not link_text or link_text in {"N/A", "*"}) and new_link:
        result.new_link = new_link
        result.append_note(f"Link added (via {via})")
        return

    if not link_text or link_text in {"N/A", "*"}:
        return

    if new_link:
        result.new_link = new_link
        result.append_note(f"Link updated (via {via})")
        return

    result.new_link = "N/A"
    if policy.has_year_data:
        result.append_note("Invalid link replaced with N/A; no URL found")
    else:
        result.append_note("Invalid link cleared; no year data")


def is_bad_link(value: Optional[str]) -> bool:
    """True if link cell is non-empty but not a real URL or allowed sentinel."""
    if value is None or str(value).strip() == "":
        return False
    text = str(value).strip()
    if text in {"N/A", "*"}:
        return False
    return not text.startswith(("http://", "https://"))


def has_year_data(adopted: Optional[str], revised: Optional[str]) -> bool:
    for year in (adopted, revised):
        try:
            val = int(str(year).strip())
            if 1900 < val < 2100:
                return True
        except (ValueError, TypeError):
            pass
    return False

