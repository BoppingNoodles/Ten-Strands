"""
download_pdfs.py — Download a PDF for every adopted policy in a scraped workbook.

Input  : a .xlsx workbook produced by the scraper (e.g. Summer_2026_Scraped_*.xlsx)
Output : a ``Ten Strands`` parent folder in the current directory, with one
         subfolder per district, containing the downloaded policy PDFs.

Naming convention (one file per adopted policy):
    <BP/AR>_<Policy Number>_<Policy Title>_<County>_<District>_<Year Adopted>_<Year Revised>.pdf
Example:
    BP_3510_Green Schools Operations_Alameda_Alameda Unified_2019_2019.pdf

Only BP/AR policy quads are considered. Resolution columns (RES-*) are skipped.

(Workbook reading and the download strategies are added in follow-up commits.)

Usage
-----
    python download_pdfs.py --input "Summer_2026_Scraped_20260702_095536.xlsx" --sheet Caden
"""

from __future__ import annotations

import argparse
import re
import sys
from typing import Optional

# Column layout (value | adopted | revised | link quads) is shared with the
# scraper so the two never drift apart.
from models import POLICY_DEFS


# ── Constants ─────────────────────────────────────────────────────────────────

PARENT_FOLDER = "Ten Strands"


# ── Filename helpers ──────────────────────────────────────────────────────────

def strip_county_suffix(county: str) -> str:
    """'Alameda County' -> 'Alameda'  (case-insensitive)."""
    return re.sub(r"\s*county\s*$", "", (county or "").strip(), flags=re.IGNORECASE)


def sanitize(value: str) -> str:
    """Remove characters that are illegal in file-system names and tidy spaces."""
    value = (value or "").strip()
    value = re.sub(r'[<>:"/\\|?*]', " ", value)
    value = re.sub(r"\s{2,}", " ", value)
    return value.strip()


def normalize_year(value) -> str:
    """'2019', 2019, 2019.0 -> '2019'; anything else -> 'NA'."""
    if value is None:
        return "NA"
    text = str(value).strip()
    if not text or text in {"N/A", "NA", "*", "None"}:
        return "NA"
    try:
        year = int(float(text))
        return str(year) if 1900 < year < 2100 else "NA"
    except (ValueError, TypeError):
        return "NA"


def parse_doc_type_and_number(policy_code: str) -> tuple[str, str]:
    """
    'BP 3510'    -> ('BP', '3510')
    'AR 3514.1'  -> ('AR', '3514.1')
    'RES-CLIMATE'-> ('', '')          (resolutions are skipped by the caller)
    """
    code = (policy_code or "").strip()
    m = re.match(r"^(BP|AR)\s+(.+)$", code, re.IGNORECASE)
    if m:
        return m.group(1).upper(), m.group(2).strip()
    return "", ""


def build_filename(policy_code: str, policy_title: str, county: str,
                   district: str, year_adopted, year_revised) -> Optional[str]:
    """Assemble the convention filename stem, or None if it cannot be built."""
    doc_type, policy_number = parse_doc_type_and_number(policy_code)
    if not doc_type or not policy_number:
        # Resolution or unknown code — caller filters these out, but be safe.
        return None
    parts = [
        doc_type,
        sanitize(policy_number),
        sanitize(policy_title) or "Untitled",
        sanitize(strip_county_suffix(county)),
        sanitize(district),
        normalize_year(year_adopted),
        normalize_year(year_revised),
    ]
    return "_".join(parts) + ".pdf"


# ── Entry point (stub — fleshed out in later commits) ─────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Download a PDF for every adopted policy in a scraped workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--input", "-i", required=True,
                        help="Path to the scraped .xlsx workbook")
    parser.add_argument("--sheet", "-s", required=True,
                        help="Sheet name to process (e.g. Caden)")
    parser.parse_args()
    print("Workbook reading and download logic land in the next commits.")


if __name__ == "__main__":
    main()
