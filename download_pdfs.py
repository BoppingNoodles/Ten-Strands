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

The download strategies are added in follow-up commits.

Usage
-----
    python download_pdfs.py --input "Summer_2026_Scraped_20260702_095536.xlsx" --sheet Caden
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from typing import Optional

import openpyxl

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


# ── Workbook predicates ───────────────────────────────────────────────────────

def _is_adopted(value) -> bool:
    """True if the policy value cell marks the policy as adopted (value == 1)."""
    if value is None:
        return False
    return str(value).strip() in {"1", "1.0"}


def _is_real_link(link) -> bool:
    """True if the link cell is a real http(s) URL (not N/A / * / blank)."""
    if not link:
        return False
    text = str(link).strip()
    if not text or text in {"N/A", "*", "None"}:
        return False
    return text.lower().startswith(("http://", "https://"))


# ── Workbook reading ──────────────────────────────────────────────────────────

def load_tasks(filepath: str, sheet: str,
               start_row: Optional[int] = None,
               end_row: Optional[int] = None) -> list[dict]:
    """
    Walk the sheet and return one task dict per adopted policy that has a real link.

    Each task: {row, county, district, policy_code, policy_title,
                year_adopted, year_revised, link}

    Rows 1-2 are headers; data starts at row 3 (matching the scraper's reader).
    Columns: A=status, B=CDS code, C=county, D=district, then per-policy quads
    (value | year_adopted | year_revised | link) per POLICY_DEFS.

    Only BP/AR policies are considered; resolution columns (RES-*) are skipped.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")
    ws = wb[sheet]

    tasks: list[dict] = []
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=3, values_only=True), start=3
    ):
        if start_row and row_idx < start_row:
            continue
        if end_row and row_idx > end_row:
            break

        county = str(row[2]).strip() if row[2] is not None else ""
        district = str(row[3]).strip() if row[3] is not None else ""
        if not district:
            continue

        for pdef in POLICY_DEFS:
            policy_code = pdef["code"]
            # Only BP/AR policies — skip RES-* resolution columns.
            if not re.match(r"^(BP|AR)\s", policy_code):
                continue

            base = pdef["col_start"] - 1  # 0-based index into the row tuple
            value = row[base]
            year_adopted = row[base + 1]
            year_revised = row[base + 2]
            link = row[base + 3]

            if not _is_adopted(value) or not _is_real_link(link):
                continue

            tasks.append({
                "row": row_idx,
                "county": county,
                "district": district,
                "policy_code": policy_code,
                "policy_title": pdef["title"],
                "year_adopted": year_adopted,
                "year_revised": year_revised,
                "link": str(link).strip(),
            })
    return tasks


# ── Entry point (download strategies land in later commits) ───────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Download a PDF for every adopted policy in a scraped workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--input", "-i", required=True,
                        help="Path to the scraped .xlsx workbook")
    parser.add_argument("--sheet", "-s", required=True,
                        help="Sheet name to process (e.g. Caden)")
    parser.add_argument("--output-dir", "-o", default=".",
                        help="Where to create the 'Ten Strands' parent folder")
    parser.add_argument("--limit", type=int, default=None,
                        help="Stop after this many download tasks (test runs)")
    parser.add_argument("--start-row", type=int, default=None,
                        help="First workbook data row to process (1-based)")
    parser.add_argument("--end-row", type=int, default=None,
                        help="Last workbook data row to process (1-based)")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        sys.exit(f"Error: workbook not found: {args.input}")

    print(f"Reading {args.input}  [sheet: {args.sheet}]")
    try:
        tasks = load_tasks(args.input, args.sheet,
                           start_row=args.start_row, end_row=args.end_row)
    except ValueError as exc:
        sys.exit(str(exc))

    if args.limit:
        tasks = tasks[: args.limit]

    if not tasks:
        print("No adopted policies with a real link found. Nothing to do.")
        return

    districts = []
    seen = set()
    for t in tasks:
        if t["district"] not in seen:
            seen.add(t["district"])
            districts.append(t["district"])

    print(f"Found {len(tasks)} adopted policy link(s) across {len(districts)} district(s).")

    base_dir = os.path.join(args.output_dir, PARENT_FOLDER)
    os.makedirs(base_dir, exist_ok=True)
    print(f"Output folder: {os.path.abspath(base_dir)}")

    print("\nPlanned downloads (download logic lands in the next commit):")
    for i, t in enumerate(tasks, 1):
        print(f"  [{i}/{len(tasks)}] {t['district']} | {t['policy_code']} -> {t['link']}")


if __name__ == "__main__":
    main()
