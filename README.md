# Ten-Strands Policy Scraper

Scrapes California school district board policies from Simbli and BoardDocs, compares them against a tracking spreadsheet, and writes an updated Excel workbook with revision highlights.

## Setup

1. Install Python 3.11+.
2. Install [Google Chrome](https://www.google.com/chrome/) (required by `undetected-chromedriver`).
3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run the scraper

```bash
python scrape_policies.py --input "Summer 2026 Board Policy Indicator Refresh Data Tracker.xlsx" --sheet Caden
```

Useful options:

- `--start-row 121 --end-row 200` — process a row range
- `--limit 5` or `--pilot` — small test run
- `--concurrency 5` — parallel districts (default 5)

Output files are created automatically:

- `Summer_2026_Scraped_YYYYMMDD_HHMMSS.xlsx` — updated workbook
- `scrape_log_YYYYMMDD_HHMMSS.json` — scrape details

## Files to share

### Required (core scraper)

| File | Purpose |
|------|---------|
| `scrape_policies.py` | Main entry point / orchestrator |
| `models.py` | Shared data models and policy column definitions |
| `reader.py` | Reads the input Excel tracker into district records |
| `writer.py` | Applies scrape results back to the workbook |
| `simbli.py` | Simbli scraper (Chrome + PolicyListing API) |
| `boarddocs.py` | BoardDocs API scraper |
| `generic.py` | Fallback HTTP link checker |
| `discover.py` | Discovers Simbli/BoardDocs IDs via Google search |
| `requirements.txt` | Python dependencies |

### Input data (not in git)

Each user needs their own copy of the tracking spreadsheet (for example `Summer 2026 Board Policy Indicator Refresh Data Tracker.xlsx`). It is excluded from git because it is project data, not code.

### Optional helpers

| File | Purpose |
|------|---------|
| `normalize_blank_policy_blocks.py` | Fill blank policy quads with `0 / N/A / N/A / N/A` |
| `strip_safe_routes_tracking.py` | Remove Safe Routes (BP/AR 5142.2) highlights and spurious "Policy Updated" tags |
| `inspect_workbook_policy_issues.py` | Scan a workbook for blank blocks and invalid links |
| `verify_scrape_output.py` | Compare a scrape log against a written workbook |

### Analysis

| File | Purpose |
|------|---------|
| `analysis/compare_policy_language.py` | Compare district policy PDFs with templates or cluster by language similarity |

### Not needed to run the scraper

- `Summer_2026_Scraped_*.xlsx`, `scrape_log_*.json` — generated output
- `playwright_user_data/` — local browser cache
- `test_*.py`, `debug_*.html` — development/debug scripts
- `make_district_folders.py`, `rename_pdfs.py` — unrelated utilities

## Notes

- Simbli scraping opens a real Chrome window via `undetected-chromedriver`.
- Safe Routes policies (BP/AR 5142.2) are updated in the sheet but do not trigger green highlighting or the "Policy Updated" row tag.
- Districts marked with `*` (no database) are skipped during platform discovery and scraping.
