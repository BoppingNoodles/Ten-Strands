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
python scrape_policies.py --input "Summer 2026 Board Policy Indicator Refresh Data Tracker.xlsx" --sheet Caden --inter-district-delay 8 --delay-min 3.5 --delay-max 6.5
```

Useful options:

- `--start-row 121 --end-row 200` — process a row range
- `--limit 5` or `--pilot` — small test run
- `--concurrency 3` — parallel districts (default 3)

> [!TIP]
> **Anti-Bot Settings:** Simbli has strict Cloudflare bot-protection. To prevent massive blocks of skipped districts, always use a slow, human-paced delay. 
> Recommended settings: `--inter-district-delay 8 --delay-min 3.5 --delay-max 6.5`

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

# Download district policy pdfs
Run this command (this will create a folder called 'Ten Strands' in your current directory and in that folder, download policy pdfs by policy). It will skip districts that have already been downloaded
```bash
python download_pdfs.py --input "Summer 2026 Board Policy Indicator Refresh Data Tracker.xlsx" --sheet Caden --start-row 3 --end-row 315 --chrome-version 149
```
---


### Run

From inside that `Ten Strands` folder:
`python reorganize_by_policy.py`
### Output

Creates a new folder called By Policy alongside the district folders inside Ten Strands. Inside By Policy, there's one subfolder per canonical policy (for example "AR 3511.1 Integrated Waste Management" or "BP 3510 Green Schools Operations"), and each of those subfolders contains copies of every district's PDF for that policy.

At the end of the run, the script prints how many PDFs were successfully copied out of the total found (e.g. `Copied 280/312 PDF(s)`), plus a list of any skipped files and why.
## Notes

- Simbli scraping opens a real Chrome window via `undetected-chromedriver`.
- Safe Routes policies (BP/AR 5142.2) are updated in the sheet but do not trigger green highlighting or the "Policy Updated" row tag.
- Districts marked with `*` (no database) are skipped during platform discovery and scraping.
