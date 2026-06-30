"""
scrape_policies.py — Orchestrator for the async scraping process.
"""

import asyncio
import argparse
from datetime import datetime
from curl_cffi.requests import AsyncSession

import reader
import writer
import simbli
import boarddocs
import generic
from models import ScrapeResult, ScapeAction, HighlightColor

async def scrape_district(district, session, simbli_ctx, sem, delay_min, delay_max):
    """Process all policies for a single district."""
    results = []
    
    async with sem:
        print(f"[{district.district_name}] Starting policy processing...")
        try:
            for i, policy in enumerate(district.policies):
                print(f"[{district.district_name}] Checking policy {i+1}/{len(district.policies)}: {policy.policy_code} (System: Simbli={policy.is_simbli}, BoardDocs={policy.is_boarddocs})")
                if policy.is_no_database:
                    res = ScrapeResult(
                        cds_code=district.cds_code,
                        district_name=district.district_name,
                        policy_code=policy.policy_code,
                        action=ScapeAction.NO_DATABASE,
                        highlight_color=HighlightColor.NONE,
                        notes="Skipped due to * indicator",
                        col_start=policy.col_start
                    )
                    results.append(res)
                    continue

                if policy.is_adopted:
                    # Check for revisions
                    if policy.is_simbli or (policy.has_real_link == False and district.simbli_id):
                        res = await simbli.check_policy(district, policy, simbli_ctx, delay_min, delay_max)
                    elif policy.is_boarddocs or (policy.has_real_link == False and district.boarddocs_slug):
                        res = await boarddocs.check_policy(district, policy, session, delay_min, delay_max)
                    elif policy.has_real_link:
                        res = await generic.check_link(district, policy)
                    else:
                        res = ScrapeResult(
                            cds_code=district.cds_code, district_name=district.district_name,
                            policy_code=policy.policy_code, action=ScapeAction.SKIPPED,
                            highlight_color=HighlightColor.NONE, notes="No link, no system ID",
                            col_start=policy.col_start
                        )
                    results.append(res)
                    
                elif policy.is_not_adopted:
                    # Check if newly adopted
                    if district.simbli_id:
                        res = await simbli.search_for_policy(district, policy, simbli_ctx, delay_min, delay_max)
                    elif district.boarddocs_slug:
                        res = await boarddocs.search_for_policy(district, policy, session, delay_min, delay_max)
                    else:
                        res = ScrapeResult(
                            cds_code=district.cds_code, district_name=district.district_name,
                            policy_code=policy.policy_code, action=ScapeAction.SKIPPED,
                            highlight_color=HighlightColor.NONE, notes="No system ID to search",
                            col_start=policy.col_start
                        )
                    results.append(res)
        except Exception as e:
            print(f"[{district.district_name}] Error: {e}")
            
    return results

async def main_async(args):
    print(f"Loading data from '{args.input}' sheet '{args.sheet}'...")
    districts = reader.load_districts(args.input, args.sheet, limit=args.limit)
    
    # Filter out districts that are completely * / N/A with no links
    valid_districts = []
    skipped_districts = 0
    for d in districts:
        has_system = bool(d.simbli_id or d.boarddocs_slug)
        has_any_link = any(p.has_real_link for p in d.policies)
        if not has_system and not has_any_link:
            skipped_districts += 1
        else:
            valid_districts.append(d)
            
    print(f"Found {len(districts)} districts. Skipped {skipped_districts} (no DB/links). Proceeding with {len(valid_districts)}.")
    
    # Increase concurrency since HTTP requests are lightweight compared to Playwright tabs
    # but don't increase too much to avoid rate limits.
    sem = asyncio.Semaphore(args.concurrency * 2) 
    
    import undetected_chromedriver as uc
    print("Initializing browser for Simbli...")
    options = uc.ChromeOptions()
    driver = uc.Chrome(options=options, version_main=149)
    driver.set_page_load_timeout(30)
    simbli_lock = asyncio.Lock()
    simbli_ctx = (driver, simbli_lock)
    
    async with AsyncSession(impersonate='chrome110') as session:
        tasks = [
            scrape_district(d, session, simbli_ctx, sem, args.delay_min, args.delay_max) 
            for d in valid_districts
        ]
        
        print("Starting scrape...")
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
    print("Closing browser...")
    driver.quit()
        
    print(f"Scrape complete. Writing output to {args.output} and log to {args.log}...")
    writer.write_output(args.input, results, args.output, args.sheet)
    writer.write_log(results, args.log)
    print("Done.")

def main():
    parser = argparse.ArgumentParser(description="Policy Tracker Scraper")
    parser.add_argument("--input", default="Summer 2026 Board Policy Indicator Refresh Data Tracker.xlsx")
    parser.add_argument("--sheet", default="Caden")
    parser.add_argument("--pilot", action="store_true", help="Run only first 5 districts")
    parser.add_argument("--limit", type=int, default=None, help="Process max N districts")
    parser.add_argument("--concurrency", type=int, default=5, help="Max parallel districts")
    parser.add_argument("--delay-min", type=int, default=1)
    parser.add_argument("--delay-max", type=int, default=3)
    
    args = parser.parse_args()
    
    if args.pilot:
        args.limit = 5
        print("PILOT MODE: Limiting to first 5 districts.")
        
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    args.output = f"Summer_2026_Scraped_{ts}.xlsx"
    args.log = f"scrape_log_{ts}.json"
    
    asyncio.run(main_async(args))

if __name__ == "__main__":
    main()
