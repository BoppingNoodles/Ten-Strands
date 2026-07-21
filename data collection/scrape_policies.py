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
import discover
from models import ScrapeResult, ScapeAction, HighlightColor, blank_not_found_result, is_safe_routes_policy_code


def _is_safe_routes_policy(policy) -> bool:
    return is_safe_routes_policy_code(policy.policy_code)


def _safe_routes_not_found_result(district, policy, notes):
    return _not_found_result(district, policy, notes)


def _is_blank_policy_block(policy) -> bool:
    return policy.is_blank_block


def _not_found_result(district, policy, notes):
    return blank_not_found_result(district, policy, notes)

async def scrape_district(district, session, simbli_ctx, delay_min, delay_max):
    """Process all policies for a single district sequentially."""
    results = []

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
                    if _is_blank_policy_block(policy):
                        res = _not_found_result(
                            district, policy, "No system ID to search; normalized blank policy cells to 0/N/A"
                        )
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


class BrowserSession:
    def __init__(self):
        self.driver = self._create_driver()

    def _create_driver(self):
        import undetected_chromedriver as uc
        print("  [BrowserSession] Initializing Chrome...")
        options = uc.ChromeOptions()
        d = uc.Chrome(options=options, version_main=149)
        d.set_page_load_timeout(30)
        return d

    def ensure_alive(self):
        try:
            self.driver.current_url
            return self.driver
        except Exception:
            print("  [BrowserSession] Chrome crashed. Restarting...")
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = self._create_driver()
            return self.driver

    def __getattr__(self, name):
        return getattr(self.driver, name)

    def quit(self):
        try:
            self.driver.quit()
        except Exception:
            pass

async def main_async(args):
    print(f"Loading data from '{args.input}' sheet '{args.sheet}'...")
    districts = reader.load_districts(
        args.input, 
        args.sheet, 
        limit=args.limit,
        start_row=args.start_row,
        end_row=args.end_row
    )
    
    # Phase 2: For districts with no system ID and no links, try Google search
    # (via real Chrome browser to bypass anti-bot) to discover their Simbli or
    # BoardDocs platform. Must happen after the browser is initialized.

    print("Initializing browser for Simbli...")
    driver = BrowserSession()
    simbli_lock = asyncio.Lock()
    simbli_ctx = (driver, simbli_lock)

    await discover.discover_missing_platforms(districts, simbli_ctx)
    
    print(f"Found {len(districts)} districts. Proceeding with all {len(districts)} sequentially.")
    print(
        f"Pacing: delay={args.delay_min:.1f}-{args.delay_max:.1f}s between Simbli requests, "
        f"{args.inter_district_delay}s between districts"
    )
    
    async with AsyncSession(impersonate='chrome110') as session:
        results = []
        print("Starting scrape (sequential mode)...")
        for i, d in enumerate(districts):
            district_results = await scrape_district(d, session, simbli_ctx, args.delay_min, args.delay_max)
            results.append(district_results)
            if i < len(districts) - 1:
                print(f"  [Pause] Waiting {args.inter_district_delay}s before next district...")
                await asyncio.sleep(args.inter_district_delay)
        
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
    parser.add_argument("--concurrency", type=int, default=3, help="Number of parallel districts to process")
    parser.add_argument("--start-row", type=int, default=None, help="Start at row number (inclusive)")
    parser.add_argument("--end-row", type=int, default=None, help="End at row number (inclusive)")
    parser.add_argument("--inter-district-delay", type=int, default=3, help="Seconds to wait between districts")
    parser.add_argument("--delay-min", type=float, default=2.5, help="Min seconds between Simbli requests")
    parser.add_argument("--delay-max", type=float, default=5.5, help="Max seconds between Simbli requests")
    
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
