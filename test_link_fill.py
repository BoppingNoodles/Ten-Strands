"""
Internal test for apply_policy_link fix and a 3-district live scrape dry-run.
Does not write any .xlsx output.
"""

import asyncio
import sys
from curl_cffi.requests import AsyncSession

import reader
import discover
from models import (
    PolicyEntry,
    ScrapeResult,
    ScapeAction,
    HighlightColor,
    apply_policy_link,
)
from scrape_policies import scrape_district


def test_apply_policy_link_unit() -> None:
    scraped = "https://simbli.eboardsolutions.com/Policy/ViewPolicy.aspx?S=123&revid=abc"

    cases = [
        ("N/A link + scraped URL", PolicyEntry(5, "BP 3510", "Test", "0", "N/A", "N/A", "N/A"), scraped, scraped),
        ("blank link + scraped URL", PolicyEntry(5, "BP 3510", "Test", "0", "N/A", "N/A", ""), scraped, scraped),
        ("N/A link + no scraped URL", PolicyEntry(5, "BP 3510", "Test", "0", "N/A", "N/A", "N/A"), None, None),
        ("garbage link + scraped URL", PolicyEntry(5, "BP 3510", "Test", "1", "2018", "2020", "see website"), scraped, scraped),
    ]

    print("=== Unit tests: apply_policy_link ===")
    failed = 0
    for name, policy, incoming, expected in cases:
        result = ScrapeResult(
            cds_code="0000000",
            district_name="Test District",
            policy_code=policy.policy_code,
            action=ScapeAction.NEWLY_FOUND,
            highlight_color=HighlightColor.GREEN,
            new_value="1",
            new_year_adopted="2020",
            new_year_revised="2020",
            col_start=policy.col_start,
        )
        apply_policy_link(result, policy, incoming, via="test")
        ok = result.new_link == expected
        status = "PASS" if ok else "FAIL"
        print(f"  [{status}] {name}: new_link={result.new_link!r}")
        if not ok:
            failed += 1

    if failed:
        raise SystemExit(f"{failed} unit test(s) failed")


async def test_three_districts_live(rows: list[int] | None = None) -> None:
    input_path = "Summer 2026 Board Policy Indicator Refresh Data Tracker.xlsx"
    sheet = "Caden"

    if rows:
        districts = []
        for row in rows:
            districts.extend(
                reader.load_districts(input_path, sheet, start_row=row, end_row=row)
            )
    else:
        districts = reader.load_districts(input_path, sheet, limit=3)
    if not districts:
        raise SystemExit("No districts loaded for live test")

    print("\n=== Live dry-run: 3 districts (no workbook written) ===")
    for d in districts:
        not_adopted = [p for p in d.policies if p.is_not_adopted]
        print(f"  {d.district_name} (row {d.row_index}): {len(not_adopted)} not-adopted policies")

    import undetected_chromedriver as uc

    options = uc.ChromeOptions()
    driver = uc.Chrome(options=options, version_main=149)
    driver.set_page_load_timeout(30)
    simbli_lock = asyncio.Lock()
    simbli_ctx = (driver, simbli_lock)

    try:
        await discover.discover_missing_platforms(districts, simbli_ctx)

        sem = asyncio.Semaphore(3)
        async with AsyncSession(impersonate="chrome110") as session:
            tasks = [
                scrape_district(d, session, simbli_ctx, sem, delay_min=2.5, delay_max=5.5)
                for d in districts
            ]
            all_results = await asyncio.gather(*tasks, return_exceptions=True)

        newly_found = []
        link_backfilled = []
        missing_link_with_scrape_note = []
        for district_results in all_results:
            if isinstance(district_results, Exception):
                print(f"  ERROR: {district_results}")
                continue
            for res in district_results:
                if res.action == ScapeAction.NEWLY_FOUND:
                    newly_found.append(res)
                if res.new_link and (res.old_link is None or str(res.old_link).strip() in {"", "N/A", "None"}):
                    link_backfilled.append(res)
                if res.action == ScapeAction.NEWLY_FOUND and res.new_value == "1" and not res.new_link:
                    missing_link_with_scrape_note.append(res)

        print(f"\n  NEWLY_FOUND policies: {len(newly_found)}")
        print(f"  Link backfilled from N/A/blank: {len(link_backfilled)}")
        for res in newly_found:
            link_status = "OK" if res.new_link else "MISSING"
            print(
                f"    [{link_status}] {res.district_name} / {res.policy_code}: "
                f"value={res.new_value}, adopted={res.new_year_adopted}, "
                f"revised={res.new_year_revised}, link={res.new_link!r}"
            )
        for res in link_backfilled[:10]:
            if res not in newly_found:
                print(
                    f"    [BACKFILL] {res.district_name} / {res.policy_code}: "
                    f"action={res.action.value}, link={res.new_link!r}"
                )

        if missing_link_with_scrape_note:
            print(
                f"\n  WARNING: {len(missing_link_with_scrape_note)} newly-found policy(ies) "
                "still missing link (likely no URL from API/index, not apply_policy_link)"
            )
        elif newly_found:
            print("\n  All newly-found policies have links populated.")
        else:
            print("\n  No newly-found policies in this sample (fix verified via unit tests).")

    finally:
        driver.quit()


def main() -> None:
    test_apply_policy_link_unit()
    if "--live" in sys.argv:
        # Districts with many N/A not-adopted policies for end-to-end coverage.
        target_rows = [204, 205, 217]
        asyncio.run(test_three_districts_live(rows=target_rows))
    else:
        print("\nSkipping live scrape (pass --live to run 3-district browser test).")


if __name__ == "__main__":
    main()
