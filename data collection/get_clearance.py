import asyncio
from playwright.async_api import async_playwright

async def get_clearance():
    print("Launching visible browser...")
    async with async_playwright() as pw:
        # We use the exact same persistent folder as the main scraper will use
        browser = await pw.chromium.launch_persistent_context(
            user_data_dir="./playwright_user_data",
            headless=False,  # VERY IMPORTANT: Keep this False so you can see it
            channel="chrome" # Use your actual Chrome browser to look more human
        )
        
        page = await browser.new_page()
        print("Navigating to Simbli. Please solve any CAPTCHA you see...")
        
        # Go to a known Simbli policy page
        await page.goto("https://simbli.eboardsolutions.com/Policy/PolicyListing.aspx?S=36030670")
        
        # Give you 60 seconds to solve any checks and verify the page loads
        print("You have 60 seconds to clear any checks and let the page load.")
        for i in range(60, 0, -5):
            print(f"  {i} seconds remaining...")
            await asyncio.sleep(5)
            
        print("Time is up! Saving cookies and closing browser...")
        await browser.close()
        print("Clearance saved. You can now run the scraper safely.")

if __name__ == "__main__":
    asyncio.run(get_clearance())
