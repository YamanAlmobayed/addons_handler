import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import os 
from credentials import email, password, vendor_name

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
desktop_path = desktop_path.replace("\\", "\\\\")

LOGIN_URL = "https://vendor.elitewherego.com/login"
MAIN_URL = "https://vendor.elitewherego.com/addoncategories"
EXCEL_PATH = f"{desktop_path}\\{vendor_name}\\addons\\addon_cat.xlsx"

def read_excel_to_dict_list(file_path: str) -> list:
    df = pd.read_excel(file_path, usecols=[0, 1, 2], names=["category", "category_status", "count"], skiprows=0, header=[0])
    return df.to_dict(orient='records')

async def automate(USERNAME, PASSWORD):
    excel_data = read_excel_to_dict_list(EXCEL_PATH)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        await page.goto(LOGIN_URL)
        await page.fill("#email", USERNAME)
        await page.fill("input[type='password']", PASSWORD)
        await page.click("button[type='submit']")
        await page.wait_for_url('https://vendor.elitewherego.com/')
        
        await page.goto(MAIN_URL)
        
        for row in excel_data:
            try:
                await page.click("a[href='https://vendor.elitewherego.com/addoncategories/create']")
                await page.wait_for_selector('div.max-w-3xl.overflow-hidden.bg-white.rounded.shadow')
                
                await page.fill("#name", row['category'])
                await page.select_option("#status", "Active")
                await page.select_option("#is_required", row['category_status'])
                if(row['category_status'] == 'No'):
                    await page.fill("#count", str(row['count']))
                
                await page.click("button[type='submit']")
                await page.wait_for_selector('table.w-full.whitespace-nowrap')
            except Exception as e:
                print("Error:", e)
                await page.goto(MAIN_URL)
                await page.wait_for_selector('table.w-full.whitespace-nowrap')
                continue
        
        await browser.close()

asyncio.run(automate(email, password))
