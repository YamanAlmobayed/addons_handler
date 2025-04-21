import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import os
from credentials import email, password, vendor_name

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
desktop_path = desktop_path.replace("\\", "\\\\")

LOGIN_URL = "https://vendor.elitewherego.com/login"
MAIN_URL = "https://vendor.elitewherego.com/addons"
EXCEL_PATH = f"{desktop_path}\\{vendor_name}\\addons\\addons.xlsx"

def read_excel_to_dict_list(file_path: str) -> list:
    df = pd.read_excel(file_path, usecols=[0, 1, 2], names=["addon_category", "addon_name", "addon_price"], skiprows=0, header=[0])
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
                await page.click("a[href='https://vendor.elitewherego.com/addons/create']")
                await page.wait_for_selector('#name')
                await page.wait_for_selector('#price')
                await page.wait_for_selector('#status')
                await page.wait_for_selector('#addon_category_id')
                
                await page.fill("#name", row['addon_name'])
                await page.fill("#price", str(row['addon_price']))
                await page.select_option("#addon_category_id", f"{row['addon_category']}")
                await page.select_option("#status", "Active")
                
                await page.click("button[type='submit']")
                await page.wait_for_selector('table.w-full.whitespace-nowrap')
            except Exception as e:
                print("Error:", e)
                
                await page.goto(MAIN_URL)
                await page.wait_for_selector('table.w-full.whitespace-nowrap')
                continue
        
        await browser.close()

asyncio.run(automate(email, password))
