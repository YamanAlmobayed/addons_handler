import pandas as pd
from playwright.sync_api import sync_playwright
import time

# Constants
LOGIN_URL = "https://vendor.elitewherego.com/login"
MAIN_URL = "https://vendor.elitewherego.com/items"
EXCEL_PATH = "C:\\Users\\Yaman_Almobayed\\Desktop\\items_addons.xlsx"

# Functions
def extract_items_with_addons(file_path):
    df = pd.read_excel(file_path, usecols=[0, 1, 2], names=['item_name', 'addon_name', 'addon_price'])
    return df.groupby('item_name')['addon_name'].apply(list).to_dict()

def normalize_text(text):
    return "".join(text.lower().split())

def scroll_down(page):
    while True:
        last_height = page.evaluate("document.body.scrollHeight")
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(2)
        new_height = page.evaluate("document.body.scrollHeight")
        if new_height == last_height:
            break

def wait_for_loading(page, timeout=3000):
    try:
        page.wait_for_function("document.documentElement.classList.contains('nprogress-busy')", timeout=timeout)
        page.wait_for_function("!document.documentElement.classList.contains('nprogress-busy')", timeout=timeout)
    except:
        time.sleep(2)

def login(page, username, password):
    page.goto(LOGIN_URL)
    page.fill("#email", username)
    page.fill("input[type='password']", password)
    page.click("button[type='submit']")
    page.wait_for_url('https://vendor.elitewherego.com/')

def process_items(page, items_data):
    page.goto(MAIN_URL)
    for item, addons in items_data.items():
        try:
            page.fill("#search", item)
            wait_for_loading(page)
            
            table = page.query_selector("table.w-full.whitespace-nowrap")
            if not table:
                continue

            edit_buttons = table.query_selector_all("a[href^='https://vendor.elitewherego.com/items/']")
            rows = table.query_selector_all("tr.hover\\:bg-gray-100.focus-within\\:bg-gray-100")

            item_links = [{
                "link": btn.get_attribute('href'),
                "title": rows[i].query_selector("td:nth-child(2)").text_content().strip(),
                "price": rows[i].query_selector("td:nth-child(3)").text_content().strip()
            } for i, btn in enumerate(edit_buttons)]
            for item_data in item_links:
                if normalize_text(item_data['title']) == normalize_text(item):
                    page.goto(item_data['link'])
                    page.wait_for_selector('tr.hover\\:bg-gray-100')
                    
                    addons_rows = page.query_selector_all('tr.hover\\:bg-gray-100')
                    for addon in addons:
                        for addon_row in addons_rows:
                            addon_name = addon_row.query_selector("td:nth-child(3)").text_content().strip()
                            if normalize_text(addon_name) == normalize_text(addon):
                                try:
                                    addon_row.query_selector('a[href^="https://vendor.elitewherego.com/addaddon/"]').click()
                                except:
                                    print(f'Addon "{addon}" already added.')
        except Exception as e:
            print(f"Error processing item '{item}':", e)
            page.goto(MAIN_URL)
            page.wait_for_selector('table.w-full.whitespace-nowrap')

def automate(username, password):
    items_data = extract_items_with_addons(EXCEL_PATH)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        login(page, username, password)
        process_items(page, items_data)
        
        browser.close()

# Execution
USERNAME = "foodbook@gmail.com"
PASSWORD = "100200300"
automate(USERNAME, PASSWORD)