import pandas as pd
from playwright.sync_api import sync_playwright
import time
import re
import os
from credentials import email, password, vendor_name

class ItemsAddonsLinker:
    LOGIN_URL = "https://vendor.elitewherego.com/login"
    MAIN_URL = "https://vendor.elitewherego.com/items"

    def __init__(self, username, password, excel_path):
        self.username = username
        self.password = password
        self.excel_path = excel_path

    @staticmethod
    def sanitize_text(text: str, allow_spaces: bool = True) -> str:
        allowed_chars = r"[^a-zA-Z0-9_\-.\\ ]"
        sanitized = re.sub(allowed_chars, "", text)
        sanitized = re.sub(r"\s+", " ", sanitized).strip()
        if not allow_spaces:
            sanitized = sanitized.replace(" ", "_")
        return sanitized

    @staticmethod
    def normalize_text(text):
        return "".join(text.lower().split())

    @staticmethod
    def scroll_to_element(page, selector):
        page.eval_on_selector(selector, "el => el.scrollIntoView({ behavior: 'smooth', block: 'center' })")

    @staticmethod
    def wait_for_loading(page, timeout=3000):
        try:
            page.wait_for_function("document.documentElement.classList.contains('nprogress-busy')", timeout=timeout)
            page.wait_for_function("!document.documentElement.classList.contains('nprogress-busy')", timeout=timeout)
        except:
            time.sleep(2)
    @staticmethod
    def format_number(value):
        if isinstance(value, (int, float)):
            if float(value).is_integer():
                return str(int(value))
            else:
                return str(value)
        else:
            raise ValueError("Input must be an int or float")

    def extract_items_with_addons(self):
        df = pd.read_excel(self.excel_path, usecols=[0, 1, 2, 3, 4, 5],
                           names=['category_name', 'item_name', 'addon_category', 'addon_name', 'addon_price', 'category_status'])
        result = {}
        for _, row in df.iterrows():
            item = row['item_name']
            addon_details = [row['category_name'], row['addon_category'], row['addon_name'], self.format_number(row['addon_price']), row['category_status']]
            result.setdefault(item, []).append(addon_details)
        return result

    def login(self, page):
        page.goto(self.LOGIN_URL)
        page.fill("#email", self.username)
        page.fill("input[type='password']", self.password)
        page.click("button[type='submit']")
        page.wait_for_url('https://vendor.elitewherego.com/')

    def process_items(self, page, items_data):
        page.goto(self.MAIN_URL)

        for item_name, attributes in items_data.items():
            try:
                sanitized_item = self.sanitize_text(item_name)
                page.fill("#search", sanitized_item)
                self.wait_for_loading(page)

                table = page.query_selector("table.w-full.whitespace-nowrap")
                if not table:
                    continue

                item_links = self.extract_item_links(table)
                for item_link in item_links:
                    if self.is_matching_item(item_link, item_name, attributes[0][0]):
                        self.process_addons_for_item(page, item_link["link"], attributes)
                        break  # If matched and processed, no need to check more links

                page.goto(self.MAIN_URL)

            except Exception as e:
                print(f"[Item Error] '{item_name}': {e}")
                page.goto(self.MAIN_URL)
                page.wait_for_selector('table.w-full.whitespace-nowrap')

    def extract_item_links(self, table):
        edit_buttons = table.query_selector_all("a[href^='https://vendor.elitewherego.com/items/']")
        rows = table.query_selector_all("tr.hover\\:bg-gray-100.focus-within\\:bg-gray-100")
        return [
            {
                "link": btn.get_attribute('href'),
                "category": rows[i].query_selector("td:nth-child(1)").text_content().strip(),
                "title": rows[i].query_selector("td:nth-child(2)").text_content().strip(),
                "price": rows[i].query_selector("td:nth-child(3)").text_content().strip(),
            }
            for i, btn in enumerate(edit_buttons)
        ]

    def is_matching_item(self, item_link, item_name, expected_category):
        return (
            self.sanitize_text(self.normalize_text(item_link['title'])) == self.sanitize_text(self.normalize_text(item_name)) and
            self.sanitize_text(self.normalize_text(item_link['category'])) == self.sanitize_text(self.normalize_text(expected_category))
        )

    def process_addons_for_item(self, page, item_url, addon_attributes):
        page.goto(item_url)
        page.wait_for_selector('tr.hover\\:bg-gray-100')
        page.wait_for_timeout(3000)

        # Step 1: Cache all current addon rows
        addon_rows = page.query_selector_all('tr.hover\\:bg-gray-100')
        addon_data = []

        # Step 2: Parse addon rows into usable dicts
        for row in addon_rows:
            try:
                data = {
                    "row": row,
                    "category": self.sanitize_text(self.normalize_text(row.query_selector("td:nth-child(2)").text_content().strip())),
                    "status": row.query_selector("td:nth-child(3)").text_content().strip(),
                    "name": self.sanitize_text(self.normalize_text(row.query_selector("td:nth-child(4)").text_content().strip())),
                    "price": row.query_selector("td:nth-child(5)").text_content().strip(),
                    "link": row.query_selector('a[href^="https://vendor.elitewherego.com/addaddon/"]').get_attribute("href")
                }
                addon_data.append(data)
            except Exception as e:
                print(f"[Parse Addon Row Error] {e}")

        # Step 3: Search for matching addon from cached data
        for addon_attr in addon_attributes:
            for addon in addon_data:
                if (addon["category"] == self.sanitize_text(self.normalize_text(addon_attr[1])) and
                    addon["name"] == self.sanitize_text(self.normalize_text(addon_attr[2])) and
                    addon["price"] == str(addon_attr[3]) and
                    addon["status"] == addon_attr[4]):

                    try:
                        print(f"Found matching addon, visiting: {addon['link']}")
                        page.goto(addon["link"])
                        page.wait_for_selector('tr.hover\\:bg-gray-100')
                        page.wait_for_timeout(1000)
                        break  # move to next addon_attr
                    except Exception as e:
                        print(f"[Addon Goto Error] {e}")
                    break

    # def is_matching_addon(self, row, addon_attr):
    #     addon_category = self.sanitize_text(self.normalize_text(row.query_selector("td:nth-child(2)").text_content().strip()))
    #     addon_status = row.query_selector("td:nth-child(3)").text_content().strip()
    #     addon_name = self.sanitize_text(self.normalize_text(row.query_selector("td:nth-child(4)").text_content().strip()))
    #     addon_price = row.query_selector("td:nth-child(5)").text_content().strip()

    #     return (
    #         addon_category == self.sanitize_text(self.normalize_text(addon_attr[1])) and
    #         addon_name == self.sanitize_text(self.normalize_text(addon_attr[2])) and
    #         addon_price == str(addon_attr[3]) and
    #         addon_status == addon_attr[4]
    #     )

    def run(self):
        items_data = self.extract_items_with_addons()
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context()
            page = context.new_page()

            self.login(page)
            self.process_items(page, items_data)

            browser.close()

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
desktop_path = desktop_path.replace("\\", "\\\\")

if __name__ == "__main__":
    EXCEL_PATH = f"{desktop_path}\\{vendor_name}\\addons\\items_addons.xlsx"

    automation = ItemsAddonsLinker(email, password, EXCEL_PATH)
    automation.run()