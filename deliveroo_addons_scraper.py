from playwright.sync_api import sync_playwright
import re
import openpyxl
from openpyxl.styles import Font
from credentials import vendor_name, vendor_url
import os

class DeliverooAddonScraper:
    def __init__(self, url, base_path, browser_context):
        self.url = url
        self.base_path = base_path
        self.browser_context = browser_context
        self.cat_attributes = []
        self.addon_attributes = []
        self.items_addons_attributes = []

    def append_to_excel(self, filename, data):
        """Append data to an Excel file, creating headers if file does not exist."""
        excel_path = f"{self.base_path}\\{filename}"
        try:
            workbook = openpyxl.load_workbook(excel_path)
            sheet = workbook.active
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            bold_font = Font(bold=True)
            sheet.append(list(data.keys()))
            for cell in sheet[1]:
                cell.font = bold_font

        sheet.append(list(data.values()))
        workbook.save(excel_path)
        
    def capitalize_sentence(self, sentence):
        return ' '.join(word.capitalize() for word in sentence.split())


    def dict_exists(self, target_dict, data_list, keys):
        """Check if a dictionary with specific keys exists in a list."""
        return any(all(d.get(key) == target_dict.get(key) for key in keys) for d in data_list)

    def extract_addon_details(self, addon):
        """Extract addon name and price from a given addon element."""
        addon_name = addon.query_selector('p.ccl-649204f2a8e630fd.ccl-a396bc55704a9c8a.ccl-0956b2f88e605eb8.ccl-40ad99f7b47f3781').text_content()
        try:                                
            addon_price  = addon.query_selector('div.ccl-a206e125970432e3').text_content()
        except:
            addon_price = ''
        return addon_name, addon_price

    def extract_addon_categories(self, page, category, selector):
        """Extract all addon categories and their items from the menu."""
        for item in category.query_selector_all(selector):
            try:
                item_name = item.query_selector('p.ccl-649204f2a8e630fd.ccl-a396bc55704a9c8a.ccl-0956b2f88e605eb8.ccl-ff5caa8a6f2b96d0.ccl-40ad99f7b47f3781').text_content().strip()
                item.click()
                page.wait_for_timeout(2000)

                addon_window = page.query_selector('div.ccl-e2683e5cd3d2680f')
                addon_categories = addon_window.query_selector_all("div.MenuItemModifiers-60c359b419ec39f6")

                for addon_category in addon_categories:
                    is_required = addon_category.query_selector("p.ccl-649204f2a8e630fd.ccl-6f43f9bb8ff2d712.ccl-08c109442f3e666d.ccl-40ad99f7b47f3781")
                    category_status = "No" if not is_required else "Yes"
                    
                    addon_category_name = self.capitalize_sentence(addon_category.query_selector("p.ccl-649204f2a8e630fd.ccl-a396bc55704a9c8a.ccl-0956b2f88e605eb8.ccl-ff5caa8a6f2b96d0.ccl-40ad99f7b47f3781").text_content().strip())
                    if not is_required:    
                        count_text = addon_category.query_selector("span.ccl-649204f2a8e630fd.ccl-6f43f9bb8ff2d712").text_content()
                        addon_count = re.search(r'\d+', count_text).group() if re.search(r'\d+', count_text) else \
                                      len(addon_category.query_selector_all('div.col-lg-5.col-md-5.col-sm-16.col-16'))
                    else:
                        addon_count = ''

                    cat_data = {'addon_category': addon_category_name, 'category_status': category_status, 'addon_count_line': addon_count}
                    if not self.dict_exists(cat_data, self.cat_attributes, cat_data.keys()):
                        self.cat_attributes.append(cat_data)

                    for addon in addon_category.query_selector_all('div.ccl-a5e1512b87ef2079'):
                        addon_name, addon_price = self.extract_addon_details(addon)
                        
                        item_data = {'item_name': item_name, 'addon_name': addon_name, 'addon_price': addon_price, 'category_status': category_status,}
                        if not self.dict_exists(item_data, self.items_addons_attributes, item_data.keys()):
                            self.items_addons_attributes.append(item_data)

                        addon_data = {'addon_category': addon_category_name, 'addon_name': addon_name, 'addon_price': addon_price, 'category_status': category_status}
                        if not self.dict_exists(addon_data, self.addon_attributes, addon_data.keys()):
                            self.addon_attributes.append(addon_data)

                addon_window.query_selector("button.ccl-4704108cacc54616.ccl-4f99b5950ce94015").click()
            except Exception as e:
                print(f"Error processing item: {e}")

    def scrape(self):
        """Scrape the menu from the given URL."""
        # with sync_playwright() as p:
        page = self.browser_context.new_page()
        page.goto(self.url)
        categories = page.query_selector_all("div.Layout-4549ebf43c78c99a")[2:]
        for category in categories:
            self.extract_addon_categories(page, category, 
                                          "div.MenuItemCard-a927b3314fc88b17")
  

    def save_to_excel(self):
        """Save scraped data to Excel files."""
        for data, filename in zip([self.cat_attributes, self.addon_attributes, self.items_addons_attributes],
                                  ["addon_cat.xlsx", "addons.xlsx", "items_addons.xlsx"]):
            for row in data:
                self.append_to_excel(filename, row)

    def start(self):
        """Start the scraping process."""
        self.scrape()
        self.save_to_excel()
        print("Scraping and saving complete!")



with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = browser.new_page( geolocation={"latitude": 25.186054760669197, "longitude": 55.27504936531868, "accuracy": 100}, permissions=["geolocation"])
    page.goto(vendor_url)
    page.wait_for_timeout(30000)

    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    desktop_path = desktop_path.replace("\\", "\\\\")

    scraper = DeliverooAddonScraper(
        url=vendor_url,
        base_path=f"{desktop_path}",
        browser_context=context
    )
    scraper.start()
    browser.close()    


