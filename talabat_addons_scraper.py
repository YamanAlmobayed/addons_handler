from playwright.sync_api import sync_playwright
import re
import openpyxl
from openpyxl.styles import Font

class TalabatAddonScraper:
    def __init__(self, url, base_path):
        self.url = url
        self.base_path = base_path
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
        label = addon.query_selector('label.control-label > span:nth-of-type(2)')
        text_span = addon.query_selector('span.text span')

        addon_name = (label or text_span).text_content().split('(')[0].strip() if label or text_span else None
        price_element = addon.query_selector('label[data-testid="radio"] span.currency') or \
                        addon.query_selector('label.control-label span.currency')
        addon_price = price_element.text_content().strip() if price_element else ''

        return addon_name, addon_price

    def extract_addon_categories(self, page, category, selector):
        """Extract all addon categories and their items from the menu."""
        for item in category.query_selector_all(selector):
            try:
                item_name = item.query_selector('div.f-15').text_content().strip()
                item.click()
                page.wait_for_timeout(2000)

                addon_window = page.query_selector('div.modal-content')
                addon_categories = addon_window.query_selector_all("div.sc-1bf12ad-0.ilBSTs")

                for addon_category in addon_categories:
                    has_checkboxes = addon_category.query_selector("div[data-testid='choices-checkboxes-component']")
                    category_status = "No" if has_checkboxes else "Yes"
                    
                    addon_category_name = self.capitalize_sentence(addon_category.query_selector("strong[data-test='sectionName']").text_content().strip())
                    count_text = addon_category.query_selector("span.dark-gray.align-middle").text_content()
                    addon_count = re.search(r'\d+', count_text).group() if re.search(r'\d+', count_text) else \
                                  len(addon_category.query_selector_all('div.col-lg-5.col-md-5.col-sm-16.col-16'))

                    cat_data = {'addon_category': addon_category_name, 'category_status': category_status, 'addon_count_line': addon_count}
                    if not self.dict_exists(cat_data, self.cat_attributes, cat_data.keys()):
                        self.cat_attributes.append(cat_data)

                    for addon in addon_category.query_selector_all('div.col-lg-5.col-md-5.col-sm-16.col-16'):
                        addon_name, addon_price = self.extract_addon_details(addon)
                        
                        item_data = {'item_name': item_name, 'addon_name': addon_name, 'addon_price': addon_price, 'category_status': category_status,}
                        if not self.dict_exists(item_data, self.items_addons_attributes, item_data.keys()):
                            self.items_addons_attributes.append(item_data)

                        addon_data = {'addon_category': addon_category_name, 'addon_name': addon_name, 'addon_price': addon_price, 'category_status': category_status}
                        if not self.dict_exists(addon_data, self.addon_attributes, addon_data.keys()):
                            self.addon_attributes.append(addon_data)

                addon_window.query_selector("span.clickable.close-span").click()
            except Exception as e:
                print(f"Error processing item: {e}")

    def scrape(self):
        """Scrape the menu from the given URL."""
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            page = browser.new_page()
            page.goto(self.url)

            categories = page.query_selector_all("div[data-testid='menu-category']")[1:]

            for category in categories:
                self.extract_addon_categories(page, category, 
                                              "div.sc-a31f9fb2-0.dyJtfK.d-flex.justify-content-between.py-2.clickable")
                self.extract_addon_categories(page, category, 
                                              "div.sc-a31f9fb2-0.eQGrrN.d-flex.justify-content-between.py-2.clickable")

            browser.close()

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


# Example usage:
scraper = TalabatAddonScraper(
    url="https://www.talabat.com/uae/restaurant/746408/AL-QUOZ-3?aid=1209",
    base_path="C:\\Users\\Yaman_Almobayed\\Desktop\\"
)
scraper.start()