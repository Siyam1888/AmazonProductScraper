import random
from bs4 import BeautifulSoup

import requests
from requests import Session
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

class ProductScraper:

    def __init__(self):
        self.HEADERS_LIST = [
            "Mozilla/5.0 (Windows; U; Windows NT 6.1; x64; fr; rv:1.9.2.13) Gecko/20101203 Firebird/3.6.13",
            "Mozilla/5.0 (Windows; U; Windows NT 6.1; rv:2.2) Gecko/20110201",
            "Opera/9.80 (X11; Linux i686; Ubuntu/14.10) Presto/2.12.388 Version/12.16",
            "Mozilla/5.0 (Windows NT 5.2; RW; rv:7.0a1) Gecko/20091211 SeaMonkey/9.23a1pre",
        ]

    def get_response(self, url):
        """Take a product URL and return the response"""

        # add schema to the url if it doesn't exist
        url = 'https://' + url if not url.startswith('http') else url
        
        # set up the request session
        session = Session()
        header = {
            "User-Agent": random.choice(self.HEADERS_LIST),
            "X-Requested-With": "XMLHttpRequest",
        }
        session.headers.update(header)

        # return the response
        try:
            response = session.get(url)
        except requests.exceptions.ConnectionError:
            print("Please Check your Internet Connection!")
            return None
        if response.ok:
            return response


    def scrape_product_info(self, url):
        response = self.get_response(url)
        soup = BeautifulSoup(response.content, 'html.parser')

        title = soup.find('span', {'id': 'productTitle'}).text.strip()
        description = soup.find('div', {'id': 'productDescription'}).p.span.text.strip()

        details_feature_div = soup.find('div', {'id': 'detailBullets_feature_div'})
        details_list = [' '.join(span.text.split()) for span in details_feature_div.find_all('span', {'class': 'a-list-item'})]
        details = ' | '.join(details_list)

        product_info = {'url': url,'title': title, 'description': description, 'details': details}
        return product_info



class Excel:
    """Work with all excel related tasks."""

    def __init__(self, file_name):
        self.file_name = file_name
        self.font = Font(color="000000", bold=True)
        self.bg_color = PatternFill(fgColor="E8E8E8", fill_type="solid")
        self.customize_excel()

    def create_sheets(self):
        """Create all the sheets required for the project."""
        self.output = (
            self.wb.create_sheet("Output")
            if "Output" not in self.wb.sheetnames
            else self.wb["Output"]
        )

        self.errors = (
            self.wb.create_sheet("Errors")
            if "Errors" not in self.wb.sheetnames
            else self.wb["Errors"]
        )


    def make_columns(self, cells_zip, sheet, general_width=20, url_width=50):
        """Takes zip values of rows and columns and puts values in place with some stylings"""
        # iterating through the column and its values to put them in place
        for col, value in cells_zip:
            cell = sheet[f"{col}1"]
            cell.value = value
            cell.font = self.font
            cell.fill = self.bg_color
            sheet.freeze_panes = cell

            # fixing the column width
            sheet.column_dimensions[col].width = general_width
        # fixing the URL column width
        sheet.column_dimensions["A"].width = url_width

    def customize_output_column(self):
        """Customize the Output column according to its values"""

        # combining columns with its values
        output_column = zip(
            ("A", "B", "C", "D"),
            (
                "Product URL",
                "Product Title",
                "Product Description",
                "Product Details",
            ),
        )
        self.make_columns(output_column, self.output, general_width=50)


    def customize_excel(self):
        """Run all the functions related to excel customization"""
        self.wb = load_workbook(self.file_name)
        self.create_sheets()
        self.customize_output_column()
        self.wb.save(self.file_name)

    def generate_inputs(self):
        """Read the first column of Input sheet and yield the values"""
        inputs = self.wb["Input"]
        for row in range(2, inputs.max_row + 1):
            # generates the links one by one
            if value := inputs[f"A{row}"].value:
                yield value

    def append_output(self, product_info):
        """Read the output and append it to excel"""
        if product_info:
            try:
                self.wb["Output"].append(
                    (
                        product_info["url"],
                        product_info["title"],
                        product_info["description"],
                        product_info["details"],
                    )
                )

                self.wb.save(self.file_name)
            except openpyxl.utils.exceptions.IllegalCharacterError as e:
                self.wb["Errors"].append((product_info["url"], str(e)))



        
    

if __name__ == '__main__':
    filename = 'AmazonProducts.xlsx'
    excel = Excel(filename)
    scraper = ProductScraper()

    for url in  excel.generate_inputs():
        print(f"SCRAPING... {url}")
        product_info = scraper.scrape_product_info(url)

        if product_info:
            excel.append_output(product_info)
    
    excel.wb.save(filename)
    
    