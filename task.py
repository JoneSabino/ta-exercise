from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem
import pandas as pd
import os
import time as t

fs = FileSystem()
browser = Selenium()
excel = Files()
tb = Tables()

URL = "https://itdashboard.gov"
AGENCY = os.getenv('CHOSEN_AGENCY')
print(os.getenv('PATH'))
agency = f'//span[contains(text(),"{AGENCY}")]/..'
workbook_path = './output/agency_info.xlsx'


def access_itdashboard(url):
    prefs = {'download.default_directory': './output/'}
    browser.open_available_browser(url, headless=True, preferences=prefs)


def create_excel_file(path):
    excel.create_workbook()
    excel.rename_worksheet('Sheet', 'Agencies')
    excel.create_worksheet('Induvidual Investments')
    excel.save_workbook(path)


def write_to_excel(data, workbook, worksheet):
    try:
        excel.open_workbook(workbook)
        excel.append_rows_to_worksheet(data, worksheet)
        excel.save_workbook(workbook)
        excel.close_workbook()
    except Exception as e:
        raise f'{e} \n Failed when trying to generate \
               the excel file'


def _get_agencies_info(selector, web_elements):
    return [x.find_element_by_css_selector(selector).text
            for x in web_elements]


def get_agencies_and_spend_amounts():
    dive_in = 'css:a[aria-controls="home-dive-in"]'
    agencies = '//span[@class=" h1 w900"]/..'
    browser.click_element_when_visible(dive_in)

    browser.wait_until_element_is_visible(agencies)

    agencies_we = browser.get_webelements(agencies)
    sheet_data = {'agency': '', 'amount': ''}
    sheet_data['agency'] = _get_agencies_info('.h4.w200', agencies_we)
    sheet_data['amount'] = _get_agencies_info('.h1.w900', agencies_we)

    return sheet_data


def main():
    try:
        # Step 1:
        #   - Get all the agencies and its spend amounts.
        #   - Write it to an Excel file with sheet name "agencies"
        create_excel_file(workbook_path)
        access_itdashboard(URL)
        agencies_info = get_agencies_and_spend_amounts()
        write_to_excel(agencies_info, workbook_path, 'Agencies')

        # Step 2:
        #   - Load an agency from a file.
        #   - Go to the agency's page
        #   - Get all "Individual Investments"
        #   - Write to a new sheet in the same excel file
        browser.wait_until_element_is_visible(agency)
        agency_we = browser.get_webelement(agency)
        # for attr in agency_we.get_property('attributes'):
        #     print(f"{attr['name']}: {attr['value']}")
        agency_ln = agency_we.get_attribute('href')
        browser.go_to(agency_ln)

        show_entries = 'css:select[name="investments-table-object_length"]'
        browser.wait_until_element_is_visible(show_entries, timeout='30s')
        browser.select_from_list_by_value(show_entries, '-1')

        next_btn_disabled = 'css:#investments-table-object_next.disabled'
        browser.wait_until_element_is_visible(next_btn_disabled, timeout='30s')

        data_table_loc = 'css:#investments-table-object'
        data_table_we = browser.get_webelement(data_table_loc)
        df = pd.read_html(data_table_we.get_attribute("outerHTML"),
                          flavor='html5lib')[0]
        write_to_excel(df.values.tolist(),
                       workbook_path,
                       'Induvidual Investments')

        # Step 3:
        #   If the "UII" column contains a link, open it and download PDF with
        #   Business Case (button "Download Business Case PDF")
        uii_loc = 'css:#investments-table-object > \
                   tbody > tr > td.left.sorting_2 > a'
        uii_we = browser.get_webelements(uii_loc)
        uii_lns = [uui.get_attribute('href') for uui in uii_we]
        for uii in uii_lns:
            print(uii)
            split_char = "/"
            filename = uii.split(split_char)[-1]
            browser.go_to(uii)
            browser.click_element_when_visible('id:business-case-pdf')
            fs.wait_until_created(f'./output/{filename}.pdf', 30)
    finally:
        browser.close_all_browsers()


def get_agencies_expenses():
    print("Done.")


if __name__ == "__main__":
    main()

