import re
import os
import configparser
from typing import List

from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
from RPA.Excel.Files import Files

from selenium.common.exceptions import NoSuchElementException


def parse_config(config_path: str) -> str:
    config = configparser.ConfigParser()
    config.read(config_path)
    return config["main"]["name_of_selected_agency"]

browser = Selenium()
name_of_selected_agency = parse_config("config.ini")
url = "https://itdashboard.gov"


def agency_link_search(agency_list: List) -> str:
    title_locator = "w200"
    amount_locator = "w900"
    btn_locator = "btn"
    for item in agency_list:
        title = item.find_element_by_class_name(title_locator).text
        if title == name_of_selected_agency:
            agency_link = item.find_element_by_class_name(btn_locator).get_property(
                "href"
            )
            break
    return agency_link


def create_main_table_with_headers(headers: List) -> List:
    table = []
    header_items = []

    for header in headers:
        header_items.append(header.text)
    table.append(header_items)

    return table


def create_exel_file(table: List):
    exel = Files()
    exel.create_workbook(f"./{name_of_selected_agency}.xlsx")
    exel.create_worksheet(name_of_selected_agency[:20], table)
    exel.save_workbook()


def select_all_investments(table_locator, browser):
    option_locator = "option"
    object_length = browser.find_element(table_locator)
    options = object_length.find_elements_by_tag_name(option_locator)
    for option in options:
        if option.text == "All":
            option.click()


def pdf_data_comparison(text, item):
    uii_number, name, link = item
    for number, text_part in text.items():
        pattern = "1. Name of this Investment\: ([\n\d()A-Za-z \,\-]+).*2. Unique Investment Identifier .UII.: ([\d\- ]+)"
        target = re.search(pattern, text_part)
        if target:
            name_from_pdf = target.group(1).strip()
            uii_from_pdf = target.group(2).strip()
            if uii_number == uii_from_pdf and name == name_from_pdf:
                print("Data match in ", uii_number)
            else:
                print("Data does not match in ", uii_number)
            break
    if not target:
        print("UII and Name data cannot be found")


def update_main_table_with_data(table, main_table):
    block_of_rows = table.find_element_by_tag_name("tbody")
    rows = block_of_rows.find_elements_by_tag_name("tr")
    pdf_links = []

    for row in rows:
        data_row = []
        values = row.find_elements_by_tag_name("td")
        for i in values:
            data_row.append(i.text)

        try:
            link_element = values[0].find_element_by_tag_name("a")
            uii = data_row[0]
            investment_name = data_row[2]
            link = [uii, investment_name, link_element.get_property("href")]
            pdf_links.append(link)
        except NoSuchElementException:
            pass
        main_table.append(data_row)
    return main_table, pdf_links


def pdf_work(pdf_links):
    try:
        os.mkdir("output")
    except:
        pass

    file_system = FileSystem()
    pdf = PDF()
    for item in pdf_links:
        uii_number, name, link = item
        try:
            browser = Selenium()
            browser.set_download_directory(os.path.abspath("output"))
            browser.open_available_browser(link)
            browser.wait_until_page_contains_element("id:business-case-pdf")
            browser.click_element("id:business-case-pdf")

            file_name = f"{uii_number}.pdf"
            pdf_save_path = os.path.abspath(os.path.join("output", file_name))
            file_system.wait_until_created(pdf_save_path, timeout=60 * 10)

            text = pdf.get_text_from_pdf(pdf_save_path)

            pdf_data_comparison(text, item)
        finally:
            browser.close_browser()


def agency_page_parse(agency_link: str, browser):
    table_locator = "name:investments-table-object_length"
    investment_table_locator = "id:investments-table-object_wrapper"
    all_check_locator = "css:div#investments-table-object_paginate span"

    browser.open_available_browser(agency_link)
    browser.wait_until_page_contains_element(table_locator, timeout=30)

    select_all_investments(table_locator, browser)

    browser.wait_until_element_does_not_contain(all_check_locator, "2", timeout=60)
    browser.wait_until_page_contains_element(investment_table_locator, timeout=50)

    table = browser.find_element(investment_table_locator)
    table_headers = table.find_element_by_class_name("dataTables_scrollHead")
    headers = table_headers.find_elements_by_tag_name("th")

    main_table = create_main_table_with_headers(headers)
    main_table, pdf_links = update_main_table_with_data(table, main_table)

    create_exel_file(main_table)
    pdf_work(pdf_links)


def agency_search():
    dive_btn = "class:btn.btn-default.btn-lg-2x.trend_sans_oneregular"
    agency_locator = "id:agency-tiles-widget"
    agency_titles_locator = "tuck-5"

    browser.wait_until_page_contains_element(dive_btn)
    browser.click_element(dive_btn)
    browser.wait_until_element_is_visible(agency_locator)

    agency_titles = browser.find_element(agency_locator)
    agencies = agency_titles.find_elements_by_class_name(agency_titles_locator)

    agency_link = agency_link_search(agencies)

    agency_page_parse(agency_link, browser)


def main():
    try:
        browser.open_available_browser(url)
        agency_search()
    finally:
        browser.close_all_browsers()


if __name__ == "__main__":
    main()
