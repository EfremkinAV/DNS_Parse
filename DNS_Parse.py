from selenium import webdriver
import re
from selenium.common.exceptions import NoSuchElementException
import sys
import sqlite3
import traceback
from datetime import datetime
import bs4


def db_conn():
    sqlite_connection = sqlite3.connect('DNS_PARSE.db')
    con = sqlite_connection.cursor()
    try:
        con.execute("INSERT INTO product (model,type) VALUES ('te','st')")
        sqlite_connection.commit()
    except sqlite3.Error as err:
        print("Класс исключения: ", err.__class__)
        print("Исключение", err.args)
        print("Печать подробноcтей исключения SQLite: ")
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(traceback.format_exception(exc_type, exc_value, exc_tb))
    con.close()
    sqlite_connection.close()
#db_conn()


def dns_parse():
    name = "" #название товара которое надо найти
    #section = "17a89a3916404e77/operativnaya-pamyat-dimm/"
    section = "17a89aab16404e77/videokarty/"
    url_start_page = "https://www.dns-shop.ru/catalog/" + section
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(url_start_page)
    driver.implicitly_wait(3)
    last_page = driver.find_element_by_class_name('pagination-widget__page-link_last ').get_attribute('href')
    last_page = re.findall('\d+$', last_page)
    last_page = int(''.join(last_page))
    #print(last_page)
    product_card = driver.find_elements_by_css_selector(".catalog-products.view-simple .catalog-product")
    #print(product_card)
    for specs in product_card:
        spec_model = specs.find_element_by_css_selector(".catalog-product__name").text
        print(spec_model)
    driver.quit()
dns_parse()