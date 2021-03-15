from selenium import webdriver
import re
from selenium.common.exceptions import NoSuchElementException
import sys
import sqlite3
import traceback
from datetime import datetime
from bs4 import BeautifulSoup
import requests
import io
from webdriver_manager.chrome import ChromeDriverManager
import urllib
from urllib.request import urlretrieve
import zipfile
import xlrd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


def fill_table():
    #sqlite_connection = sqlite3.connect("DNS_PARSE.db")
    #con = sqlite_connection.cursor()
    book = xlrd.open_workbook('price-tomsk.xls')
    sheets = book.sheet_names()
    sheets_count = book.nsheets
    worksheet = book.sheet_by_index(4)
    print(worksheet)
    print("Количество листов: ", sheets_count)
    # sheets[4] = компьютеры и комплектующие
    print("имя листа 4: ", sheets[4])

    #createTable = ("INSERT INTO product (kod,prod,M1,M2,M3,M4,M5,M6,M7,M8,M9,M10,M11,price,date)"
    #                        "VALUES (?,?,?,?,?,?)", gname, gprice_old, gprice_dis, gmarket, g_add_date, city)


#fill_table()



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
    section = "17a89a3916404e77/operativnaya-pamyat-dimm/"
    section = "17a89aab16404e77/videokarty/"
    url_start_page = "https://www.dns-shop.ru/catalog/" + section
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.maximize_window()
    driver.get(url_start_page)
    driver.implicitly_wait(3)
    driver.page_source.encode('utf-8')
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
#dns_parse()


def parse_soup():
    section = "17a89aab16404e77/videokarty/"
    url = 'https://www.dns-shop.ru/catalog/' + section#"https://www.dns-shop.ru/catalog/17a8dae116404e77/nastolnye-i-napolnye-svetilniki/"
    headers = headers = {'User-Agent': "App/0.0.1.1"}
    r = requests.get(url, headers=headers)
    r.encoding = "UTF-8"
    with open("test.txt", "w", encoding="UTF-8") as f:
        f.write(r.text)
    print(r.text)
    soup = BeautifulSoup(r.text, 'html.parser')
    print(soup.text)
#parse_soup()


def save_file():
    destination = "price_tomsk.zip"
    url = 'https://www.dns-shop.ru/files/price/price-tomsk.zip'
    urllib.request.urlretrieve(url, destination)
    print()
    return destination


#save_file()


def extract_file():
    file_zip = zipfile.ZipFile("price_tomsk.zip", "r")#(save_file())
    file_zip.extractall('')
    file_name = str(''.join(file_zip.namelist()))
    file_size = file_zip.getinfo(file_name).file_size
    print("Распакован файл: ", file_name)
    print("Размер файла: ", file_size/1000000, "Mb")
    file_zip.close()


#extract_file()


def xl_get():
    sqlite_connection = sqlite3.connect("DNS_PARSE.db")
    con = sqlite_connection.cursor()

    book = xlrd.open_workbook('price-tomsk.xls')
    #sheets = book.sheet_names()
    rows = book.sheet_by_index(4)
    #print(rows)
    #sheets_count = book.nsheets
    #worksheet = book.sheet_by_index(4)
    #print("Количество листов: ", sheets_count)
    #sheets[4] = компьютеры и комплектующие
    #print("имя листа 4: ", sheets[4])


    for r in range(13, rows.nrows):
        kod = rows.cell_value(r, 0)
        prod = rows.cell_value(r, 1)
        M1 = rows.cell_value(r, 2)
        M2 = rows.cell_value(r, 3)
        M3 = rows.cell_value(r, 4)
        M4 = rows.cell_value(r, 5)
        M5 = rows.cell_value(r, 6)
        M6 = rows.cell_value(r, 7)
        M7 = rows.cell_value(r, 8)
        M8 = rows.cell_value(r, 9)
        M9 = rows.cell_value(r, 10)
        M10 = rows.cell_value(r, 11)
        M11 = rows.cell_value(r, 12)
        price = rows.cell_value(r, 13)

        #print(rows.cell_value(r, 1))
        con.execute("INSERT INTO product (kod,prod,M1,M2,M3,M4,M5,M6,M7,M8,M9,M10,M11,price)"
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (kod,prod,M1,M2,M3,M4,M5,M6,M7,M8,M9,M10,M11,price))
        sqlite_connection.commit()
    con.close()
    sqlite_connection.close()
        # sqlite_connection = sqlite3.connect("DNS_PARSE.db")
        # con = sqlite_connection.cursor()
        #book = xlrd.open_workbook('price-tomsk.xls')
        #sheets = book.sheet_names()
        #sheets_count = book.nsheets
        ##worksheet = book.sheet_by_index(4)
        #print(worksheet)
        #print("Количество листов: ", sheets_count)
        # sheets[4] = компьютеры и комплектующие
        #print("имя листа 4: ", sheets[4])

        # createTable = ("INSERT INTO product (kod,prod,M1,M2,M3,M4,M5,M6,M7,M8,M9,M10,M11,price,date)"
        #                        "VALUES (?,?,?,?,?,?)", gname, gprice_old, gprice_dis, gmarket, g_add_date, city)

    #s = "Видеокарта"#"Видеокарта Asus AMD Radeon RX 550 Phoenix [PH-550-2G]"
    #reg = "*" + s + "*"
    #fs = re.findall(reg)
    #for sheet in book.sheets():
    #    for rowidx in range(sheet.nrows):
    #        row = sheet.row(rowidx)
    #        for colidx, cell in enumerate(row):
    #            if cell.value == (fs):
    #                print(sheet.name)
    #               print(colidx)
    #                print(rowidx)
    #                print(cell.value)



xl_get()