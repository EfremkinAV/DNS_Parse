import re
import sys
import sqlite3
import traceback
from datetime import datetime
import io
import urllib
from urllib.request import urlretrieve
import zipfile
import xlrd
import time
import os
from shutil import copyfile

def save_file(city):
    city = "price-" + city + ".zip"
    url = 'https://www.dns-shop.ru/files/price/'+city
    check_file = os.path.exists(city)
    file_time = time.ctime(os.path.getctime(city))
    urllib.request.urlretrieve(url, city)
    #копирование файла из temp в основную папку. Нужно для сравнения файлов
    file_zip = zipfile.ZipFile(city, "r")
    file_zip.extractall('./temp/')
    file_name = str(''.join(file_zip.namelist()))
    file_size = file_zip.getinfo(file_name).file_size
    f1 = 'price-tomsk.xls'
    f2 = './temp/price-tomsk.xls'
    copyfile(f2, f1)
    print(time.ctime(os.path.getmtime(f1)))
    print(time.ctime(os.path.getmtime(f2)))
    if time.ctime(os.path.getmtime(f1)) != time.ctime(os.path.getmtime(f2)):
        copyfile(f2, f1)

        print("файл", city, "обновлен")
        print("Размер файла: ", file_size / 1000000, "Mb")
        print("Дата создания файла: ", file_time)
    file_zip.close()
    #return city


#save_file(input("Введите город:"))


#def file_time_diff(file_name):
#    file_name =




def fill_db():
    start_time = time.time()
    sqlite_connection = sqlite3.connect("DNS_PARSE.db")
    con = sqlite_connection.cursor()

    city = "tomsk"  # input("Введите город: ")
    file_name = "price-" + city + ".xls"
    book = xlrd.open_workbook(file_name)
    sheets_num = book.nsheets
    for i in range(1, sheets_num - 1):
        rows = book.sheet_by_index(i)
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
            parse_date = time.ctime(os.path.getctime("price-tomsk.xls")) #Добыть дату файла
            try:
                con.execute("INSERT INTO product (kod,prod,M1,M2,M3,M4,M5,M6,M7,M8,M9,M10,M11,price,date)"
                            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (kod, prod, M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11, price, parse_date))
                sqlite_connection.commit()
            except sqlite3.Error as error:
                print("Класс исключения: ", error.__class__)
                print("Исключение", error.args)
                print("Печать подробноcтей исключения SQLite: ")
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(traceback.format_exception(exc_type, exc_value, exc_tb))
    con.close()
    sqlite_connection.close()
    print("Таблица распарсилась за", str(time.time() - start_time), "секунд")
    print("---КОНЕЦ---")


fill_db()


def db_delete_table(table_name):
    sqlite_connection = sqlite3.connect('DNS_PARSE.db')
    con = sqlite_connection.cursor()
    con.execute("DELETE from " + table_name)
    sqlite_connection.commit()
    con.close()
    sqlite_connection.close()
    print("ТАБЛИЦА", table_name, "ОЧИЩЕНА!")

#db_delete_table("product")
