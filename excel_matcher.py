import logging, openpyxl, pprint
import sqlite3
import pandas as pd


def column_counter(xls_filename): #checks row 1 and counts how many cells have data, therefore how many columns in xls
    logging.debug('Counting columns in xls')
    wb = openpyxl.load_workbook(xls_filename)
    sheet = wb.active
    cols = 1 # start from 1st column
    while 1:
        cell = sheet.cell(row = 1, column = cols)
        v = cell.value
        if v != None: # if there is data in the cell
            cols += 1 # check the next column along
        else:    # if no data in the cell, then that's the last column, so break
            break
    c = int(cols)-1  # need to be minus one because it increments cols, then realises it's an empty cell
    logging.debug(f'# cols = {c}')
    return c


def row_counter(xls_filename): #checks column 1 and counts how many cells have data, therefore how many rows in xls
    logging.debug('Counting rows in xls')
    wb = openpyxl.load_workbook(xls_filename)
    sheet = wb.active
    rows = 1  # start from 1st column
    while 1:
        check_cell = sheet.cell(row = rows, column = 1)
        v = check_cell.value
        if v != None:  #if there is data in the cell
            rows += 1  # check the next column along
        else:    # if no data in the cell, then that's the last row, so break
             break
    r = int(rows)-1 # need to be minus one because it increments rows, then realises it's an empty cell
    logging.debug(f'#rows = {r}')
    return r


def excel_headings_grabber(xls_filename): # checks row 1 of xls and returns a dictionary showing col# & heading
    logging.debug('excel_headings_grabber - establishing headings/columns dict')
    wb = openpyxl.load_workbook(xls_filename)
    sheet = wb.active
    cols = 1
    dic = {}  # start from 1st column
    while 1:
        check_cell = sheet.cell(row = 1, column = cols)
        v = check_cell.value
        if v != None:  # if there is data in the cell
            dic.setdefault(cols, v)
            cols += 1 # check the next column along
        else:    # if no data in the cell, then that's the last column, so break
            break
    return dic


def xls_to_sql(filename):
    # print(f'attempting to open {filename + ".db"}')
    con = sqlite3.connect(filename + ".db")
    wb = pd.read_excel(filename + '.xlsx', sheetname=None)
    for sheet in wb:
        wb[sheet].to_sql(sheet, con, index=False)
    con.commit()
    con.close()


# excel_filename = "H:\WorkingDir\member_data\All members"
# xls_to_sql(excel_filename)


