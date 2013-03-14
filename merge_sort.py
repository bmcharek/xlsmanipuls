#!bin/python
"""
Klein scriptje om twee reeds gemanipuleerde xls te mergen
en te sorteren.
Kan natuurlijk uitgebreid worden tot een lijst van files.
"""
#TODO: Add specification for sheet number
import xlrd
import re
import xlwt
a_file = "adr.xls"
a_book = (xlrd.open_workbook(a_file))
a_sheet = a_book.sheet_by_index(0)
b_file = "adr.xls"
b_book = (xlrd.open_workbook(a_file))
b_sheet = b_book.sheet_by_index(0)
for counter in range(a_sheet.nrows):
    rowValue = b_sheet.row_values(counter, start_colx=0, end_colx=2)
    print rowValue
