#coding=utf-8
from xlrd import open_workbook
from xlutils.copy import copy
import xlrd
from xlutils.copy import copy
import pandas as pd
import numpy as np

#打开一个workbook
xlrd.Book.encoding = "gbk"
workbook = xlrd.open_workbook('d:/test/test.xlsx','w+b')
workbooknew = copy(workbook)
#抓取所有sheet页的名称
worksheets = workbook.sheet_names()
#定位到sheet1
worksheet1 = workbook.sheet_by_name(u'Sheet1')

#遍历sheet1中所有行row
num_rows = worksheet1.nrows
for curr_row in range(num_rows):
    row = worksheet1.row_values(curr_row)
    #遍历sheet1中所有列col
    num_cols = worksheet1.ncols
    for curr_col in range(num_cols):
        col = worksheet1.col_values(curr_col)
    #遍历sheet1中所有单元格cell
    for rown in range(num_rows):
        for coln in range(num_cols):
            cell = worksheet1.cell_value(rown,coln)
            if  '要替换的字符串'in cell:
                ws = workbooknew.get_sheet(0)
                ws.write(rown, coln, '替换后的字符串')
                workbooknew.save(u'C:/Users/yimei.wen/Desktop/test.xls')
