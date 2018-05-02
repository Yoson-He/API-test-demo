# -*- coding: utf-8 -*-
# @Time    : 2018/4/11 13:56
# @Author  : Yoson
# @File    : learn.py
# @Software: PyCharm

import xlrd
a=[]

data = xlrd.open_workbook(r'C:\Users\Administrator\Downloads\03_按订单查看明细.xlsx')
table = data.sheet_by_index(0)
rows = table.nrows
for i in range(1,rows):
   a.append(table.cell_value(i, 3))

b=[]
data2 = xlrd.open_workbook(r'C:\Users\Administrator\Downloads\导出订单(2018-04-11).xlsx')
table2 = data2.sheet_by_index(0)
rows2 = table2.nrows
for i in range(1,rows2):
   b.append(table2.cell_value(i, 4))

for each in a:
    if each not in b:
        print(each)
print('---------')
for each in b:
    if each not in a:
        print(each)