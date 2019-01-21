'''
Created on 2019年1月21日
@author: Vostory
'''
# 读取excel数据
# 小罗的需求，取第二行以下的数据，然后取每行前13列的数据

# -*- coding: utf-8 -*-
import xlwt
import xlrd

data = xlrd.open_workbook('D:\\PracFolder\\Data\\111.xlsx')  # 打开xls文件
table = data.sheets()[0]  # 打开第一张表
nrows = table.nrows  # 获取表的行数
for i in range(nrows):  # 循环逐行打印
    if i == 0:  # 跳过第一行
        continue
    print(table.row_values(i)[:2])  # 取前2列
