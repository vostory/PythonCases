'''
功能：用xlrd和xlwt读写excel
Created on 2019年1月21日
@author: Vostory
'''
# coding=utf-8
import xlrd
import xlwt

# 读取数据
file_path = r'D:\\PracFolder\\Data\\111.xlsx'  # 路径前加 r，读取的文件路径
data = xlrd.open_workbook(file_path)  # 获取数据
table = data.sheet_by_name('Sheet1')  # 获取sheet
# table = data.sheet_by_index(0)#索引的方式，从0开始

nrows = table.nrows  # 获取总行数
ncols = table.ncols  # 获取总列数
row_value = table.row_values(2)  # 获取一行的数值，例如第5行
col_values = table.col_values(2)  # 获取一列的数值，例如第6列
cell_value = table.cell(2, 2).value  # 获取一个单元格的数值，例如第5行第6列
print(row_value)
print(col_values)
print(cell_value)

# 写入数据
writebook = xlwt.Workbook()  # 打开一个excel
test1 = writebook.add_sheet('test1')  # 在打开的excel中添加一个sheet，名称为test1

test1.write(0, 0, 'Englishname')  # 第0行第0列
test1.write(1, 0, 'Hellen')  # 第1行第0列
test1.write(0, 1, '中文名字')  # 第0行第1列
test1.write(1, 1, '海伦')  # 第1行第1列
writebook.save('D:\\PracFolder\\Data\\222.xlsx')  # 一定要记得保存
