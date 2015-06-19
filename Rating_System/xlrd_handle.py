__author__ = 'kido'
import xlrd

data = xlrd.open_workbook('Rating.xlsx')

def get_info()

table = data.sheets()[1]
nrows = table.nrows
ncols = table.ncols

for i in range(nrows):
    print table.row_values(i)