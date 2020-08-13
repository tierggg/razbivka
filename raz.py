#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
from tkinter.filedialog import askopenfilename
import subprocess
import os

filename = askopenfilename()
rb = xlrd.open_workbook(filename)
sheetR = rb.sheet_by_index(0)

font = xlwt.Font()               # font
font.name = 'Times New Roman'
font.height = 11 * 20            # Нужный размер шрифта нужно умножить на 20

borders = xlwt.Borders()         # borders
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1

date_style = xlwt.XFStyle()
date_style.num_format_str = "M/D/YY"
date_style.borders = borders
date_style.font = font
style = xlwt.XFStyle()
style.borders = borders
style.font = font

regNumber = set()
for row in range(3, sheetR.nrows):
    if sheetR.cell_type(row, 9) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):  # Проверка пуста ли ячейка
        regNumber.add(sheetR.cell_value(row, 9))

colsWidths = [3, 16, 26, 27, 24, 12, 4, 85, 7, 22, 13]  # ширина столбцов от 0 до последнего
colDict = {}

for k in range(len(colsWidths)):
    colDict[k] = colsWidths[k]

cont = 1

for val in regNumber:
    print(cont, 'из', len(regNumber))
    cnt = 2
    outBook = xlwt.Workbook()
    outSheet = outBook.add_sheet(str(int(val)))

    for i in colDict:
        outSheet.col(i).width = colDict[i] * 256    # Нужную ширину стобца надо умножить на 256

    for row in range(1, 2):
        for col in range(sheetR.ncols):
            outSheet.write(row, col, sheetR.cell_value(row, col), style=style)

    for row in range(3, sheetR.nrows):
        if val == sheetR.cell_value(row, 9):
            for col in range(sheetR.ncols):
                if col == 5:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=date_style)
                else:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=style)
            cnt += 1
    outBook.save('D:\\1\\out\\' + str(int(val)) + '.xls')
    cont += 1

path = 'D:\\1\\out\\'
path=os.path.realpath(path)
os.startfile(path)
