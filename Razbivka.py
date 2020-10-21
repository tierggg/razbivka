#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from os import startfile
from os.path import realpath


font = xlwt.Font()                      # font
font.name = 'Times New Roman'
font.height = 11 * 20
borders = xlwt.Borders()                # borders
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1
style = xlwt.XFStyle()
style.borders = borders
style.font = font

filename = askopenfilename()
rb = xlrd.open_workbook(filename)
sheetR = rb.sheet_by_index(0)
path = askdirectory()

regNumbers = set()
for row in range(1, sheetR.nrows):
    #if sheetR.cell_type(row, 9) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):  # Проверка пуста ли ячейка
    regNumbers.add(sheetR.cell_value(row, 0))

print(regNumbers)
cont = 1
for regnumber in regNumbers:
    cnt = 1
    outBook = xlwt.Workbook()
    outSheet = outBook.add_sheet(str(regnumber))

    for col in range(sheetR.ncols):
        outSheet.write(0, col, sheetR.cell_value(0, col), style=style)

    for row in range(1, sheetR.nrows):
        if regnumber == sheetR.cell_value(row, 0):
            for col in range(sheetR.ncols):
                outSheet.write(cnt, col, sheetR.cell_value(row, col), style=style)
            cnt += 1
    outBook.save(path + '\\' + str(regnumber) + '.xls')
    print(cont, 'из', len(regNumbers), '   ', realpath(path + '\\' + str(regnumber) + '.xls'))
    cont += 1
startfile(path)
