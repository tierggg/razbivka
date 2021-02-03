#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from os import startfile
from os.path import realpath


filename = askopenfilename()
rb = xlrd.open_workbook(filename)
sheetR = rb.sheet_by_index(0)
path = askdirectory()

font = xlwt.Font()                      # font
font.name = 'Times New Roman'
font.height = 11 * 20                   # Нужный размер шрифта нужно умножить на 20

minifont = xlwt.Font()                  # font
minifont.name = 'Times New Roman'
minifont.height = 8 * 20                # Нужный размер шрифта нужно умножить на 20

borders = xlwt.Borders()                # borders
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1

style = xlwt.XFStyle()
style.borders = borders
style.font = font

date_style = xlwt.XFStyle()
date_style.num_format_str = "M/D/YY"
date_style.borders = borders
date_style.font = font

ministyle = xlwt.XFStyle()
ministyle.borders = borders
ministyle.font = minifont

# Ширина столбцов от первого до последнего, если ширина 0, столбец будет скрыт [3, 16, 26, 27, 24, 12, 4, 85, 7, 22, 13]
colsWidths = [0, 16, 26, 27, 24, 0, 0, 0, 0, 22, 13]
colDict = {}
for k in range(len(colsWidths)):
    colDict[k] = colsWidths[k]

count = 1

outBook = xlwt.Workbook()
outSheet = outBook.add_sheet(str(count))

for i in colDict:
    if colDict[i] == 0:
        outSheet.col(i).hidden = 1                  # Скрытие столбца
    else:
        outSheet.col(i).width = colDict[i] * 256    # Нужную ширину стобца надо умножить на 256

