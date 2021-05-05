#!/usr/bin/env python
# -*- coding: utf-8 -*-

from xlrd import open_workbook
from xlrd import XL_CELL_EMPTY
from xlrd import XL_CELL_BLANK
from xlwt import Borders
from xlwt import XFStyle
from xlwt import Font
from xlwt import Workbook
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from os import startfile
from os.path import realpath

rb = open_workbook(askopenfilename())
path = askdirectory()
sheetR = rb.sheet_by_index(0)

borders = Borders()
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1

style = XFStyle()
style.borders = borders
normalFont = Font()
normalFont.name = 'Times New Roman'
normalFont.height = 11 * 20  # Нужный размер шрифта нужно умножить на 20
style.font = normalFont

date_style = XFStyle()
date_style.num_format_str = "M/D/YY"
date_style.borders = borders
date_style.font = normalFont

ministyle = XFStyle()
ministyle.borders = borders
miniFont = Font()
miniFont.name = 'Times New Roman'
miniFont.height = 8 * 20  # Нужный размер шрифта нужно умножить на 20
ministyle.font = miniFont

fatStyle = XFStyle()
fatStyle.borders = borders
fatFont = Font()
fatFont.name = 'Times New Roman'
fatFont.height = 11 * 20  # Нужный размер шрифта нужно умножить на 20
fatFont.bold = True
fatStyle.font = fatFont

firstReg = 1  # Номер первой строки после шапки (счёт идёт от 0)

regColumn = 3  # Номер столбца с рег.номерами (счёт идёт от 0)

columnsWithDate = [10]  # Номера столбцов, через запятую (счёт идёт от 0), для которых нужен формат "ДАТА" - ЧЧ.ММ.ГГГГ

regNumber = set()
for row in range(firstReg, sheetR.nrows):
    if sheetR.cell_type(row, regColumn) not in (XL_CELL_EMPTY, XL_CELL_BLANK):  # Проверка пуста ли ячейка
        regNumber.add(sheetR.cell_value(row, regColumn))
print('Найдено', len(regNumber), 'рег. номеров')

cont = 1  # Просто счётчик
for val in regNumber:
    rowCounter = firstReg
    outBook = Workbook()
    outSheet = outBook.add_sheet(str(val))

    for row in range(0, rowCounter):  # Заполнение шапки
        '''if row == 2:               # Номер строки в которой нужен мелкий шрифт
            for col in range(sheetR.ncols):
                outSheet.write(row, col, sheetR.cell_value(row, col), style=ministyle)
        else:'''
        for col in range(sheetR.ncols):
            outSheet.write(row, col, sheetR.cell_value(row, col), style=fatStyle)

    for row in range(rowCounter, sheetR.nrows):
        if val == sheetR.cell_value(row, regColumn):
            for col in range(sheetR.ncols):

                cwidth = outSheet.col(col).width
                if (len(str(sheetR.cell_value(row, col))) * 367) > cwidth:
                    if len(str(sheetR.cell_value(row, col))) < 175:
                        outSheet.col(col).width = len(str(sheetR.cell_value(row, col))) * 367
                    else:
                        outSheet.col(col).width = 30 * 367

                if col in columnsWithDate:
                    outSheet.write(rowCounter, col, sheetR.cell_value(row, col), style=date_style)
                else:
                    outSheet.write(rowCounter, col, sheetR.cell_value(row, col), style=style)
            rowCounter += 1

    if str(val)[-2:] == '.0':
        valname = str(val)[:-2]
    else:
        valname = str(val)

    outBook.save(path + '\\' + valname + '.xls')
    print(cont, 'из', len(regNumber), '   ', realpath(path + '\\' + valname + '.xls'))
    cont += 1

startfile(path)
