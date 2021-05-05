#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from os import startfile
from os.path import realpath

rb = xlrd.open_workbook(askopenfilename())
path = askdirectory()
sheetR = rb.sheet_by_index(0)


normalFont = xlwt.Font()
normalFont.name = 'Times New Roman'
normalFont.height = 11 * 20  # Нужный размер шрифта нужно умножить на 20
miniFont = xlwt.Font()
miniFont.name = 'Times New Roman'
miniFont.height = 8 * 20  # Нужный размер шрифта нужно умножить на 20
fatFont = xlwt.Font()
fatFont.name = 'Times New Roman'
fatFont.height = 11 * 20
fatFont.bold = True

borders = xlwt.Borders()  # borders
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1

date_style = xlwt.XFStyle()
date_style.num_format_str = "M/D/YY"
date_style.borders = borders
date_style.font = normalFont

style = xlwt.XFStyle()
style.borders = borders
style.font = normalFont

ministyle = xlwt.XFStyle()
ministyle.borders = borders
ministyle.font = miniFont

fatStyle = xlwt.XFStyle()
fatStyle.borders = borders
fatStyle.font = fatFont

firstReg = 1  # Номер первой строки после шапки (счёт идёт от 0)

regColumn = 2  # Номер столбца с регномерами (счёт идёт от 0)

columnsWithDate = [10]  # Столбцы для которых нужен формат "ДАТА" - ЧЧ.ММ.ГГГГ

regNumber = set()
for row in range(firstReg, sheetR.nrows):
    if sheetR.cell_type(row, regColumn) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):  # Проверка пуста ли ячейка
        regNumber.add(sheetR.cell_value(row, regColumn))
        # print(sheetR.cell_value(row, regColumn))

print('Найдено', len(regNumber), 'рег. номеров')

cont = 1
for val in regNumber:
    cnt = firstReg
    outBook = xlwt.Workbook()
    outSheet = outBook.add_sheet(str(val))

    for row in range(0, cnt):  # Заполнение шапки
        '''if row == 2:
            for col in range(sheetR.ncols):
                outSheet.write(row, col, sheetR.cell_value(row, col), style=ministyle)
        else:'''
        for col in range(sheetR.ncols):
            outSheet.write(row, col, sheetR.cell_value(row, col), style=fatStyle)

    for row in range(cnt, sheetR.nrows):
        if val == sheetR.cell_value(row, regColumn):
            for col in range(sheetR.ncols):

                cwidth = outSheet.col(col).width
                if (len(str(sheetR.cell_value(row, col))) * 367) > cwidth:
                    if len(str(sheetR.cell_value(row, col))) < 175:
                        outSheet.col(col).width = len(str(sheetR.cell_value(row, col))) * 367
                    else:
                        outSheet.col(col).width = 30 * 367

                if col in columnsWithDate:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=date_style)
                else:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=style)
            cnt += 1

    if str(val)[-2:] == '.0':
        valname = str(val)[:-2]
    else:
        valname = str(val)

    outBook.save(path + '\\' + valname + '.xls')
    print(cont, 'из', len(regNumber), '   ', realpath(path + '\\' + valname + '.xls'))
    cont += 1

startfile(path)
