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


font = xlwt.Font()  # font
font.name = 'Times New Roman'
font.height = 11 * 20  # Нужный размер шрифта нужно умножить на 20
minifont = xlwt.Font()  # font
minifont.name = 'Times New Roman'
minifont.height = 8 * 20  # Нужный размер шрифта нужно умножить на 20
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
date_style.font = font

style = xlwt.XFStyle()
style.borders = borders
style.font = font

ministyle = xlwt.XFStyle()
ministyle.borders = borders
ministyle.font = minifont

fatStyle = xlwt.XFStyle()
fatStyle.borders = borders
fatStyle.font = fatFont

firstReg = 1  # Номер первой строки после шапки (счёт идёт от 0)

regColumn = 3  # Номер столбца с регномерами (счёт идёт от 0)

columnsWithDate = [10]  # Столбцы для которых нужен формат "ДАТА" - ЧЧ.ММ.ГГГГ

regNumber = set()
for row in range(firstReg, sheetR.nrows):
    if sheetR.cell_type(row, regColumn) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):  # Проверка пуста ли ячейка
        regNumber.add(sheetR.cell_value(row, regColumn))
        print(sheetR.cell_value(row, regColumn))

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
                        outSheet.col(col).width = 30*367

                if col in columnsWithDate:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=date_style)
                else:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=style)
            cnt += 1

    outBook.save(path + '\\' + str(val)[:-2] + '.xls')
    print(cont, 'из', len(regNumber), '   ', realpath(path + '\\' + str(val)[:-2] + '.xls'))
    cont += 1

startfile(path)
