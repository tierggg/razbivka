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

bigfont = xlwt.Font()                      # font
bigfont.name = 'Times New Roman'
bigfont.height = 16 * 20

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

bigstyle = xlwt.XFStyle()
bigstyle.font = bigfont

filename = askopenfilename()
rb = xlrd.open_workbook(filename)
sheetR = rb.sheet_by_index(0)
path = askdirectory()

regNumbers = set()
for row in range(1, sheetR.nrows):
    #if sheetR.cell_type(row, 9) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):  # Проверка пуста ли ячейка
    regNumbers.add(sheetR.cell_value(row, 0))

# Ширина столбцов от первого до последнего, если ширина 0, столбец будет скрыт [3, 16, 26, 27, 24, 12, 4, 85, 7, 22, 13]
colsWidths = [8, 14, 40, 10, 8, 10, 25, 25]
colDict = {}
for k in range(len(colsWidths)):
    colDict[k] = colsWidths[k]

cont = 1
for regnumber in regNumbers:
    cnt = 4
    outBook = xlwt.Workbook()
    outSheet = outBook.add_sheet(str(int(regnumber)))

    for i in colDict:
        if colDict[i] == 0:
            outSheet.col(i).hidden = 1                  # Скрытие столбца
        else:
            outSheet.col(i).width = colDict[i] * 256    # Нужную ширину стобца надо умножить на 256

    outSheet.write(0, 0, 'ПЕНСИОНЕРЫ, которых нет в СЗВ-М за ОКТЯБРЬ 2020', style=bigstyle)
    outSheet.write(0, 6, 'Написать Объяснения в колонке ПРИМЕЧАНИЕ', style=bigstyle)
    outSheet.write(1, 0, 'Образец: работа по ГПХ в сентябре 20; уволен в сентябре-дата, СЗВ-ТД сдан(дата); расторгнут',
                   style=bigstyle)
    outSheet.write(2, 0, 'трудовой договор по совместительству, СЗВ-ТД сдан(дата)',style=bigstyle)

    for col in range(sheetR.ncols):
        outSheet.write(3, col, sheetR.cell_value(0, col), style=style)

    for row in range(1, sheetR.nrows):
        if regnumber == sheetR.cell_value(row, 0):
            for col in range(sheetR.ncols):
                if col == 3 or col == 5:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=date_style)
                else:
                    outSheet.write(cnt, col, sheetR.cell_value(row, col), style=style)
            cnt += 1
    outBook.save(path + '\\' + str(regnumber) + '.xls')
    print(cont, 'из', len(regNumbers), '   ', realpath(path + '\\' + (str(int(regnumber))) + '.xls'))
    cont += 1
startfile(path)