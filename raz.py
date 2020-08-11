import xlrd
import xlwt
from tkinter.filedialog import askopenfilename


filename = askopenfilename()
rb = xlrd.open_workbook(filename)

# font
font = xlwt.Font()
font.name = 'Times New Roman'
font.height = 11
# borders
borders = xlwt.Borders()
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1

date_style = xlwt.XFStyle()
date_style.num_format_str = "M/D/YY"
date_style.borders = borders

style = xlwt.XFStyle()
style.borders = borders

sheetR = rb.sheet_by_index(0)
regNumber = set()
nrows = sheetR.nrows
ncols = sheetR.ncols

for row in range(3, nrows):
    if sheetR.cell_type(row, 9) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        regNumber.add(sheetR.cell_value(row, 9))

cont = 1
abc = len(regNumber)
for val in regNumber:
    name = str(int(val))
    print(cont,'из',abc)
    cont += 1
    cnt = 2
    outbook = xlwt.Workbook()
    outsheet = outbook.add_sheet(name)

    outsheet.col(0).width = 2 * 256
    outsheet.col(1).width = 16 * 256
    outsheet.col(2).width = 26 * 256
    outsheet.col(3).width = 27 * 256
    outsheet.col(4).width = 24 * 256
    outsheet.col(5).width = 11 * 256
    outsheet.col(6).width = 4 * 256
    outsheet.col(7).width = 85 * 256
    outsheet.col(8).width = 7 * 256
    outsheet.col(9).width = 22 * 256
    outsheet.col(0).width = 13 * 256

    for row in range(1,2):
        for col in range(ncols):
            outsheet.write(row, col, sheetR.cell_value(row, col), style=style)
    for row in range(3, nrows):
        if val == sheetR.cell_value(row, 9):
            for col in range(ncols):
                if col == 5:
                    outsheet.write(cnt, col, sheetR.cell_value(row, col), style=date_style)
                else:
                    outsheet.write(cnt, col, sheetR.cell_value(row, col), style=style)
            cnt += 1
    outbook.save('D:\\1\\out\\' + name + '.xls')
