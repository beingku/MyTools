# -*- coding: utf-8 -*-
"""
用于提取excel文件信息
"""

import xlrd
import xlwt

L = 4


def open_excel(file='file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))


# 参数:file：Excel文件路径     colnameindex：表头列名所在行 by_index：表的索引
def excel_table_byindex(file='file.xls', by_index=0, l=3):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数

    row_list = []
    for rownum in range(nrows):
        row = table.row_values(rownum)
        if row:
            if ncols % l == 0:
                for col in range(ncols + 1):
                    if col != 0 and col % l == 0:
                        app = row[col - l:col]
                        row_list.append(app)
            else:
                for col in range((ncols // l + 1) * l + 1):
                    if (col != 0 and col % l == 0):
                        app = row[col - l:col]
                        row_list.append(app)

    return row_list


def write_to_excel(filename, sheetname):
    wb = xlwt.Workbook()
    sheet = wb.add_sheet(sheetname)  # sheet的名称为test

    # 单元格的格式
    # style = 'pattern: pattern solid, fore_colour yellow; '  # 背景颜色为黄色
    # style += 'font: bold on; '  # 粗体字
    # style += 'align: horz centre, vert center; '  # 居中
    # header_style = xlwt.easyxf(style)
    datas = excel_table_byindex("d:/test.xlsx", 0, L)
    for col in range(len(datas)):
        for row in range(len(datas[col])):
            sheet.write(row, col, datas[col][row])
    wb.save(filename)


if __name__ == "__main__":
    # l = excel_table_byindex("d:/test.xlsx",0,3)
    # print(l)
    # for n in l:
    #     print(len(n))
    write_to_excel("d:/test1.xlsx", "Sheet1")
