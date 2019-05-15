# -*- coding: utf-8 -*-
"""
Created on Sun Jun 11 16:06:59 2017

@author: manna
"""

import os
# import glob
import xlrd
import xlsxwriter

        
# 根据索引获取Excel表格中的数据   
# 参数:file：Excel文件路径, colnameindex：表头列名所在行，by_index：表的索引
def excel_table_byindex(file='file.xls', colnameindex=0, by_index=0):
    data = xlrd.open_workbook(file)
    table = data.sheets()[by_index]
    nrows, ncols = table.nrows, table.ncols  # 行数, 列数
    colnames = table.row_values(colnameindex)  # 列名
    # print(colnames)
    
    data_list = []
    for rownum in range(colnameindex+1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            data_list.append(app)
    return colnames, data_list
    
    
# 根据名称获取Excel表格中的数据
# 参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
def excel_table_byname(file='file.xls', colnameindex=0, by_name=u'Sheet1'):
    data = xlrd.open_workbook(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 某一行数据
    print(colnames)
    list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    return list
    
    
def get_xls_data(file):
    return excel_table_byindex(file, 1)


def to_xls(file):
    workbook = xlsxwriter.Workbook(out_excel_file)
    worksheet = workbook.add_worksheet(u'已脱贫数据')
    # worksheet.set_column(0,len(colnames)+1,22)
    title_bold = workbook.add_format({'bold': True, 'border': 2, 'bg_color': 'blue'})
    border = workbook.add_format({'border': 1})

    for i, j in enumerate(colnames):
        worksheet.set_column(i, i, len(j) + 1)
        worksheet.write_string(0, i, j, title_bold)

    for i, row in enumerate(data_arry):
        for j in range(len(row)):
            cell_value = row.get(colnames[j])
            # print("colname={},i={}, j={}, cell_value={}".format(colnames[j],i,j,cell_value))
            worksheet.write_string(i + 1, j, cell_value, border)
        # i += 1
        # table.set_column(1,1,30)
        # table.set_column(0,0,16)
    # print(len(colnames))
    # worksheet.set_column(0,len(colnames)+1,30)

    workbook.close()
    
    
if __name__ == '__main__':
    out_excel_file = u'data\\test.xlsx'
    wdir = u"data"

    colnames = []
    data_arry = []
    
    # for file in glob.glob(input_excel_files):
    types = ("xls")
    for root, dirs, files in os.walk(wdir):
        print(files)
        for f in files:
            if os.path.isfile(os.path.join(root, f)) and os.path.splitext(f)[1][1:] in types:
                print(os.path.join(root, f))
                colnames, dict_content = get_xls_data(os.path.join(root, f))
                
                data_arry.extend(dict_content)
    
    dict_ing = colnames

