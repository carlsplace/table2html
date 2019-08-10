# -*- encoding=utf8 -*-

'''
FileName:   table2html.py
Author:     Chang YC
@contact:   chang.yc@outlook.com
@version:   $Id$
Description:
Changelog:
'''

import csv
import argparse

def read_csv(csv_path):
    cell_info = {"content": "word", "colspan": "1", "rowspan": "1"}
    cell_info_rows = []

    with open(csv_path) as f:
        f_csv = csv.reader(f)
        list_table = list(f_csv)

    for i in range(len(list_table)):
        row = list_table[i]
        cell_info_row = []
        for j in range(len(row)):
            cell = row[j]
            colspan = 1
            rowspan = 1
            if cell or (i ==0 and j == 0): # cell为第一个或有内容才处理
                # 往右遍历行，确认colspan
                for cell_right in row[j+1:]:
                    if cell_right :
                        break
                    else:
                        colspan += 1
                # 往下遍历列，确认rowspan
                for row_down in list_table[i+1:]:
                    cell_down = row_down[j]
                    if cell_down:
                        break
                    else:
                        rowspan += 1
                cell_info = {
                    "content": cell,
                    "colspan": str(colspan),
                    "rowspan": str(rowspan)
                }
                cell_info_row.append(cell_info)
        cell_info_rows.append(cell_info_row)
    return cell_info_rows

def html_table(context):
    left = '<table border="1">'
    right = "</table>"
    return left + context + right

def html_cell(cell_info):
    rowspan = cell_info.get("rowspan", "1")
    colspan = cell_info.get("colspan", "1")
    content = cell_info.get("content")
    if rowspan == "1" and colspan == "1":
        cell = "<td>" + content + "</td>"
    else:
        cell = '<th rowspan="%s" colspan="%s">' % (rowspan, colspan) \
             + content + "</th>"
    return cell

def html_row(context, color=None):
    if color:
        left = '<tr bgcolor="%s">' % color
    else:
        left = "<tr>"
    right = "</tr>"
    return left + context + right

def csv2html(csv_path, header_num=0, color="LightGray"):
    csv = read_csv(csv_path)
    row_list = []
    row_num = 0
    for row in csv:
        cell_list = []
        for cell_info in row:
            cell = html_cell(cell_info)
            cell_list.append(cell)
        row_context = '\n'.join(cell_list)
        if row_num < header_num: # 如果行数为表头行，更改表格底色
            row_list.append(html_row(row_context, color=color))
        else:
            row_list.append(html_row(row_context, color="white"))
        row_num += 1
    table_context = '\n'.join(row_list)
    return html_table(table_context)

def excel2csv(excel_path):
    import xlrd
    workbook = xlrd.open_workbook(excel_path)
    csv_path = excel_path + '.csv'
    table = workbook.sheet_by_index(0)
    with open(csv_path, 'w', encoding='utf-8') as f:
        write = csv.writer(f)
        for row_num in range(table.nrows):
            row_value = table.row_values(row_num)
            write.writerow(row_value)
    return csv_path

def table2html(file_path, header_num=0):
    if ".xls" in file_path:
        file_path = excel2csv(file_path)
    res = csv2html(file_path, header_num=header_num)
    with open('样表.html', 'w') as f:
        f.write(res)
    return res

if __name__ == "__main__":
    table2html("样表.xlsx", header_num=2)